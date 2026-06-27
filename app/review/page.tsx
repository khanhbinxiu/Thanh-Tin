'use client'
import { useEffect, useState } from 'react'
import { supabase, Transaction, Customer } from '@/lib/supabase'
import * as XLSX from 'xlsx'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }

type Edit = { gas_returned: string; unit_price: string }

export default function ReviewPage() {
  const [rows, setRows] = useState<Transaction[]>([])
  const [customers, setCustomers] = useState<Customer[]>([])
  const [filterCode, setFilterCode] = useState('')
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo, setDateTo] = useState('')
  const [loading, setLoading] = useState(false)
  const [edits, setEdits] = useState<Record<string, Edit>>({})
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')

  useEffect(() => {
    supabase.from('customers').select('*').order('code').then(({ data }) => setCustomers(data || []))
    load()
  }, [])

  async function load() {
    setLoading(true)
    setMsg('')
    let q = supabase.from('transactions').select('*').order('delivery_date', { ascending: false })
    if (filterCode) q = q.eq('customer_code', filterCode)
    if (dateFrom) q = q.gte('delivery_date', dateFrom)
    if (dateTo) q = q.lte('delivery_date', dateTo)
    const { data } = await q.limit(500)
    setRows(data || [])
    setEdits({})
    setSelected(new Set())
    setLoading(false)
  }

  function startEdit(r: Transaction) {
    if (edits[r.id]) return
    setEdits(prev => ({
      ...prev,
      [r.id]: { gas_returned: String(r.gas_returned || 0), unit_price: String(r.unit_price || 0) }
    }))
  }

  function updateEdit(id: string, field: keyof Edit, value: string) {
    setEdits(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } }))
  }

  function cancelEdit(id: string) {
    setEdits(prev => { const n = { ...prev }; delete n[id]; return n })
  }

  function toggleSelect(id: string) {
    setSelected(prev => {
      const n = new Set(prev)
      n.has(id) ? n.delete(id) : n.add(id)
      return n
    })
  }

  function toggleAll() {
    if (selected.size === rows.length) {
      setSelected(new Set())
    } else {
      setSelected(new Set(rows.map(r => r.id)))
    }
  }

  async function saveEdits() {
    const ids = Object.keys(edits)
    if (ids.length === 0) return
    setSaving(true)
    let count = 0
    for (const id of ids) {
      const e = edits[id]
      const r = rows.find(r => r.id === id)
      if (!r) continue
      const gasReturned = parseFloat(e.gas_returned) || 0
      const unitPrice = parseFloat(e.unit_price) || 0
      const gasPaid = r.b45_delivered * 45 + r.b12_delivered * 12 - gasReturned
      const totalAmount = gasPaid * unitPrice
      await supabase.from('transactions').update({
        gas_returned: gasReturned, gas_paid: gasPaid,
        unit_price: unitPrice, total_amount: totalAmount,
      }).eq('id', id)
      count++
    }
    setMsg(`Đã lưu ${count} dòng`)
    setSaving(false)
    await load()
  }

  async function deleteSelected() {
    if (selected.size === 0) return
    if (!confirm(`Xóa ${selected.size} dòng đã chọn?`)) return
    setSaving(true)
    const ids = [...selected]
    for (let i = 0; i < ids.length; i += 50) {
      const batch = ids.slice(i, i + 50)
      await supabase.from('transactions').delete().in('id', batch)
    }
    setMsg(`Đã xóa ${ids.length} dòng`)
    setSaving(false)
    await load()
  }

  function exportExcel() {
    const ws = XLSX.utils.json_to_sheet(rows.map(r => ({
      'Ngày giao': fmtDate(r.delivery_date),
      'Mã KH': r.customer_code,
      'Địa điểm': r.location,
      'PIC': r.pic,
      'B45 Giao': r.b45_delivered,
      'B45 Trả': r.b45_returned,
      'B12 Giao': r.b12_delivered,
      'B12 Trả': r.b12_returned,
      'Gas giao (kg)': r.gas_delivered,
      'Gas trả (kg)': r.gas_returned,
      'Gas TT (kg)': r.gas_paid,
      'Đơn giá': r.unit_price,
      'Thành tiền': r.total_amount,
      'Ghi chú': r.note,
    })))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Giao dịch')
    XLSX.writeFile(wb, `ThanhTin_DuLieu_${dateFrom || 'all'}_${dateTo || 'all'}.xlsx`)
  }

  const editCount = Object.keys(edits).length

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Xem & xuất dữ liệu</h1>

      <div className="bg-white rounded-xl border p-4 flex flex-wrap gap-3 items-end">
        <div>
          <label className="block text-xs text-gray-500 mb-1">Khách hàng</label>
          <select value={filterCode} onChange={e => setFilterCode(e.target.value)}
            className="border rounded px-3 py-2 text-sm">
            <option value="">Tất cả</option>
            {customers.map(c => <option key={c.code} value={c.code}>{c.code} – {c.name}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Từ ngày</label>
          <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)}
            className="border rounded px-3 py-2 text-sm" />
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Đến ngày</label>
          <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)}
            className="border rounded px-3 py-2 text-sm" />
        </div>
        <button onClick={load} className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700">
          Lọc
        </button>
        <button onClick={exportExcel} className="border px-4 py-2 rounded text-sm hover:bg-gray-50">
          Xuất Excel
        </button>
      </div>

      {(editCount > 0 || selected.size > 0) && (
        <div className="bg-blue-50 border border-blue-200 rounded-lg px-4 py-3 flex items-center gap-3">
          {editCount > 0 && (
            <button onClick={saveEdits} disabled={saving}
              className="bg-green-600 text-white px-4 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
              {saving ? 'Đang lưu...' : `Lưu ${editCount} thay đổi`}
            </button>
          )}
          {selected.size > 0 && (
            <button onClick={deleteSelected} disabled={saving}
              className="bg-red-600 text-white px-4 py-2 rounded text-sm hover:bg-red-700 disabled:opacity-50">
              Xóa {selected.size} dòng
            </button>
          )}
          <span className="text-xs text-gray-500">Bấm vào Gas trả / Đơn giá để sửa · Tab chuyển ô</span>
        </div>
      )}

      {msg && <div className="bg-green-50 border border-green-200 rounded-lg px-4 py-3 text-sm text-green-800">{msg}</div>}

      <div className="bg-white rounded-xl border overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b">
            <tr>
              <th className="px-2 py-2"><input type="checkbox" checked={selected.size === rows.length && rows.length > 0} onChange={toggleAll} /></th>
              {['Ngày', 'Mã KH', 'Địa điểm', 'B45↓', 'B45↑', 'B12↓', 'B12↑', 'Gas giao', 'Gas trả', 'Gas TT', 'Đơn giá', 'Thành tiền', 'Ghi chú'].map(h => (
                <th key={h} className="text-left px-3 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading && <tr><td colSpan={14} className="px-4 py-8 text-center text-gray-400">Đang tải...</td></tr>}
            {!loading && rows.map(r => {
              const e = edits[r.id]
              const isEditing = !!e
              const gasRet = isEditing ? parseFloat(e.gas_returned) || 0 : r.gas_returned
              const uPrice = isEditing ? parseFloat(e.unit_price) || 0 : r.unit_price
              const gasPaid = r.b45_delivered * 45 + r.b12_delivered * 12 - gasRet
              const total = gasPaid * uPrice
              return (
                <tr key={r.id} className={`border-b last:border-0 ${selected.has(r.id) ? 'bg-red-50' : isEditing ? 'bg-blue-50' : 'hover:bg-gray-50'}`}>
                  <td className="px-2 py-2 text-center">
                    <input type="checkbox" checked={selected.has(r.id)} onChange={() => toggleSelect(r.id)} />
                  </td>
                  <td className="px-3 py-2 whitespace-nowrap">{fmtDate(r.delivery_date)}</td>
                  <td className="px-3 py-2 font-medium">{r.customer_code}</td>
                  <td className="px-3 py-2">{r.location}</td>
                  <td className="px-3 py-2 text-right">{r.b45_delivered || ''}</td>
                  <td className="px-3 py-2 text-right">{r.b45_returned || ''}</td>
                  <td className="px-3 py-2 text-right">{r.b12_delivered || ''}</td>
                  <td className="px-3 py-2 text-right">{r.b12_returned || ''}</td>
                  <td className="px-3 py-2 text-right">{r.gas_delivered || ''}</td>
                  <td className="px-1 py-1">
                    <input type="number" step="0.1"
                      value={isEditing ? e.gas_returned : (r.gas_returned || '')}
                      onFocus={() => startEdit(r)}
                      onChange={ev => updateEdit(r.id, 'gas_returned', ev.target.value)}
                      className={`border rounded px-2 py-1 text-sm w-16 text-right ${isEditing ? 'border-blue-400 bg-blue-50' : 'border-gray-200'}`} />
                  </td>
                  <td className="px-3 py-2 text-right text-blue-600 font-medium">{gasPaid ? gasPaid.toFixed(1) : ''}</td>
                  <td className="px-1 py-1">
                    <input type="number"
                      value={isEditing ? e.unit_price : (r.unit_price || '')}
                      onFocus={() => startEdit(r)}
                      onChange={ev => updateEdit(r.id, 'unit_price', ev.target.value)}
                      className={`border rounded px-2 py-1 text-sm w-20 text-right ${isEditing ? 'border-blue-400 bg-blue-50' : 'border-gray-200'}`} />
                  </td>
                  <td className="px-3 py-2 text-right font-medium">{total ? total.toLocaleString('vi-VN') : ''}</td>
                  <td className="px-3 py-2 text-gray-500">{r.note}</td>
                </tr>
              )
            })}
            {!loading && rows.length === 0 && (
              <tr><td colSpan={14} className="px-4 py-8 text-center text-gray-400">Không có dữ liệu</td></tr>
            )}
          </tbody>
        </table>
        {rows.length === 500 && <p className="text-xs text-center text-gray-400 py-2">Hiển thị tối đa 500 dòng</p>}
      </div>
    </div>
  )
}
