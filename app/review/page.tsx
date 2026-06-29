'use client'
import { useEffect, useState } from 'react'
import { supabase, Transaction, Customer } from '@/lib/supabase'
import * as XLSX from 'xlsx'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }

type EditRow = {
  delivery_date: string
  customer_code: string
  location: string
  b45_delivered: string
  b45_returned: string
  b12_delivered: string
  b12_returned: string
  gas_returned: string
  unit_price: string
  note: string
}

function toEdit(r: Transaction): EditRow {
  return {
    delivery_date: r.delivery_date,
    customer_code: r.customer_code,
    location: r.location,
    b45_delivered: String(r.b45_delivered || 0),
    b45_returned: String(r.b45_returned || 0),
    b12_delivered: String(r.b12_delivered || 0),
    b12_returned: String(r.b12_returned || 0),
    gas_returned: String(r.gas_returned || 0),
    unit_price: String(r.unit_price || 0),
    note: r.note || '',
  }
}

export default function ReviewPage() {
  const [rows, setRows] = useState<Transaction[]>([])
  const [customers, setCustomers] = useState<Customer[]>([])
  const [filterCode, setFilterCode] = useState('')
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo, setDateTo] = useState('')
  const [loading, setLoading] = useState(false)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [edits, setEdits] = useState<Record<string, EditRow>>({})
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')

  useEffect(() => {
    supabase.from('customers').select('*').order('code').then(({ data }) => setCustomers(data || []))
    load()
  }, [])

  async function load() {
    setLoading(true); setMsg('')
    let q = supabase.from('transactions').select('*').order('delivery_date', { ascending: false })
    if (filterCode) q = q.eq('customer_code', filterCode)
    if (dateFrom) q = q.gte('delivery_date', dateFrom)
    if (dateTo) q = q.lte('delivery_date', dateTo)
    const { data } = await q.limit(500)
    setRows(data || [])
    setEdits({}); setSelected(new Set())
    setLoading(false)
  }

  function toggleSelect(id: string) {
    const ns = new Set(selected)
    if (ns.has(id)) {
      ns.delete(id)
      const ne = { ...edits }; delete ne[id]; setEdits(ne)
    } else {
      ns.add(id)
      const r = rows.find(r => r.id === id)
      if (r) setEdits(prev => ({ ...prev, [id]: toEdit(r) }))
    }
    setSelected(ns)
  }

  function toggleAll() {
    if (selected.size === rows.length) {
      setSelected(new Set()); setEdits({})
    } else {
      const ns = new Set(rows.map(r => r.id))
      const ne: Record<string, EditRow> = {}
      for (const r of rows) ne[r.id] = toEdit(r)
      setSelected(ns); setEdits(ne)
    }
  }

  function updateEdit(id: string, field: keyof EditRow, value: string) {
    setEdits(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } }))
  }

  async function saveEdits() {
    setSaving(true); setMsg('')
    let count = 0
    for (const id of Object.keys(edits)) {
      const e = edits[id]
      const b45d = parseFloat(e.b45_delivered) || 0
      const b45r = parseFloat(e.b45_returned) || 0
      const b12d = parseFloat(e.b12_delivered) || 0
      const b12r = parseFloat(e.b12_returned) || 0
      const gasRet = parseFloat(e.gas_returned) || 0
      const unitPrice = parseFloat(e.unit_price) || 0
      const gasDelivered = b45d * 45 + b12d * 12
      const gasPaid = gasDelivered - gasRet
      const totalAmount = gasPaid * unitPrice
      await supabase.from('transactions').update({
        delivery_date: e.delivery_date,
        b45_delivered: b45d, b45_returned: b45r,
        b12_delivered: b12d, b12_returned: b12r,
        gas_delivered: gasDelivered, gas_returned: gasRet, gas_paid: gasPaid,
        unit_price: unitPrice, total_amount: totalAmount,
        note: e.note,
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
      await supabase.from('transactions').delete().in('id', ids.slice(i, i + 50))
    }
    setMsg(`Đã xóa ${ids.length} dòng`)
    setSaving(false)
    await load()
  }

  function exportExcel() {
    const ws = XLSX.utils.json_to_sheet(rows.map(r => ({
      'Ngày giao': fmtDate(r.delivery_date), 'Mã KH': r.customer_code, 'Địa điểm': r.location,
      'B45 Giao': r.b45_delivered, 'B45 Trả': r.b45_returned, 'B12 Giao': r.b12_delivered, 'B12 Trả': r.b12_returned,
      'Gas giao': r.gas_delivered, 'Gas trả': r.gas_returned, 'Gas TT': r.gas_paid,
      'Đơn giá': r.unit_price, 'Thành tiền': r.total_amount, 'Ghi chú': r.note,
    })))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Giao dịch')
    XLSX.writeFile(wb, `ThanhTin_DuLieu.xlsx`)
  }

  const editCount = Object.keys(edits).length

  return (
    <div className="space-y-4">
      <h1 className="text-lg md:text-xl font-bold text-gray-800">Xem & sửa dữ liệu</h1>

      <div className="bg-white rounded-xl border p-3 flex flex-wrap gap-3 items-end">
        <div>
          <label className="block text-xs text-gray-500 mb-1">Khách hàng</label>
          <select value={filterCode} onChange={e => setFilterCode(e.target.value)} className="border rounded px-3 py-2 text-sm">
            <option value="">Tất cả</option>
            {customers.map(c => <option key={c.code} value={c.code}>{c.code} – {c.name}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Từ ngày</label>
          <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} className="border rounded px-3 py-2 text-sm" />
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Đến ngày</label>
          <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} className="border rounded px-3 py-2 text-sm" />
        </div>
        <button onClick={load} className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700">Lọc</button>
        <button onClick={exportExcel} className="border px-4 py-2 rounded text-sm hover:bg-gray-50">Xuất Excel</button>
      </div>

      <p className="text-xs text-gray-400">Tick ô checkbox để chọn dòng cần sửa hoặc xóa. Sửa trực tiếp trên bảng rồi bấm Lưu.</p>

      {(editCount > 0 || selected.size > 0) && (
        <div className="bg-blue-50 border border-blue-200 rounded-lg px-4 py-3 flex flex-wrap items-center gap-3">
          {editCount > 0 && (
            <button onClick={saveEdits} disabled={saving} className="bg-green-600 text-white px-4 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
              {saving ? 'Đang lưu...' : `Lưu ${editCount} dòng`}
            </button>
          )}
          {selected.size > 0 && (
            <button onClick={deleteSelected} disabled={saving} className="bg-red-600 text-white px-4 py-2 rounded text-sm hover:bg-red-700 disabled:opacity-50">
              Xóa {selected.size} dòng
            </button>
          )}
        </div>
      )}

      {msg && <div className="bg-green-50 border border-green-200 rounded-lg px-3 py-2.5 text-sm text-green-800">{msg}</div>}

      <div className="bg-white rounded-xl border overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b sticky top-0">
            <tr>
              <th className="px-2 py-2"><input type="checkbox" checked={selected.size === rows.length && rows.length > 0} onChange={toggleAll} /></th>
              {['Ngày','Mã KH','Địa điểm','B45↓','B45↑','B12↓','B12↑','Gas giao','Gas trả','Gas TT','Đơn giá','Thành tiền','Ghi chú'].map(h => (
                <th key={h} className="text-left px-2 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading && <tr><td colSpan={14} className="px-4 py-8 text-center text-gray-400">Đang tải...</td></tr>}
            {!loading && rows.map(r => {
              const isEditing = selected.has(r.id)
              const e = edits[r.id]
              const b45d = isEditing ? parseFloat(e.b45_delivered) || 0 : r.b45_delivered
              const b12d = isEditing ? parseFloat(e.b12_delivered) || 0 : r.b12_delivered
              const gasRet = isEditing ? parseFloat(e.gas_returned) || 0 : r.gas_returned
              const uPrice = isEditing ? parseFloat(e.unit_price) || 0 : r.unit_price
              const gasDelivered = b45d * 45 + b12d * 12
              const gasPaid = gasDelivered - gasRet
              const total = gasPaid * uPrice
              const inputCls = 'border border-blue-300 bg-blue-50 rounded px-1.5 py-1 text-xs text-right w-full'

              return (
                <tr key={r.id} className={`border-b ${isEditing ? 'bg-blue-50' : 'hover:bg-gray-50'}`}>
                  <td className="px-2 py-1.5 text-center">
                    <input type="checkbox" checked={isEditing} onChange={() => toggleSelect(r.id)} />
                  </td>
                  <td className="px-2 py-1.5 whitespace-nowrap">
                    {isEditing ? <input type="date" value={e.delivery_date} onChange={ev => updateEdit(r.id, 'delivery_date', ev.target.value)} className="border border-blue-300 bg-blue-50 rounded px-1 py-1 text-xs w-28" /> : fmtDate(r.delivery_date)}
                  </td>
                  <td className="px-2 py-1.5 font-medium">{r.customer_code}</td>
                  <td className="px-2 py-1.5">{r.location}</td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b45_delivered} onChange={ev => updateEdit(r.id, 'b45_delivered', ev.target.value)} className={inputCls} /> : <span className="px-2">{r.b45_delivered || ''}</span>}
                  </td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b45_returned} onChange={ev => updateEdit(r.id, 'b45_returned', ev.target.value)} className={inputCls} /> : <span className="px-2">{r.b45_returned || ''}</span>}
                  </td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b12_delivered} onChange={ev => updateEdit(r.id, 'b12_delivered', ev.target.value)} className={inputCls} /> : <span className="px-2">{r.b12_delivered || ''}</span>}
                  </td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b12_returned} onChange={ev => updateEdit(r.id, 'b12_returned', ev.target.value)} className={inputCls} /> : <span className="px-2">{r.b12_returned || ''}</span>}
                  </td>
                  <td className="px-2 py-1.5 text-right text-gray-500">{gasDelivered || ''}</td>
                  <td className="px-1 py-1" style={{minWidth:60}}>
                    {isEditing ? <input type="number" step="0.1" value={e.gas_returned} onChange={ev => updateEdit(r.id, 'gas_returned', ev.target.value)} className={inputCls} /> : <span className="px-2">{r.gas_returned || ''}</span>}
                  </td>
                  <td className="px-2 py-1.5 text-right text-blue-600 font-medium">{gasPaid ? gasPaid.toFixed(1) : ''}</td>
                  <td className="px-1 py-1" style={{minWidth:70}}>
                    {isEditing ? <input type="number" value={e.unit_price} onChange={ev => updateEdit(r.id, 'unit_price', ev.target.value)} className={inputCls} /> : <span className="px-2">{r.unit_price ? r.unit_price.toLocaleString('vi-VN') : ''}</span>}
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium">{total ? total.toLocaleString('vi-VN') : ''}</td>
                  <td className="px-1 py-1" style={{minWidth:60}}>
                    {isEditing ? <input value={e.note} onChange={ev => updateEdit(r.id, 'note', ev.target.value)} className="border border-blue-300 bg-blue-50 rounded px-1.5 py-1 text-xs w-full" /> : <span className="px-2 text-gray-500">{r.note}</span>}
                  </td>
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
