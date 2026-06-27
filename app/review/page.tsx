'use client'
import { useEffect, useState } from 'react'
import { supabase, Transaction, Customer } from '@/lib/supabase'
import * as XLSX from 'xlsx'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }

export default function ReviewPage() {
  const [rows, setRows] = useState<Transaction[]>([])
  const [customers, setCustomers] = useState<Customer[]>([])
  const [filterCode, setFilterCode] = useState('')
  const [dateFrom, setDateFrom] = useState('')
  const [dateTo, setDateTo] = useState('')
  const [loading, setLoading] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editGasReturned, setEditGasReturned] = useState('')
  const [editUnitPrice, setEditUnitPrice] = useState('')
  const [saving, setSaving] = useState(false)

  useEffect(() => {
    supabase.from('customers').select('*').order('code').then(({ data }) => setCustomers(data || []))
    load()
  }, [])

  async function load() {
    setLoading(true)
    let q = supabase.from('transactions').select('*').order('delivery_date', { ascending: false })
    if (filterCode) q = q.eq('customer_code', filterCode)
    if (dateFrom) q = q.gte('delivery_date', dateFrom)
    if (dateTo) q = q.lte('delivery_date', dateTo)
    const { data } = await q.limit(500)
    setRows(data || [])
    setLoading(false)
  }

  function startEdit(r: Transaction) {
    setEditingId(r.id)
    setEditGasReturned(String(r.gas_returned || 0))
    setEditUnitPrice(String(r.unit_price || 0))
  }

  async function saveEdit(r: Transaction) {
    setSaving(true)
    const gasReturned = parseFloat(editGasReturned) || 0
    const unitPrice = parseFloat(editUnitPrice) || 0
    const gasPaid = r.b45_delivered * 45 + r.b12_delivered * 12 - gasReturned
    const totalAmount = gasPaid * unitPrice

    await supabase.from('transactions').update({
      gas_returned: gasReturned,
      gas_paid: gasPaid,
      unit_price: unitPrice,
      total_amount: totalAmount,
    }).eq('id', r.id)

    setEditingId(null)
    setSaving(false)
    await load()
  }

  function handleEditKeyDown(e: React.KeyboardEvent, r: Transaction) {
    if (e.key === 'Enter') { e.preventDefault(); saveEdit(r) }
    if (e.key === 'Escape') setEditingId(null)
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
        <span className="text-xs text-gray-400 ml-2">Bấm vào dòng để sửa Gas trả / Đơn giá · Enter lưu · Esc hủy</span>
      </div>

      <div className="bg-white rounded-xl border overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Ngày', 'Mã KH', 'Địa điểm', 'B45↓', 'B45↑', 'B12↓', 'B12↑', 'Gas giao', 'Gas trả', 'Gas TT', 'Đơn giá', 'Thành tiền', 'Ghi chú'].map(h => (
                <th key={h} className="text-left px-3 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading && <tr><td colSpan={13} className="px-4 py-8 text-center text-gray-400">Đang tải...</td></tr>}
            {!loading && rows.map(r => {
              const isEditing = editingId === r.id
              return (
                <tr key={r.id}
                  className={`border-b last:border-0 cursor-pointer ${isEditing ? 'bg-blue-50' : 'hover:bg-gray-50'}`}
                  onClick={() => !isEditing && startEdit(r)}>
                  <td className="px-3 py-2 whitespace-nowrap">{fmtDate(r.delivery_date)}</td>
                  <td className="px-3 py-2 font-medium">{r.customer_code}</td>
                  <td className="px-3 py-2">{r.location}</td>
                  <td className="px-3 py-2 text-right">{r.b45_delivered || ''}</td>
                  <td className="px-3 py-2 text-right">{r.b45_returned || ''}</td>
                  <td className="px-3 py-2 text-right">{r.b12_delivered || ''}</td>
                  <td className="px-3 py-2 text-right">{r.b12_returned || ''}</td>
                  <td className="px-3 py-2 text-right">{r.gas_delivered || ''}</td>
                  {isEditing ? (
                    <>
                      <td className="px-1 py-1" onClick={e => e.stopPropagation()}>
                        <input type="number" step="0.1" value={editGasReturned}
                          onChange={e => setEditGasReturned(e.target.value)}
                          onKeyDown={e => handleEditKeyDown(e, r)}
                          autoFocus
                          className="border-2 border-blue-400 rounded px-2 py-1 text-sm w-16 text-right" />
                      </td>
                      <td className="px-3 py-2 text-right text-blue-600 font-medium">
                        {(r.b45_delivered * 45 + r.b12_delivered * 12 - (parseFloat(editGasReturned) || 0)).toFixed(1)}
                      </td>
                      <td className="px-1 py-1" onClick={e => e.stopPropagation()}>
                        <input type="number" value={editUnitPrice}
                          onChange={e => setEditUnitPrice(e.target.value)}
                          onKeyDown={e => handleEditKeyDown(e, r)}
                          className="border-2 border-blue-400 rounded px-2 py-1 text-sm w-20 text-right" />
                      </td>
                      <td className="px-3 py-2 text-right font-medium">
                        {((r.b45_delivered * 45 + r.b12_delivered * 12 - (parseFloat(editGasReturned) || 0)) * (parseFloat(editUnitPrice) || 0)).toLocaleString('vi-VN')}
                      </td>
                    </>
                  ) : (
                    <>
                      <td className="px-3 py-2 text-right">{r.gas_returned || ''}</td>
                      <td className="px-3 py-2 text-right">{r.gas_paid || ''}</td>
                      <td className="px-3 py-2 text-right">{r.unit_price ? r.unit_price.toLocaleString('vi-VN') : ''}</td>
                      <td className="px-3 py-2 text-right">{r.total_amount ? r.total_amount.toLocaleString('vi-VN') : ''}</td>
                    </>
                  )}
                  <td className="px-3 py-2 text-gray-500">{r.note}</td>
                </tr>
              )
            })}
            {!loading && rows.length === 0 && (
              <tr><td colSpan={13} className="px-4 py-8 text-center text-gray-400">Không có dữ liệu</td></tr>
            )}
          </tbody>
        </table>
        {rows.length === 500 && <p className="text-xs text-center text-gray-400 py-2">Hiển thị tối đa 500 dòng</p>}
      </div>
    </div>
  )
}
