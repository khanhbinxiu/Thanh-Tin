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
      </div>

      <div className="bg-white rounded-xl border overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Ngày', 'Mã KH', 'Địa điểm', 'B45↓', 'B45↑', 'B12↓', 'B12↑', 'Gas giao', 'Gas TT', 'Đơn giá', 'Thành tiền', 'Ghi chú'].map(h => (
                <th key={h} className="text-left px-3 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {loading && <tr><td colSpan={12} className="px-4 py-8 text-center text-gray-400">Đang tải...</td></tr>}
            {!loading && rows.map(r => (
              <tr key={r.id} className="border-b last:border-0 hover:bg-gray-50">
                <td className="px-3 py-2 whitespace-nowrap">{fmtDate(r.delivery_date)}</td>
                <td className="px-3 py-2 font-medium">{r.customer_code}</td>
                <td className="px-3 py-2">{r.location}</td>
                <td className="px-3 py-2 text-right">{r.b45_delivered || ''}</td>
                <td className="px-3 py-2 text-right">{r.b45_returned || ''}</td>
                <td className="px-3 py-2 text-right">{r.b12_delivered || ''}</td>
                <td className="px-3 py-2 text-right">{r.b12_returned || ''}</td>
                <td className="px-3 py-2 text-right">{r.gas_delivered || ''}</td>
                <td className="px-3 py-2 text-right">{r.gas_paid || ''}</td>
                <td className="px-3 py-2 text-right">{r.unit_price ? r.unit_price.toLocaleString('vi-VN') : ''}</td>
                <td className="px-3 py-2 text-right">{r.total_amount ? r.total_amount.toLocaleString('vi-VN') : ''}</td>
                <td className="px-3 py-2 text-gray-500">{r.note}</td>
              </tr>
            ))}
            {!loading && rows.length === 0 && (
              <tr><td colSpan={12} className="px-4 py-8 text-center text-gray-400">Không có dữ liệu</td></tr>
            )}
          </tbody>
        </table>
        {rows.length === 500 && <p className="text-xs text-center text-gray-400 py-2">Hiển thị tối đa 500 dòng</p>}
      </div>
    </div>
  )
}
