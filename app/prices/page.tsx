'use client'
import { useEffect, useState } from 'react'
import { supabase } from '@/lib/supabase'

type PriceRow = {
  customer_code: string
  customer_name: string
  location: string
  output_file_name: string
  unit_price: number
  count: number
}

export default function PricesPage() {
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [rows, setRows] = useState<PriceRow[]>([])
  const [loading, setLoading] = useState(false)
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')
  const [editPrices, setEditPrices] = useState<Record<string, string>>({})

  async function load() {
    setLoading(true)
    setMsg('')
    const from = `${year}-${String(month).padStart(2, '0')}-01`
    const lastDay = new Date(year, month, 0).getDate()
    const to = `${year}-${String(month).padStart(2, '0')}-${lastDay}`

    const { data: txs } = await supabase.from('transactions')
      .select('customer_code, location, output_file_name, unit_price')
      .gte('delivery_date', from)
      .lte('delivery_date', to)
      .order('customer_code')

    const { data: custs } = await supabase.from('customers').select('code, name')
    const custMap = new Map((custs || []).map(c => [c.code, c.name]))

    // Group by customer_code + location
    const grouped = new Map<string, PriceRow>()
    for (const tx of (txs || [])) {
      const key = `${tx.customer_code}||${tx.location}`
      if (!grouped.has(key)) {
        grouped.set(key, {
          customer_code: tx.customer_code,
          customer_name: custMap.get(tx.customer_code) || tx.customer_code,
          location: tx.location,
          output_file_name: tx.output_file_name || '',
          unit_price: tx.unit_price || 0,
          count: 0,
        })
      }
      const row = grouped.get(key)!
      row.count++
      if (tx.unit_price && !row.unit_price) row.unit_price = tx.unit_price
    }

    const result = [...grouped.values()].sort((a, b) =>
      a.output_file_name.localeCompare(b.output_file_name) || a.location.localeCompare(b.location)
    )
    setRows(result)

    const prices: Record<string, string> = {}
    for (const r of result) {
      const key = `${r.customer_code}||${r.location}`
      prices[key] = r.unit_price ? String(r.unit_price) : ''
    }
    setEditPrices(prices)
    setLoading(false)
  }

  async function saveAll() {
    setSaving(true)
    setMsg('')
    const from = `${year}-${String(month).padStart(2, '0')}-01`
    const lastDay = new Date(year, month, 0).getDate()
    const to = `${year}-${String(month).padStart(2, '0')}-${lastDay}`

    let updated = 0
    for (const row of rows) {
      const key = `${row.customer_code}||${row.location}`
      const newPrice = parseFloat(editPrices[key] || '0')
      if (newPrice !== row.unit_price) {
        await supabase.from('transactions')
          .update({ unit_price: newPrice })
          .eq('customer_code', row.customer_code)
          .eq('location', row.location)
          .gte('delivery_date', from)
          .lte('delivery_date', to)
        updated++
      }
    }
    setMsg(`Đã cập nhật ${updated} nhóm giá!`)
    await load()
    setSaving(false)
  }

  function applyToFile(fileName: string, price: string) {
    const newPrices = { ...editPrices }
    for (const row of rows) {
      if (row.output_file_name === fileName) {
        newPrices[`${row.customer_code}||${row.location}`] = price
      }
    }
    setEditPrices(newPrices)
  }

  useEffect(() => { load() }, [])

  const noPrice = rows.filter(r => !r.unit_price)
  const fileGroups = [...new Set(rows.map(r => r.output_file_name))].sort()

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Nhập giá theo tháng</h1>

      <div className="bg-white rounded-xl border p-4 flex gap-4 items-end">
        <div>
          <label className="block text-xs text-gray-500 mb-1">Tháng</label>
          <select value={month} onChange={e => setMonth(Number(e.target.value))}
            className="border rounded px-3 py-2 text-sm">
            {Array.from({ length: 12 }, (_, i) => i + 1).map(m =>
              <option key={m} value={m}>Tháng {m}</option>
            )}
          </select>
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Năm</label>
          <input type="number" value={year} onChange={e => setYear(Number(e.target.value))}
            className="border rounded px-3 py-2 text-sm w-24" />
        </div>
        <button onClick={load} disabled={loading}
          className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
          {loading ? 'Đang tải...' : 'Tải dữ liệu'}
        </button>
        {rows.length > 0 && (
          <button onClick={saveAll} disabled={saving}
            className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
            {saving ? 'Đang lưu...' : 'Lưu tất cả'}
          </button>
        )}
      </div>

      {msg && <div className="bg-green-50 border border-green-200 rounded-lg px-4 py-3 text-sm text-green-800">{msg}</div>}

      {noPrice.length > 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3 text-sm text-amber-800">
          {noPrice.length} địa điểm chưa có giá
        </div>
      )}

      {rows.length > 0 && (
        <div className="space-y-4">
          {fileGroups.map(fname => {
            const fileRows = rows.filter(r => r.output_file_name === fname)
            const allSamePrice = fileRows.every(r => {
              const key = `${r.customer_code}||${r.location}`
              return editPrices[key] === editPrices[`${fileRows[0].customer_code}||${fileRows[0].location}`]
            })
            return (
              <div key={fname} className="bg-white rounded-xl border overflow-hidden">
                <div className="bg-gray-50 border-b px-4 py-3 flex items-center justify-between">
                  <div>
                    <span className="font-semibold text-gray-800">{fname || 'UNMAPPED'}</span>
                    <span className="ml-2 text-xs text-gray-400">{fileRows.length} địa điểm · {fileRows.reduce((s, r) => s + r.count, 0)} dòng</span>
                  </div>
                  {fileRows.length > 1 && (
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-gray-500">Áp dụng cùng giá:</span>
                      <input
                        type="number"
                        placeholder="Giá chung"
                        className="border rounded px-2 py-1 text-sm w-28"
                        onKeyDown={e => {
                          if (e.key === 'Enter') applyToFile(fname, (e.target as HTMLInputElement).value)
                        }}
                        onBlur={e => { if (e.target.value) applyToFile(fname, e.target.value) }}
                      />
                    </div>
                  )}
                </div>
                <table className="w-full text-sm">
                  <thead className="bg-gray-50 border-b">
                    <tr>
                      <th className="text-left px-4 py-2 font-medium text-gray-600">Địa điểm</th>
                      <th className="text-left px-4 py-2 font-medium text-gray-600">Khách hàng</th>
                      <th className="text-center px-4 py-2 font-medium text-gray-600">Số dòng</th>
                      <th className="text-right px-4 py-2 font-medium text-gray-600">Đơn giá (vnđ/kg)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {fileRows.map(r => {
                      const key = `${r.customer_code}||${r.location}`
                      const hasPrice = editPrices[key] && parseFloat(editPrices[key]) > 0
                      return (
                        <tr key={key} className={`border-b last:border-0 ${!hasPrice ? 'bg-amber-50' : 'hover:bg-gray-50'}`}>
                          <td className="px-4 py-2.5 font-medium">{r.location}</td>
                          <td className="px-4 py-2.5 text-xs text-gray-500">{r.customer_name}</td>
                          <td className="px-4 py-2.5 text-center text-gray-500">{r.count}</td>
                          <td className="px-4 py-2.5 text-right">
                            <input
                              type="number"
                              value={editPrices[key] || ''}
                              onChange={e => setEditPrices({ ...editPrices, [key]: e.target.value })}
                              placeholder="Nhập giá"
                              className={`border rounded px-3 py-1.5 text-sm w-36 text-right ${!hasPrice ? 'border-amber-400 bg-amber-50' : ''}`}
                            />
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
