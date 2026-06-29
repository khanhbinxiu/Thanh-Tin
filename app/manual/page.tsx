'use client'
import { useEffect, useState } from 'react'
import { supabase } from '@/lib/supabase'
import crypto from 'crypto'

type Mapping = {
  input_key: string
  customer_code: string
  output_file_name: string
  output_sheet_name: string
  output_location_name: string
}

type Row = {
  day: number
  input_key: string
  b45_delivered: number
  b45_returned: number
  b12_delivered: number
  b12_returned: number
  gas_returned: number
  unit_price: number
  note: string
}

const emptyRow = (): Row => ({
  day: new Date().getDate(), input_key: '',
  b45_delivered: 0, b45_returned: 0, b12_delivered: 0, b12_returned: 0,
  gas_returned: 0, unit_price: 0, note: '',
})

type PriceMap = Record<string, number>

export default function ManualPage() {
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [mappings, setMappings] = useState<Mapping[]>([])
  const [rows, setRows] = useState<Row[]>([emptyRow()])
  const [msg, setMsg] = useState('')
  const [saving, setSaving] = useState(false)
  const [prices, setPrices] = useState<PriceMap>({})
  const trangThai = 'đã giao'

  useEffect(() => {
    supabase.from('location_mappings').select('*').order('input_key')
      .then(({ data }) => setMappings(data || []))
  }, [])

  useEffect(() => {
    async function loadPrices() {
      const lastDay = new Date(year, month, 0).getDate()
      const from = `${year}-${String(month).padStart(2, '0')}-01`
      const to = `${year}-${String(month).padStart(2, '0')}-${lastDay}`
      const { data } = await supabase.from('transactions')
        .select('customer_code, unit_price')
        .gte('delivery_date', from).lte('delivery_date', to)
        .gt('unit_price', 0)
      const pm: PriceMap = {}
      for (const tx of (data || [])) {
        if (tx.unit_price && !pm[tx.customer_code]) pm[tx.customer_code] = tx.unit_price
      }
      setPrices(pm)
    }
    loadPrices()
  }, [month, year])

  function updateRow(idx: number, field: keyof Row, value: string | number) {
    const updated = [...rows]
    updated[idx] = { ...updated[idx], [field]: value }
    if (field === 'input_key') {
      const key = String(value).toLowerCase()
      const m = mappings.find(m => m.input_key === key)
      if (m && prices[m.customer_code] && !updated[idx].unit_price) {
        updated[idx].unit_price = prices[m.customer_code]
      }
    }
    setRows(updated)
  }

  function addRow() {
    const last = rows[rows.length - 1]
    setRows([...rows, { ...emptyRow(), day: last.day }])
    setTimeout(() => {
      const el = document.getElementById(`row-key-${rows.length}`)
      el?.focus()
    }, 50)
  }

  function removeRow(idx: number) {
    if (rows.length === 1) return
    setRows(rows.filter((_, i) => i !== idx))
  }

  function gasDelivered(r: Row) { return r.b45_delivered * 45 + r.b12_delivered * 12 }
  function gasPaid(r: Row) { return gasDelivered(r) - r.gas_returned }

  async function save() {
    const valid = rows.filter(r => r.input_key)
    if (valid.length === 0) { setMsg('Chưa có dòng nào hợp lệ'); return }
    setSaving(true)
    setMsg('')
    const mm = new Map(mappings.map(m => [m.input_key, m]))
    const toInsert = valid.map(r => {
      const key = r.input_key.toLowerCase()
      const m = mm.get(key)
      const date = `${year}-${String(month).padStart(2, '0')}-${String(r.day).padStart(2, '0')}`
      const gd = gasDelivered(r)
      const gp = gasPaid(r)
      const total = gp * r.unit_price
      const hash = crypto.createHash('md5').update(`${date}|${key}|${gd}|${total}`).digest('hex')
      return {
        input_key: key, delivery_date: date,
        customer_code: m?.customer_code || 'UNMAPPED',
        location: m?.output_location_name || key,
        output_file_name: m?.output_file_name || null,
        output_sheet_name: m?.output_sheet_name || null,
        pic: '',
        b45_delivered: r.b45_delivered, b45_returned: r.b45_returned,
        b12_delivered: r.b12_delivered, b12_returned: r.b12_returned,
        gas_delivered: gd, gas_returned: r.gas_returned, gas_paid: gp,
        unit_price: r.unit_price, total_amount: total,
        note: r.note, month, year, dedup_hash: hash, trang_thai: trangThai,
      }
    })
    const hashes = toInsert.map(r => r.dedup_hash)
    const { data: existing } = await supabase.from('transactions').select('dedup_hash').in('dedup_hash', hashes)
    const exSet = new Set((existing || []).map(e => e.dedup_hash))
    const newRows = toInsert.filter(r => !exSet.has(r.dedup_hash))
    const dupes = toInsert.length - newRows.length
    if (newRows.length === 0) {
      setMsg(`Tất cả ${toInsert.length} dòng đã tồn tại (trùng)`)
      setSaving(false)
      return
    }
    const { error } = await supabase.from('transactions').insert(newRows)
    if (error) {
      setMsg('Lỗi: ' + error.message)
    } else {
      setMsg(`Đã lưu ${newRows.length} dòng!${dupes > 0 ? ` (${dupes} trùng)` : ''}`)
      setRows([emptyRow()])
    }
    setSaving(false)
  }

  const inputKeys = mappings.map(m => m.input_key).sort()

  return (
    <div className="space-y-4 pb-24">
      <h1 className="text-lg md:text-xl font-bold text-gray-800">Nhập tay giao dịch</h1>

      <div className="bg-white rounded-xl border p-3 flex flex-wrap gap-3 items-end">
        <div className="flex gap-2">
          <div>
            <label className="block text-xs text-gray-500 mb-1">Tháng</label>
            <select value={month} onChange={e => setMonth(Number(e.target.value))}
              className="border rounded px-2 py-2 text-sm">
              {Array.from({ length: 12 }, (_, i) => i + 1).map(m =>
                <option key={m} value={m}>T{m}</option>
              )}
            </select>
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Năm</label>
            <input type="number" value={year} onChange={e => setYear(Number(e.target.value))}
              className="border rounded px-2 py-2 text-sm w-20" />
          </div>
        </div>
      </div>

      {msg && (
        <div className={`rounded-lg px-3 py-2.5 text-sm border ${msg.includes('Lỗi') ? 'bg-red-50 border-red-200 text-red-800' : 'bg-green-50 border-green-200 text-green-800'}`}>
          {msg}
        </div>
      )}

      <datalist id="input-keys">
        {inputKeys.map(k => <option key={k} value={k} />)}
      </datalist>

      {/* Desktop: table */}
      <div className="hidden md:block bg-white rounded-xl border overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Ngày','Mã KH','B45↓','B45↑','B12↓','B12↑','Gas giao','Gas trả','Gas TT','Đơn giá','Thành tiền','Ghi chú',''].map(h => (
                <th key={h} className="px-2 py-2.5 font-medium text-gray-600 whitespace-nowrap text-center">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => {
              const gp = gasPaid(r)
              const total = gp * r.unit_price
              const m = mappings.find(m => m.input_key === r.input_key)
              return (
                <tr key={i} className={`border-b ${!m && r.input_key ? 'bg-amber-50' : 'hover:bg-gray-50'}`}>
                  <td className="px-1 py-1"><input type="number" min={1} max={31} value={r.day} onChange={e => updateRow(i,'day',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-14 text-center" /></td>
                  <td className="px-1 py-1"><input id={`row-key-${i}`} list="input-keys" value={r.input_key} onChange={e => updateRow(i,'input_key',e.target.value.toLowerCase())} placeholder="Mã KH" className="border rounded px-2 py-1.5 text-sm w-36" /></td>
                  <td className="px-1 py-1"><input type="number" value={r.b45_delivered||''} onChange={e => updateRow(i,'b45_delivered',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                  <td className="px-1 py-1"><input type="number" value={r.b45_returned||''} onChange={e => updateRow(i,'b45_returned',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                  <td className="px-1 py-1"><input type="number" value={r.b12_delivered||''} onChange={e => updateRow(i,'b12_delivered',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                  <td className="px-1 py-1"><input type="number" value={r.b12_returned||''} onChange={e => updateRow(i,'b12_returned',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                  <td className="px-2 py-1 text-right text-gray-700">{gasDelivered(r)||''}</td>
                  <td className="px-1 py-1"><input type="number" step="0.1" value={r.gas_returned||''} onChange={e => updateRow(i,'gas_returned',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-16 text-right" /></td>
                  <td className="px-2 py-1 text-right text-blue-600 font-medium">{gp>0?gp.toFixed(1):''}</td>
                  <td className="px-1 py-1"><input type="number" value={r.unit_price||''} onChange={e => updateRow(i,'unit_price',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-20 text-right" /></td>
                  <td className="px-2 py-1 text-right font-medium">{total>0?total.toLocaleString('vi-VN'):''}</td>
                  <td className="px-1 py-1"><input value={r.note} onChange={e => updateRow(i,'note',e.target.value)} className="border rounded px-2 py-1.5 text-sm w-20" /></td>
                  <td className="px-1 py-1"><button onClick={() => removeRow(i)} className="text-red-400 hover:text-red-600 text-lg">×</button></td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      {/* Mobile: cards */}
      <div className="md:hidden space-y-3">
        {rows.map((r, i) => {
          const gp = gasPaid(r)
          const total = gp * r.unit_price
          const m = mappings.find(m => m.input_key === r.input_key)
          return (
            <div key={i} className={`rounded-xl border p-3 space-y-2 ${!m && r.input_key ? 'bg-amber-50 border-amber-200' : 'bg-white'}`}>
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-gray-400">#{i + 1}</span>
                <button onClick={() => removeRow(i)} className="text-red-400 hover:text-red-600 text-sm">Xóa</button>
              </div>

              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block text-xs text-gray-500">Ngày</label>
                  <input type="number" min={1} max={31} value={r.day}
                    onChange={e => updateRow(i, 'day', Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" />
                </div>
                <div className="col-span-2">
                  <label className="block text-xs text-gray-500">Mã KH</label>
                  <input id={`row-key-${i}`} list="input-keys" value={r.input_key}
                    onChange={e => updateRow(i, 'input_key', e.target.value.toLowerCase())}
                    placeholder="Chọn/nhập mã KH"
                    className="border rounded px-2 py-2 text-sm w-full" />
                </div>
              </div>

              <div className="grid grid-cols-4 gap-2">
                <div>
                  <label className="block text-xs text-gray-500">B45↓</label>
                  <input type="number" value={r.b45_delivered||''} onChange={e => updateRow(i,'b45_delivered',Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" />
                </div>
                <div>
                  <label className="block text-xs text-gray-500">B45↑</label>
                  <input type="number" value={r.b45_returned||''} onChange={e => updateRow(i,'b45_returned',Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" />
                </div>
                <div>
                  <label className="block text-xs text-gray-500">B12↓</label>
                  <input type="number" value={r.b12_delivered||''} onChange={e => updateRow(i,'b12_delivered',Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" />
                </div>
                <div>
                  <label className="block text-xs text-gray-500">B12↑</label>
                  <input type="number" value={r.b12_returned||''} onChange={e => updateRow(i,'b12_returned',Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" />
                </div>
              </div>

              <div className="grid grid-cols-3 gap-2">
                <div className="bg-gray-50 rounded px-2 py-2 text-center">
                  <label className="block text-xs text-gray-400">Gas giao</label>
                  <span className="text-sm font-medium">{gasDelivered(r)||'-'}</span>
                </div>
                <div>
                  <label className="block text-xs text-gray-500">Gas trả</label>
                  <input type="number" step="0.1" value={r.gas_returned||''} onChange={e => updateRow(i,'gas_returned',Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="decimal" />
                </div>
                <div className="bg-blue-50 rounded px-2 py-2 text-center">
                  <label className="block text-xs text-blue-500">Gas TT</label>
                  <span className="text-sm font-bold text-blue-700">{gp>0?gp.toFixed(1):'-'}</span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-xs text-gray-500">Đơn giá {(() => { const mp = mappings.find(mp => mp.input_key === r.input_key); return mp && prices[mp.customer_code] && !r.unit_price ? `(gợi ý: ${prices[mp.customer_code].toLocaleString('vi-VN')})` : '' })()}</label>
                  <input type="number" value={r.unit_price||''} onChange={e => updateRow(i,'unit_price',Number(e.target.value))}
                    className="border rounded px-2 py-2 text-sm w-full text-right" inputMode="numeric" />
                </div>
                <div>
                  <label className="block text-xs text-gray-500">Ghi chú</label>
                  <input value={r.note} onChange={e => updateRow(i,'note',e.target.value)}
                    className="border rounded px-2 py-2 text-sm w-full" />
                </div>
              </div>

              {total > 0 && (
                <div className="text-right text-sm font-bold text-green-700">
                  = {total.toLocaleString('vi-VN')} đ
                </div>
              )}
            </div>
          )
        })}
      </div>

      {/* Fixed bottom bar */}
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t shadow-lg p-3 flex gap-3 justify-center z-40">
        <button onClick={addRow}
          className="border border-blue-300 text-blue-600 px-4 py-2.5 rounded-lg text-sm font-medium hover:bg-blue-50 flex-1 max-w-40">
          + Thêm dòng
        </button>
        <button onClick={save} disabled={saving}
          className="bg-green-600 text-white px-4 py-2.5 rounded-lg text-sm font-medium hover:bg-green-700 disabled:opacity-50 flex-1 max-w-52">
          {saving ? 'Đang lưu...' : `Lưu ${rows.filter(r => r.input_key).length} dòng`}
        </button>
      </div>
    </div>
  )
}
