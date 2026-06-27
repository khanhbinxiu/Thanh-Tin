'use client'
import { useEffect, useState, useCallback } from 'react'
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
  pic: string
  b45_delivered: number
  b45_returned: number
  b12_delivered: number
  b12_returned: number
  gas_delivered: number
  gas_returned: number
  unit_price: number
  note: string
}

const emptyRow = (): Row => ({
  day: new Date().getDate(), input_key: '', pic: '',
  b45_delivered: 0, b45_returned: 0, b12_delivered: 0, b12_returned: 0,
  gas_delivered: 0, gas_returned: 0, unit_price: 0, note: '',
})

export default function ManualPage() {
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [mappings, setMappings] = useState<Mapping[]>([])
  const [rows, setRows] = useState<Row[]>([emptyRow()])
  const [msg, setMsg] = useState('')
  const [saving, setSaving] = useState(false)

  useEffect(() => {
    supabase.from('location_mappings').select('*').order('input_key')
      .then(({ data }) => setMappings(data || []))
  }, [])

  function updateRow(idx: number, field: keyof Row, value: string | number) {
    const updated = [...rows]
    updated[idx] = { ...updated[idx], [field]: value }
    setRows(updated)
  }

  function addRow() {
    const last = rows[rows.length - 1]
    setRows([...rows, { ...emptyRow(), day: last.day, pic: last.pic }])
    setTimeout(() => {
      const inputs = document.querySelectorAll<HTMLInputElement>('table input[list="input-keys"]')
      inputs[inputs.length - 1]?.focus()
    }, 50)
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (e.key === 'Enter') { e.preventDefault(); addRow() }
  }

  function removeRow(idx: number) {
    if (rows.length === 1) return
    setRows(rows.filter((_, i) => i !== idx))
  }

  function gasPaid(r: Row) {
    return r.b45_delivered * 45 + r.b12_delivered * 12 - r.gas_returned
  }

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
      const gp = gasPaid(r)
      const total = gp * r.unit_price
      const gasDelivered = r.b45_delivered * 45 + r.b12_delivered * 12
      const hash = crypto.createHash('md5').update(`${date}|${key}|${gasDelivered}|${total}`).digest('hex')
      return {
        input_key: key,
        delivery_date: date,
        customer_code: m?.customer_code || 'UNMAPPED',
        location: m?.output_location_name || key,
        output_file_name: m?.output_file_name || null,
        output_sheet_name: m?.output_sheet_name || null,
        pic: r.pic,
        b45_delivered: r.b45_delivered,
        b45_returned: r.b45_returned,
        b12_delivered: r.b12_delivered,
        b12_returned: r.b12_returned,
        gas_delivered: gasDelivered,
        gas_returned: r.gas_returned,
        gas_paid: gp,
        unit_price: r.unit_price,
        total_amount: total,
        note: r.note,
        month, year,
        dedup_hash: hash,
      }
    })

    // Dedup check
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
      setMsg(`Đã lưu ${newRows.length} dòng!${dupes > 0 ? ` (${dupes} dòng trùng, bỏ qua)` : ''}`)
      setRows([emptyRow()])
    }
    setSaving(false)
  }

  const inputKeys = mappings.map(m => m.input_key).sort()

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Nhập tay giao dịch</h1>

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
      </div>

      {msg && (
        <div className={`rounded-lg px-4 py-3 text-sm border ${msg.includes('Lỗi') ? 'bg-red-50 border-red-200 text-red-800' : 'bg-green-50 border-green-200 text-green-800'}`}>
          {msg}
        </div>
      )}

      <div className="bg-white rounded-xl border overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Ngày', 'Mã KH', 'PIC', 'B45↓', 'B45↑', 'B12↓', 'B12↑', 'Gas giao', 'Gas trả', 'Gas TT', 'Đơn giá', 'Thành tiền', 'Ghi chú', ''].map(h => (
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
                  <td className="px-1 py-1.5">
                    <input type="number" min={1} max={31} value={r.day}
                      onChange={e => updateRow(i, 'day', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-14 text-center" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <input list="input-keys" value={r.input_key}
                      onChange={e => updateRow(i, 'input_key', e.target.value.toLowerCase())}
                      placeholder="Chọn/nhập"
                      className="border rounded px-2 py-1.5 text-sm w-40" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <input value={r.pic} onChange={e => updateRow(i, 'pic', e.target.value)}
                      className="border rounded px-2 py-1.5 text-sm w-16" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <input type="number" value={r.b45_delivered || ''}
                      onChange={e => updateRow(i, 'b45_delivered', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-14 text-right" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <input type="number" value={r.b45_returned || ''}
                      onChange={e => updateRow(i, 'b45_returned', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-14 text-right" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <input type="number" value={r.b12_delivered || ''}
                      onChange={e => updateRow(i, 'b12_delivered', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-14 text-right" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <input type="number" value={r.b12_returned || ''}
                      onChange={e => updateRow(i, 'b12_returned', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-14 text-right" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium text-gray-700 whitespace-nowrap">
                    {(r.b45_delivered * 45 + r.b12_delivered * 12) || ''}
                  </td>
                  <td className="px-1 py-1.5">
                    <input type="number" step="0.1" value={r.gas_returned || ''}
                      onChange={e => updateRow(i, 'gas_returned', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-16 text-right" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium text-blue-600 whitespace-nowrap">
                    {gp > 0 ? gp.toFixed(1) : ''}
                  </td>
                  <td className="px-1 py-1.5">
                    <input type="number" value={r.unit_price || ''}
                      onChange={e => updateRow(i, 'unit_price', Number(e.target.value))}
                      className="border rounded px-2 py-1.5 text-sm w-20 text-right" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium whitespace-nowrap">
                    {total > 0 ? total.toLocaleString('vi-VN') : ''}
                  </td>
                  <td className="px-1 py-1.5">
                    <input value={r.note} onChange={e => updateRow(i, 'note', e.target.value)}
                      className="border rounded px-2 py-1.5 text-sm w-24" onKeyDown={handleKeyDown} />
                  </td>
                  <td className="px-1 py-1.5">
                    <button onClick={() => removeRow(i)} className="text-red-400 hover:text-red-600 text-lg px-1"
                      title="Xóa dòng">×</button>
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
        <datalist id="input-keys">
          {inputKeys.map(k => <option key={k} value={k} />)}
        </datalist>
      </div>

      <div className="flex gap-3">
        <button onClick={addRow}
          className="border border-blue-300 text-blue-600 px-4 py-2 rounded text-sm hover:bg-blue-50">
          + Thêm dòng
        </button>
        <button onClick={save} disabled={saving}
          className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
          {saving ? 'Đang lưu...' : `Lưu ${rows.filter(r => r.input_key).length} dòng`}
        </button>
      </div>
    </div>
  )
}
