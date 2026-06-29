'use client'
import { useEffect, useState } from 'react'
import { supabase, Transaction } from '@/lib/supabase'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }

type EditRow = {
  delivery_date: string
  b45_delivered: string; b45_returned: string
  b12_delivered: string; b12_returned: string
  gas_returned: string; unit_price: string; note: string
}

function toEdit(r: Transaction): EditRow {
  return {
    delivery_date: r.delivery_date,
    b45_delivered: String(r.b45_delivered||0), b45_returned: String(r.b45_returned||0),
    b12_delivered: String(r.b12_delivered||0), b12_returned: String(r.b12_returned||0),
    gas_returned: String(r.gas_returned||0), unit_price: String(r.unit_price||0), note: r.note||'',
  }
}

export default function PreorderPage() {
  const [rows, setRows] = useState<Transaction[]>([])
  const [loading, setLoading] = useState(false)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [edits, setEdits] = useState<Record<string, EditRow>>({})
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')

  async function load() {
    setLoading(true); setMsg('')
    const { data } = await supabase.from('transactions')
      .select('*').eq('trang_thai', 'đặt trước')
      .order('delivery_date', { ascending: true })
    setRows(data || [])
    setEdits({}); setSelected(new Set())
    setLoading(false)
  }

  useEffect(() => { load() }, [])

  function toggleSelect(id: string) {
    const ns = new Set(selected)
    if (ns.has(id)) { ns.delete(id); const ne = { ...edits }; delete ne[id]; setEdits(ne) }
    else { ns.add(id); const r = rows.find(r => r.id === id); if (r) setEdits(prev => ({ ...prev, [id]: toEdit(r) })) }
    setSelected(ns)
  }

  function toggleAll() {
    if (selected.size === rows.length) { setSelected(new Set()); setEdits({}) }
    else {
      setSelected(new Set(rows.map(r => r.id)))
      const ne: Record<string, EditRow> = {}; for (const r of rows) ne[r.id] = toEdit(r)
      setEdits(ne)
    }
  }

  function updateEdit(id: string, field: keyof EditRow, value: string) {
    setEdits(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } }))
  }

  async function confirmSelected() {
    if (selected.size === 0) return
    setSaving(true); setMsg('')
    let count = 0
    for (const id of selected) {
      const e = edits[id]
      if (e) {
        const b45d = parseFloat(e.b45_delivered)||0; const b12d = parseFloat(e.b12_delivered)||0
        const gasRet = parseFloat(e.gas_returned)||0; const unitPrice = parseFloat(e.unit_price)||0
        const gasDelivered = b45d * 45 + b12d * 12
        const gasPaid = gasDelivered - gasRet
        await supabase.from('transactions').update({
          trang_thai: 'đã giao', delivery_date: e.delivery_date,
          b45_delivered: b45d, b45_returned: parseFloat(e.b45_returned)||0,
          b12_delivered: b12d, b12_returned: parseFloat(e.b12_returned)||0,
          gas_delivered: gasDelivered, gas_returned: gasRet, gas_paid: gasPaid,
          unit_price: unitPrice, total_amount: gasPaid * unitPrice, note: e.note,
        }).eq('id', id)
      } else {
        await supabase.from('transactions').update({ trang_thai: 'đã giao' }).eq('id', id)
      }
      count++
    }
    setMsg(`Đã xác nhận giao ${count} đơn`)
    setSaving(false); await load()
  }

  async function deleteSelected() {
    if (selected.size === 0) return
    if (!confirm(`Xóa ${selected.size} đơn đặt trước?`)) return
    setSaving(true)
    await supabase.from('transactions').delete().in('id', [...selected])
    setMsg(`Đã xóa ${selected.size} đơn`)
    setSaving(false); await load()
  }

  return (
    <div className="space-y-4">
      <h1 className="text-lg md:text-xl font-bold text-gray-800">Đơn đặt trước</h1>
      <p className="text-sm text-gray-500">Tick chọn đơn → sửa nếu cần → bấm "Xác nhận đã giao" để chuyển thành giao dịch thật.</p>

      {msg && <div className="bg-green-50 border border-green-200 rounded-lg px-3 py-2.5 text-sm text-green-800">{msg}</div>}

      {(selected.size > 0) && (
        <div className="bg-blue-50 border border-blue-200 rounded-lg px-4 py-3 flex flex-wrap gap-3">
          <button onClick={confirmSelected} disabled={saving}
            className="bg-green-600 text-white px-4 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
            {saving ? 'Đang xử lý...' : `Xác nhận đã giao (${selected.size})`}
          </button>
          <button onClick={deleteSelected} disabled={saving}
            className="bg-red-600 text-white px-4 py-2 rounded text-sm hover:bg-red-700 disabled:opacity-50">
            Xóa ({selected.size})
          </button>
        </div>
      )}

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
              const b45d = isEditing ? parseFloat(e.b45_delivered)||0 : r.b45_delivered
              const b12d = isEditing ? parseFloat(e.b12_delivered)||0 : r.b12_delivered
              const gasRet = isEditing ? parseFloat(e.gas_returned)||0 : r.gas_returned
              const uPrice = isEditing ? parseFloat(e.unit_price)||0 : r.unit_price
              const gasDelivered = b45d * 45 + b12d * 12
              const gasPaid = gasDelivered - gasRet
              const total = gasPaid * uPrice
              const ic = 'border border-amber-300 bg-amber-50 rounded px-1.5 py-1 text-xs text-right w-full'
              return (
                <tr key={r.id} className={`border-b ${isEditing ? 'bg-amber-50' : 'hover:bg-gray-50'}`}>
                  <td className="px-2 py-1.5 text-center"><input type="checkbox" checked={isEditing} onChange={() => toggleSelect(r.id)} /></td>
                  <td className="px-2 py-1.5 whitespace-nowrap">
                    {isEditing ? <input type="date" value={e.delivery_date} onChange={ev => updateEdit(r.id,'delivery_date',ev.target.value)} className="border border-amber-300 bg-amber-50 rounded px-1 py-1 text-xs w-28" /> : fmtDate(r.delivery_date)}
                  </td>
                  <td className="px-2 py-1.5 font-medium">{r.customer_code}</td>
                  <td className="px-2 py-1.5">{r.location}</td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b45_delivered} onChange={ev => updateEdit(r.id,'b45_delivered',ev.target.value)} className={ic} /> : r.b45_delivered||''}
                  </td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b45_returned} onChange={ev => updateEdit(r.id,'b45_returned',ev.target.value)} className={ic} /> : r.b45_returned||''}
                  </td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b12_delivered} onChange={ev => updateEdit(r.id,'b12_delivered',ev.target.value)} className={ic} /> : r.b12_delivered||''}
                  </td>
                  <td className="px-1 py-1" style={{minWidth:50}}>
                    {isEditing ? <input type="number" value={e.b12_returned} onChange={ev => updateEdit(r.id,'b12_returned',ev.target.value)} className={ic} /> : r.b12_returned||''}
                  </td>
                  <td className="px-2 py-1.5 text-right text-gray-500">{gasDelivered||''}</td>
                  <td className="px-1 py-1" style={{minWidth:60}}>
                    {isEditing ? <input type="number" step="0.1" value={e.gas_returned} onChange={ev => updateEdit(r.id,'gas_returned',ev.target.value)} className={ic} /> : r.gas_returned||''}
                  </td>
                  <td className="px-2 py-1.5 text-right text-blue-600 font-medium">{gasPaid?gasPaid.toFixed(1):''}</td>
                  <td className="px-1 py-1" style={{minWidth:70}}>
                    {isEditing ? <input type="number" value={e.unit_price} onChange={ev => updateEdit(r.id,'unit_price',ev.target.value)} className={ic} /> : (r.unit_price?r.unit_price.toLocaleString('vi-VN'):'')}
                  </td>
                  <td className="px-2 py-1.5 text-right font-medium">{total?total.toLocaleString('vi-VN'):''}</td>
                  <td className="px-1 py-1" style={{minWidth:60}}>
                    {isEditing ? <input value={e.note} onChange={ev => updateEdit(r.id,'note',ev.target.value)} className="border border-amber-300 bg-amber-50 rounded px-1.5 py-1 text-xs w-full" /> : <span className="text-gray-500">{r.note}</span>}
                  </td>
                </tr>
              )
            })}
            {!loading && rows.length === 0 && (
              <tr><td colSpan={14} className="px-4 py-8 text-center text-gray-400">Không có đơn đặt trước</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}
