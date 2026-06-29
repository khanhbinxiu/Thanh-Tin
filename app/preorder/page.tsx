'use client'
import { useEffect, useState } from 'react'
import { supabase, Transaction } from '@/lib/supabase'
import crypto from 'crypto'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }

type Mapping = { input_key: string; customer_code: string; output_file_name: string; output_sheet_name: string; output_location_name: string }
type Row = { day: number; input_key: string; b45_delivered: number; b45_returned: number; b12_delivered: number; b12_returned: number; gas_returned: number; unit_price: number; note: string }
type EditRow = { delivery_date: string; b45_delivered: string; b45_returned: string; b12_delivered: string; b12_returned: string; gas_returned: string; unit_price: string; note: string }
type PriceMap = Record<string, number>

const emptyRow = (): Row => ({ day: new Date().getDate(), input_key: '', b45_delivered: 0, b45_returned: 0, b12_delivered: 0, b12_returned: 0, gas_returned: 0, unit_price: 0, note: '' })

function toEdit(r: Transaction): EditRow {
  return { delivery_date: r.delivery_date, b45_delivered: String(r.b45_delivered||0), b45_returned: String(r.b45_returned||0), b12_delivered: String(r.b12_delivered||0), b12_returned: String(r.b12_returned||0), gas_returned: String(r.gas_returned||0), unit_price: String(r.unit_price||0), note: r.note||'' }
}

export default function PreorderPage() {
  const [tab, setTab] = useState<'input' | 'pending'>('input')
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [mappings, setMappings] = useState<Mapping[]>([])
  const [prices, setPrices] = useState<PriceMap>({})

  // Input state
  const [rows, setRows] = useState<Row[]>([emptyRow()])
  const [msg, setMsg] = useState('')
  const [saving, setSaving] = useState(false)

  // Pending state
  const [pending, setPending] = useState<Transaction[]>([])
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [edits, setEdits] = useState<Record<string, EditRow>>({})
  const [loadingPending, setLoadingPending] = useState(false)

  useEffect(() => {
    supabase.from('location_mappings').select('*').order('input_key').then(({ data }) => setMappings(data || []))
  }, [])

  useEffect(() => {
    async function lp() {
      const lastDay = new Date(year, month, 0).getDate()
      const { data } = await supabase.from('transactions').select('customer_code, unit_price')
        .gte('delivery_date', `${year}-${String(month).padStart(2,'0')}-01`)
        .lte('delivery_date', `${year}-${String(month).padStart(2,'0')}-${lastDay}`)
        .gt('unit_price', 0)
      const pm: PriceMap = {}
      for (const tx of (data || [])) { if (tx.unit_price && !pm[tx.customer_code]) pm[tx.customer_code] = tx.unit_price }
      setPrices(pm)
    }
    lp()
  }, [month, year])

  async function loadPending() {
    setLoadingPending(true)
    const { data } = await supabase.from('transactions').select('*').eq('trang_thai', 'đặt trước').order('delivery_date')
    setPending(data || []); setEdits({}); setSelected(new Set())
    setLoadingPending(false)
  }

  useEffect(() => { if (tab === 'pending') loadPending() }, [tab])

  // Input functions
  function updateRow(idx: number, field: keyof Row, value: string | number) {
    const updated = [...rows]; updated[idx] = { ...updated[idx], [field]: value }
    if (field === 'input_key') {
      const key = String(value).toLowerCase()
      const m = mappings.find(m => m.input_key === key)
      if (m && prices[m.customer_code] && !updated[idx].unit_price) updated[idx].unit_price = prices[m.customer_code]
    }
    setRows(updated)
  }
  function addRow() {
    const last = rows[rows.length - 1]
    setRows([...rows, { ...emptyRow(), day: last.day }])
    setTimeout(() => document.getElementById(`po-key-${rows.length}`)?.focus(), 50)
  }
  function removeRow(idx: number) { if (rows.length > 1) setRows(rows.filter((_, i) => i !== idx)) }
  function gasDelivered(r: Row) { return r.b45_delivered * 45 + r.b12_delivered * 12 }
  function gasPaid(r: Row) { return gasDelivered(r) - r.gas_returned }

  async function saveOrders() {
    const valid = rows.filter(r => r.input_key)
    if (!valid.length) { setMsg('Chưa có dòng hợp lệ'); return }
    setSaving(true); setMsg('')
    const mm = new Map(mappings.map(m => [m.input_key, m]))
    const toInsert = valid.map(r => {
      const key = r.input_key.toLowerCase(); const m = mm.get(key)
      const date = `${year}-${String(month).padStart(2,'0')}-${String(r.day).padStart(2,'0')}`
      const gd = gasDelivered(r); const gp = gasPaid(r); const total = gp * r.unit_price
      const hash = crypto.createHash('md5').update(`${date}|${key}|${gd}|${total}|preorder`).digest('hex')
      return {
        input_key: key, delivery_date: date, customer_code: m?.customer_code || 'UNMAPPED',
        location: m?.output_location_name || key, output_file_name: m?.output_file_name || null,
        output_sheet_name: m?.output_sheet_name || null, pic: '',
        b45_delivered: r.b45_delivered, b45_returned: r.b45_returned,
        b12_delivered: r.b12_delivered, b12_returned: r.b12_returned,
        gas_delivered: gd, gas_returned: r.gas_returned, gas_paid: gp,
        unit_price: r.unit_price, total_amount: total, note: r.note,
        month, year, dedup_hash: hash, trang_thai: 'đặt trước',
      }
    })
    const { error } = await supabase.from('transactions').insert(toInsert)
    if (error) { setMsg('Lỗi: ' + error.message) }
    else { setMsg(`Đã đặt trước ${toInsert.length} đơn!`); setRows([emptyRow()]) }
    setSaving(false)
  }

  // Pending functions
  function toggleSelect(id: string) {
    const ns = new Set(selected)
    if (ns.has(id)) { ns.delete(id); const ne = { ...edits }; delete ne[id]; setEdits(ne) }
    else { ns.add(id); const r = pending.find(r => r.id === id); if (r) setEdits(prev => ({ ...prev, [id]: toEdit(r) })) }
    setSelected(ns)
  }
  function toggleAll() {
    if (selected.size === pending.length) { setSelected(new Set()); setEdits({}) }
    else { setSelected(new Set(pending.map(r => r.id))); const ne: Record<string, EditRow> = {}; for (const r of pending) ne[r.id] = toEdit(r); setEdits(ne) }
  }
  function updateEdit(id: string, field: keyof EditRow, value: string) { setEdits(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } })) }

  async function confirmSelected() {
    if (!selected.size) return; setSaving(true); setMsg('')
    for (const id of selected) {
      const e = edits[id]; if (!e) { await supabase.from('transactions').update({ trang_thai: 'đã giao' }).eq('id', id); continue }
      const b45d = parseFloat(e.b45_delivered)||0; const b12d = parseFloat(e.b12_delivered)||0
      const gasRet = parseFloat(e.gas_returned)||0; const unitPrice = parseFloat(e.unit_price)||0
      const gd = b45d * 45 + b12d * 12; const gp = gd - gasRet
      await supabase.from('transactions').update({
        trang_thai: 'đã giao', delivery_date: e.delivery_date,
        b45_delivered: b45d, b45_returned: parseFloat(e.b45_returned)||0,
        b12_delivered: b12d, b12_returned: parseFloat(e.b12_returned)||0,
        gas_delivered: gd, gas_returned: gasRet, gas_paid: gp,
        unit_price: unitPrice, total_amount: gp * unitPrice, note: e.note,
      }).eq('id', id)
    }
    setMsg(`Đã xác nhận giao ${selected.size} đơn`); setSaving(false); await loadPending()
  }

  async function deleteSelected() {
    if (!selected.size || !confirm(`Xóa ${selected.size} đơn?`)) return
    setSaving(true)
    await supabase.from('transactions').delete().in('id', [...selected])
    setMsg(`Đã xóa ${selected.size} đơn`); setSaving(false); await loadPending()
  }

  const inputKeys = mappings.map(m => m.input_key).sort()

  return (
    <div className="space-y-4 pb-24">
      <h1 className="text-lg md:text-xl font-bold text-gray-800">Đặt trước</h1>

      <div className="flex gap-1 bg-gray-100 rounded-lg p-1 w-fit">
        <button onClick={() => setTab('input')} className={`px-4 py-2 rounded-md text-sm font-medium transition ${tab === 'input' ? 'bg-amber-500 text-white shadow' : 'text-gray-500'}`}>
          Nhập đơn đặt trước
        </button>
        <button onClick={() => setTab('pending')} className={`px-4 py-2 rounded-md text-sm font-medium transition ${tab === 'pending' ? 'bg-amber-500 text-white shadow' : 'text-gray-500'}`}>
          Chờ giao ({pending.length || '...'})
        </button>
      </div>

      {msg && <div className={`rounded-lg px-3 py-2.5 text-sm border ${msg.includes('Lỗi') ? 'bg-red-50 border-red-200 text-red-800' : 'bg-green-50 border-green-200 text-green-800'}`}>{msg}</div>}

      {/* Tab: Nhập đơn */}
      {tab === 'input' && (
        <>
          <div className="bg-white rounded-xl border p-3 flex flex-wrap gap-3 items-end">
            <div className="flex gap-2">
              <div>
                <label className="block text-xs text-gray-500 mb-1">Tháng</label>
                <select value={month} onChange={e => setMonth(Number(e.target.value))} className="border rounded px-2 py-2 text-sm">
                  {Array.from({ length: 12 }, (_, i) => i + 1).map(m => <option key={m} value={m}>T{m}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Năm</label>
                <input type="number" value={year} onChange={e => setYear(Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-20" />
              </div>
            </div>
          </div>

          <datalist id="po-keys">{inputKeys.map(k => <option key={k} value={k} />)}</datalist>

          {/* Mobile cards */}
          <div className="md:hidden space-y-3">
            {rows.map((r, i) => {
              const gp = gasPaid(r); const total = gp * r.unit_price
              return (
                <div key={i} className="rounded-xl border bg-white p-3 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-amber-500">Đặt trước #{i + 1}</span>
                    <button onClick={() => removeRow(i)} className="text-red-400 text-sm">Xóa</button>
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    <div><label className="block text-xs text-gray-500">Ngày</label><input type="number" min={1} max={31} value={r.day} onChange={e => updateRow(i,'day',Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" /></div>
                    <div className="col-span-2"><label className="block text-xs text-gray-500">Mã KH</label><input id={`po-key-${i}`} list="po-keys" value={r.input_key} onChange={e => updateRow(i,'input_key',e.target.value.toLowerCase())} placeholder="Chọn/nhập" className="border rounded px-2 py-2 text-sm w-full" /></div>
                  </div>
                  <div className="grid grid-cols-4 gap-2">
                    <div><label className="block text-xs text-gray-500">B45↓</label><input type="number" value={r.b45_delivered||''} onChange={e => updateRow(i,'b45_delivered',Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" /></div>
                    <div><label className="block text-xs text-gray-500">B45↑</label><input type="number" value={r.b45_returned||''} onChange={e => updateRow(i,'b45_returned',Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" /></div>
                    <div><label className="block text-xs text-gray-500">B12↓</label><input type="number" value={r.b12_delivered||''} onChange={e => updateRow(i,'b12_delivered',Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" /></div>
                    <div><label className="block text-xs text-gray-500">B12↑</label><input type="number" value={r.b12_returned||''} onChange={e => updateRow(i,'b12_returned',Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="numeric" /></div>
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    <div className="bg-gray-50 rounded px-2 py-2 text-center"><label className="block text-xs text-gray-400">Gas giao</label><span className="text-sm font-medium">{gasDelivered(r)||'-'}</span></div>
                    <div><label className="block text-xs text-gray-500">Gas trả</label><input type="number" step="0.1" value={r.gas_returned||''} onChange={e => updateRow(i,'gas_returned',Number(e.target.value))} className="border rounded px-2 py-2 text-sm w-full text-center" inputMode="decimal" /></div>
                    <div className="bg-amber-50 rounded px-2 py-2 text-center"><label className="block text-xs text-amber-500">Gas TT</label><span className="text-sm font-bold text-amber-700">{gp>0?gp.toFixed(1):'-'}</span></div>
                  </div>
                  <div><label className="block text-xs text-gray-500">Ghi chú</label><input value={r.note} onChange={e => updateRow(i,'note',e.target.value)} className="border rounded px-2 py-2 text-sm w-full" /></div>
                  {total > 0 && <div className="text-right text-sm font-bold text-amber-700">= {total.toLocaleString('vi-VN')} đ</div>}
                </div>
              )
            })}
          </div>

          {/* Desktop table */}
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
                  const gp = gasPaid(r); const total = gp * r.unit_price
                  return (
                    <tr key={i} className="border-b hover:bg-gray-50">
                      <td className="px-1 py-1"><input type="number" min={1} max={31} value={r.day} onChange={e => updateRow(i,'day',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-14 text-center" /></td>
                      <td className="px-1 py-1"><input id={`po-key-${i}`} list="po-keys" value={r.input_key} onChange={e => updateRow(i,'input_key',e.target.value.toLowerCase())} placeholder="Mã KH" className="border rounded px-2 py-1.5 text-sm w-36" /></td>
                      <td className="px-1 py-1"><input type="number" value={r.b45_delivered||''} onChange={e => updateRow(i,'b45_delivered',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                      <td className="px-1 py-1"><input type="number" value={r.b45_returned||''} onChange={e => updateRow(i,'b45_returned',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                      <td className="px-1 py-1"><input type="number" value={r.b12_delivered||''} onChange={e => updateRow(i,'b12_delivered',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                      <td className="px-1 py-1"><input type="number" value={r.b12_returned||''} onChange={e => updateRow(i,'b12_returned',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-12 text-right" /></td>
                      <td className="px-2 py-1 text-right text-gray-700">{gasDelivered(r)||''}</td>
                      <td className="px-1 py-1"><input type="number" step="0.1" value={r.gas_returned||''} onChange={e => updateRow(i,'gas_returned',Number(e.target.value))} className="border rounded px-2 py-1.5 text-sm w-16 text-right" /></td>
                      <td className="px-2 py-1 text-right text-amber-600 font-medium">{gp>0?gp.toFixed(1):''}</td>
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

          <div className="fixed bottom-0 left-0 right-0 bg-white border-t shadow-lg p-3 flex gap-3 justify-center z-40">
            <button onClick={addRow} className="border border-amber-300 text-amber-600 px-4 py-2.5 rounded-lg text-sm font-medium hover:bg-amber-50 flex-1 max-w-40">+ Thêm dòng</button>
            <button onClick={saveOrders} disabled={saving} className="bg-amber-500 text-white px-4 py-2.5 rounded-lg text-sm font-medium hover:bg-amber-600 disabled:opacity-50 flex-1 max-w-52">
              {saving ? 'Đang lưu...' : `Đặt trước ${rows.filter(r => r.input_key).length} đơn`}
            </button>
          </div>
        </>
      )}

      {/* Tab: Chờ giao */}
      {tab === 'pending' && (
        <>
          {selected.size > 0 && (
            <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3 flex flex-wrap gap-3">
              <button onClick={confirmSelected} disabled={saving} className="bg-green-600 text-white px-4 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
                {saving ? 'Đang xử lý...' : `Xác nhận đã giao (${selected.size})`}
              </button>
              <button onClick={deleteSelected} disabled={saving} className="bg-red-600 text-white px-4 py-2 rounded text-sm hover:bg-red-700 disabled:opacity-50">
                Xóa ({selected.size})
              </button>
            </div>
          )}

          <div className="bg-white rounded-xl border overflow-auto">
            <table className="w-full text-xs">
              <thead className="bg-gray-50 border-b sticky top-0">
                <tr>
                  <th className="px-2 py-2"><input type="checkbox" checked={selected.size === pending.length && pending.length > 0} onChange={toggleAll} /></th>
                  {['Ngày','Mã KH','Địa điểm','B45↓','B45↑','B12↓','B12↑','Gas giao','Gas trả','Gas TT','Đơn giá','Thành tiền','Ghi chú'].map(h => (
                    <th key={h} className="text-left px-2 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {loadingPending && <tr><td colSpan={14} className="px-4 py-8 text-center text-gray-400">Đang tải...</td></tr>}
                {!loadingPending && pending.map(r => {
                  const isEd = selected.has(r.id); const e = edits[r.id]
                  const b45d = isEd ? parseFloat(e.b45_delivered)||0 : r.b45_delivered
                  const b12d = isEd ? parseFloat(e.b12_delivered)||0 : r.b12_delivered
                  const gasRet = isEd ? parseFloat(e.gas_returned)||0 : r.gas_returned
                  const uP = isEd ? parseFloat(e.unit_price)||0 : r.unit_price
                  const gd = b45d * 45 + b12d * 12; const gp = gd - gasRet; const total = gp * uP
                  const ic = 'border border-amber-300 bg-amber-50 rounded px-1.5 py-1 text-xs text-right w-full'
                  return (
                    <tr key={r.id} className={`border-b ${isEd ? 'bg-amber-50' : 'hover:bg-gray-50'}`}>
                      <td className="px-2 py-1.5 text-center"><input type="checkbox" checked={isEd} onChange={() => toggleSelect(r.id)} /></td>
                      <td className="px-2 py-1.5 whitespace-nowrap">
                        {isEd ? <input type="date" value={e.delivery_date} onChange={ev => updateEdit(r.id,'delivery_date',ev.target.value)} className="border border-amber-300 bg-amber-50 rounded px-1 py-1 text-xs w-28" /> : fmtDate(r.delivery_date)}
                      </td>
                      <td className="px-2 py-1.5 font-medium">{r.customer_code}</td>
                      <td className="px-2 py-1.5">{r.location}</td>
                      <td className="px-1 py-1" style={{minWidth:50}}>{isEd ? <input type="number" value={e.b45_delivered} onChange={ev => updateEdit(r.id,'b45_delivered',ev.target.value)} className={ic} /> : r.b45_delivered||''}</td>
                      <td className="px-1 py-1" style={{minWidth:50}}>{isEd ? <input type="number" value={e.b45_returned} onChange={ev => updateEdit(r.id,'b45_returned',ev.target.value)} className={ic} /> : r.b45_returned||''}</td>
                      <td className="px-1 py-1" style={{minWidth:50}}>{isEd ? <input type="number" value={e.b12_delivered} onChange={ev => updateEdit(r.id,'b12_delivered',ev.target.value)} className={ic} /> : r.b12_delivered||''}</td>
                      <td className="px-1 py-1" style={{minWidth:50}}>{isEd ? <input type="number" value={e.b12_returned} onChange={ev => updateEdit(r.id,'b12_returned',ev.target.value)} className={ic} /> : r.b12_returned||''}</td>
                      <td className="px-2 py-1.5 text-right text-gray-500">{gd||''}</td>
                      <td className="px-1 py-1" style={{minWidth:60}}>{isEd ? <input type="number" step="0.1" value={e.gas_returned} onChange={ev => updateEdit(r.id,'gas_returned',ev.target.value)} className={ic} /> : r.gas_returned||''}</td>
                      <td className="px-2 py-1.5 text-right text-amber-600 font-medium">{gp?gp.toFixed(1):''}</td>
                      <td className="px-1 py-1" style={{minWidth:70}}>{isEd ? <input type="number" value={e.unit_price} onChange={ev => updateEdit(r.id,'unit_price',ev.target.value)} className={ic} /> : (r.unit_price?r.unit_price.toLocaleString('vi-VN'):'')}</td>
                      <td className="px-2 py-1.5 text-right font-medium">{total?total.toLocaleString('vi-VN'):''}</td>
                      <td className="px-1 py-1" style={{minWidth:60}}>{isEd ? <input value={e.note} onChange={ev => updateEdit(r.id,'note',ev.target.value)} className="border border-amber-300 bg-amber-50 rounded px-1.5 py-1 text-xs w-full" /> : <span className="text-gray-500">{r.note}</span>}</td>
                    </tr>
                  )
                })}
                {!loadingPending && pending.length === 0 && <tr><td colSpan={14} className="px-4 py-8 text-center text-gray-400">Không có đơn đặt trước</td></tr>}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  )
}
