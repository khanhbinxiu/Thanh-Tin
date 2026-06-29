'use client'
import { useEffect, useState } from 'react'
import { supabase } from '@/lib/supabase'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }

type InRecord = { id: string; receive_date: string; b45_count: number; b12_count: number; gas_kg: number; note: string }

export default function InventoryPage() {
  const [records, setRecords] = useState<InRecord[]>([])
  const [totals, setTotals] = useState({ inGas: 0, inB45: 0, inB12: 0, outGas: 0, outB45d: 0, outB45r: 0, outB12d: 0, outB12r: 0 })
  const [loading, setLoading] = useState(false)
  const [msg, setMsg] = useState('')
  const [form, setForm] = useState({ receive_date: new Date().toISOString().slice(0, 10), b45_count: '', b12_count: '', gas_kg: '', note: '' })
  const [editing, setEditing] = useState<string | null>(null)

  async function load() {
    setLoading(true)
    const [{ data: inv }, { data: txs }] = await Promise.all([
      supabase.from('inventory_in').select('*').order('receive_date', { ascending: false }),
      supabase.from('transactions').select('gas_delivered, b45_delivered, b45_returned, b12_delivered, b12_returned').eq('trang_thai', 'đã giao'),
    ])
    setRecords(inv || [])
    const inGas = (inv || []).reduce((s, r) => s + (r.gas_kg || 0), 0)
    const inB45 = (inv || []).reduce((s, r) => s + (r.b45_count || 0), 0)
    const inB12 = (inv || []).reduce((s, r) => s + (r.b12_count || 0), 0)
    const outGas = (txs || []).reduce((s, r) => s + (r.gas_delivered || 0), 0)
    const outB45d = (txs || []).reduce((s, r) => s + (r.b45_delivered || 0), 0)
    const outB45r = (txs || []).reduce((s, r) => s + (r.b45_returned || 0), 0)
    const outB12d = (txs || []).reduce((s, r) => s + (r.b12_delivered || 0), 0)
    const outB12r = (txs || []).reduce((s, r) => s + (r.b12_returned || 0), 0)
    setTotals({ inGas, inB45, inB12, outGas, outB45d, outB45r, outB12d, outB12r })
    setLoading(false)
  }

  useEffect(() => { load() }, [])

  async function save() {
    if (!form.gas_kg && !form.b45_count && !form.b12_count) { setMsg('Nhập ít nhất 1 trường'); return }
    setMsg('')
    const data = {
      receive_date: form.receive_date,
      b45_count: parseInt(form.b45_count) || 0,
      b12_count: parseInt(form.b12_count) || 0,
      gas_kg: parseFloat(form.gas_kg) || 0,
      note: form.note,
    }
    if (editing) {
      await supabase.from('inventory_in').update(data).eq('id', editing)
      setEditing(null)
    } else {
      await supabase.from('inventory_in').insert(data)
    }
    setForm({ receive_date: new Date().toISOString().slice(0, 10), b45_count: '', b12_count: '', gas_kg: '', note: '' })
    setMsg(editing ? 'Đã cập nhật' : 'Đã thêm')
    await load()
  }

  async function del(id: string) {
    if (!confirm('Xóa bản ghi này?')) return
    await supabase.from('inventory_in').delete().eq('id', id)
    await load()
  }

  function edit(r: InRecord) {
    setEditing(r.id)
    setForm({ receive_date: r.receive_date, b45_count: String(r.b45_count || ''), b12_count: String(r.b12_count || ''), gas_kg: String(r.gas_kg || ''), note: r.note || '' })
  }

  const gasStock = totals.inGas - totals.outGas
  const b45Stock = totals.inB45 - totals.outB45d + totals.outB45r
  const b12Stock = totals.inB12 - totals.outB12d + totals.outB12r

  return (
    <div className="space-y-4">
      <h1 className="text-lg md:text-xl font-bold text-gray-800">Tồn kho</h1>

      {/* Summary cards */}
      <div className="grid grid-cols-3 gap-3">
        <div className="bg-white rounded-xl border p-4 text-center">
          <div className="text-xs text-gray-500">Gas tồn (kg)</div>
          <div className={`text-2xl font-bold ${gasStock >= 0 ? 'text-green-700' : 'text-red-600'}`}>
            {gasStock.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}
          </div>
          <div className="text-xs text-gray-400 mt-1">Nhập: {totals.inGas.toLocaleString('vi-VN')} · Xuất: {totals.outGas.toLocaleString('vi-VN')}</div>
        </div>
        <div className="bg-white rounded-xl border p-4 text-center">
          <div className="text-xs text-gray-500">Bình 45kg tồn</div>
          <div className={`text-2xl font-bold ${b45Stock >= 0 ? 'text-green-700' : 'text-red-600'}`}>{b45Stock}</div>
          <div className="text-xs text-gray-400 mt-1">Nhập: {totals.inB45} · Giao: {totals.outB45d} · Thu: {totals.outB45r}</div>
        </div>
        <div className="bg-white rounded-xl border p-4 text-center">
          <div className="text-xs text-gray-500">Bình 12kg tồn</div>
          <div className={`text-2xl font-bold ${b12Stock >= 0 ? 'text-green-700' : 'text-red-600'}`}>{b12Stock}</div>
          <div className="text-xs text-gray-400 mt-1">Nhập: {totals.inB12} · Giao: {totals.outB12d} · Thu: {totals.outB12r}</div>
        </div>
      </div>

      {/* Input form */}
      <div className="bg-white rounded-xl border p-4 space-y-3">
        <h2 className="font-semibold text-gray-700">{editing ? 'Sửa' : 'Nhập hàng từ Phát Vinh / LPG'}</h2>
        <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
          <div>
            <label className="block text-xs text-gray-500 mb-1">Ngày nhận</label>
            <input type="date" value={form.receive_date} onChange={e => setForm({ ...form, receive_date: e.target.value })}
              className="border rounded px-3 py-2 text-sm w-full" />
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Bình 45kg</label>
            <input type="number" value={form.b45_count} onChange={e => setForm({ ...form, b45_count: e.target.value })}
              placeholder="0" className="border rounded px-3 py-2 text-sm w-full" inputMode="numeric" />
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Bình 12kg</label>
            <input type="number" value={form.b12_count} onChange={e => setForm({ ...form, b12_count: e.target.value })}
              placeholder="0" className="border rounded px-3 py-2 text-sm w-full" inputMode="numeric" />
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Gas (kg)</label>
            <input type="number" step="0.1" value={form.gas_kg} onChange={e => setForm({ ...form, gas_kg: e.target.value })}
              placeholder="0" className="border rounded px-3 py-2 text-sm w-full" inputMode="decimal" />
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Ghi chú</label>
            <input value={form.note} onChange={e => setForm({ ...form, note: e.target.value })}
              className="border rounded px-3 py-2 text-sm w-full" />
          </div>
        </div>
        <div className="flex gap-2">
          <button onClick={save} className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700">
            {editing ? 'Cập nhật' : 'Thêm'}
          </button>
          {editing && <button onClick={() => { setEditing(null); setForm({ receive_date: new Date().toISOString().slice(0, 10), b45_count: '', b12_count: '', gas_kg: '', note: '' }) }}
            className="border px-4 py-2 rounded text-sm hover:bg-gray-50">Hủy</button>}
        </div>
        {msg && <p className="text-sm text-green-600">{msg}</p>}
      </div>

      {/* Records table */}
      <div className="bg-white rounded-xl border overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Ngày nhận', 'Bình 45', 'Bình 12', 'Gas (kg)', 'Ghi chú', ''].map(h => (
                <th key={h} className="text-left px-4 py-3 font-medium text-gray-600">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {records.map(r => (
              <tr key={r.id} className="border-b last:border-0 hover:bg-gray-50">
                <td className="px-4 py-2.5">{fmtDate(r.receive_date)}</td>
                <td className="px-4 py-2.5 text-right">{r.b45_count || ''}</td>
                <td className="px-4 py-2.5 text-right">{r.b12_count || ''}</td>
                <td className="px-4 py-2.5 text-right">{r.gas_kg || ''}</td>
                <td className="px-4 py-2.5 text-gray-500">{r.note}</td>
                <td className="px-4 py-2.5 flex gap-2 justify-end">
                  <button onClick={() => edit(r)} className="text-blue-600 hover:underline text-xs">Sửa</button>
                  <button onClick={() => del(r.id)} className="text-red-500 hover:underline text-xs">Xóa</button>
                </td>
              </tr>
            ))}
            {records.length === 0 && !loading && (
              <tr><td colSpan={6} className="px-4 py-8 text-center text-gray-400">Chưa có dữ liệu nhập hàng</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}
