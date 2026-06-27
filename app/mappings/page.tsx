'use client'
import { useEffect, useState, useRef } from 'react'
import { supabase, Customer } from '@/lib/supabase'

type Mapping = {
  id: string
  input_key: string
  customer_code: string
  output_file_name: string
  output_location_name: string
}

const empty = { input_key: '', customer_code: '', output_file_name: '', output_location_name: '' }

export default function MappingsPage() {
  const [mappings, setMappings] = useState<Mapping[]>([])
  const [customers, setCustomers] = useState<Customer[]>([])
  const [form, setForm] = useState(empty)
  const [editing, setEditing] = useState<string | null>(null)
  const [msg, setMsg] = useState('')
  const [msgType, setMsgType] = useState<'error' | 'success'>('error')
  const [search, setSearch] = useState('')
  const inputKeyRef = useRef<HTMLInputElement>(null)

  async function load() {
    const [{ data: m }, { data: c }] = await Promise.all([
      supabase.from('location_mappings').select('*').order('input_key'),
      supabase.from('customers').select('*').order('code'),
    ])
    setMappings(m || [])
    setCustomers(c || [])
  }

  useEffect(() => { load() }, [])

  async function save() {
    if (!form.input_key || !form.customer_code || !form.output_file_name || !form.output_location_name) {
      setMsg('Cần điền đủ các trường bắt buộc')
      setMsgType('error')
      return
    }
    setMsg('')
    if (editing) {
      const { error } = await supabase.from('location_mappings').update({
        ...form,
        output_sheet_name: 'công nợ',
      }).eq('id', editing)
      if (error) { setMsg('Lỗi: ' + error.message); setMsgType('error'); return }
      setMsg('Đã cập nhật!')
      setMsgType('success')
    } else {
      const { error } = await supabase.from('location_mappings').insert({
        ...form,
        output_sheet_name: 'công nợ',
      })
      if (error) {
        setMsg(error.message.includes('duplicate') ? 'Input key đã tồn tại!' : 'Lỗi: ' + error.message)
        setMsgType('error')
        return
      }
      setMsg(`Đã thêm "${form.input_key}" → ${form.customer_code}`)
      setMsgType('success')
    }
    setForm(empty)
    setEditing(null)
    load()
    setTimeout(() => inputKeyRef.current?.focus(), 100)
  }

  async function del(id: string) {
    if (!confirm('Xóa mapping này?')) return
    await supabase.from('location_mappings').delete().eq('id', id)
    load()
  }

  function edit(m: Mapping) {
    setEditing(m.id)
    setForm({ input_key: m.input_key, customer_code: m.customer_code, output_file_name: m.output_file_name, output_location_name: m.output_location_name })
    setMsg('')
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (e.ctrlKey && e.key === 'Enter') {
      e.preventDefault()
      save()
    }
  }

  const filtered = mappings.filter(m =>
    !search || m.input_key.includes(search.toLowerCase()) || m.customer_code.includes(search.toLowerCase())
  )

  const existingFileNames = [...new Set(mappings.map(m => m.output_file_name))].sort()

  return (
    <div className="space-y-6" onKeyDown={handleKeyDown}>
      <h1 className="text-xl font-bold text-gray-800">Mapping Mã KH → Khách hàng</h1>
      <p className="text-sm text-gray-500">
        Mỗi giá trị trong cột &quot;Mã KH&quot; của file Excel input phải được map vào khách hàng và định dạng output tương ứng.
        <span className="ml-2 text-xs text-gray-400">Ctrl+Enter để lưu nhanh</span>
      </p>

      <div className="bg-white rounded-xl border p-5 space-y-3">
        <h2 className="font-semibold text-gray-700">{editing ? 'Sửa mapping' : 'Thêm mapping mới'}</h2>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-xs text-gray-500 mb-1">Mã KH trong input *</label>
            <input ref={inputKeyRef} disabled={!!editing} value={form.input_key}
              onChange={e => setForm({ ...form, input_key: e.target.value.toLowerCase() })}
              placeholder="vd: bùi văn ba"
              className="border rounded px-3 py-2 text-sm w-full font-mono disabled:bg-gray-100" />
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Khách hàng *</label>
            <select value={form.customer_code} onChange={e => setForm({ ...form, customer_code: e.target.value })}
              className="border rounded px-3 py-2 text-sm w-full">
              <option value="">-- Chọn KH --</option>
              {customers.map(c => <option key={c.code} value={c.code}>{c.code} – {c.name}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Tên file output *</label>
            <input list="file-names" value={form.output_file_name}
              onChange={e => setForm({ ...form, output_file_name: e.target.value })}
              placeholder="Chọn hoặc nhập mới"
              className="border rounded px-3 py-2 text-sm w-full font-mono" />
            <datalist id="file-names">
              {existingFileNames.map(f => <option key={f} value={f} />)}
            </datalist>
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Tên địa điểm (= tên sheet) *</label>
            <input value={form.output_location_name}
              onChange={e => setForm({ ...form, output_location_name: e.target.value })}
              placeholder="vd: BẾP BUIVANBA"
              className="border rounded px-3 py-2 text-sm w-full" />
          </div>
        </div>
        {msg && (
          <p className={`text-sm ${msgType === 'error' ? 'text-red-500' : 'text-green-600'}`}>{msg}</p>
        )}
        <div className="flex gap-2">
          <button onClick={save} className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700">
            {editing ? 'Cập nhật' : 'Thêm mới'}
          </button>
          {editing && <button onClick={() => { setEditing(null); setForm(empty); setMsg('') }}
            className="border px-4 py-2 rounded text-sm hover:bg-gray-50">Hủy</button>}
        </div>
      </div>

      <div className="bg-white rounded-xl border overflow-hidden">
        <div className="p-3 border-b">
          <input value={search} onChange={e => setSearch(e.target.value)}
            placeholder="Tìm theo input key hoặc mã KH..."
            className="border rounded px-3 py-2 text-sm w-72" />
          <span className="ml-3 text-sm text-gray-500">{filtered.length} mappings</span>
        </div>
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Input key (Mã KH)', 'Khách hàng', 'File output', 'Tên địa điểm (= tên sheet)', ''].map(h => (
                <th key={h} className="text-left px-4 py-3 font-medium text-gray-600">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map(m => (
              <tr key={m.id} className="border-b last:border-0 hover:bg-gray-50">
                <td className="px-4 py-2.5 font-mono text-xs">{m.input_key}</td>
                <td className="px-4 py-2.5 text-xs">{m.customer_code}</td>
                <td className="px-4 py-2.5 font-mono text-xs text-gray-500">{m.output_file_name}</td>
                <td className="px-4 py-2.5 text-xs">{m.output_location_name}</td>
                <td className="px-4 py-2.5 flex gap-2 justify-end">
                  <button onClick={() => edit(m)} className="text-blue-600 hover:underline text-xs">Sửa</button>
                  <button onClick={() => del(m.id)} className="text-red-500 hover:underline text-xs">Xóa</button>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={5} className="px-4 py-8 text-center text-gray-400">Chưa có mapping</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  )
}
