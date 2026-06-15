'use client'
import { useEffect, useState } from 'react'
import { supabase, Customer } from '@/lib/supabase'

const empty = { code: '', name: '', address: '', tax_code: '' }

export default function CustomersPage() {
  const [customers, setCustomers] = useState<Customer[]>([])
  const [form, setForm] = useState(empty)
  const [editing, setEditing] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [msg, setMsg] = useState('')

  async function load() {
    const { data } = await supabase.from('customers').select('*').order('code')
    setCustomers(data || [])
  }

  useEffect(() => { load() }, [])

  async function save() {
    if (!form.code || !form.name) { setMsg('Cần nhập Mã KH và Tên KH'); return }
    setLoading(true)
    setMsg('')
    if (editing) {
      await supabase.from('customers').update({ name: form.name, address: form.address, tax_code: form.tax_code }).eq('id', editing)
    } else {
      const { error } = await supabase.from('customers').insert(form)
      if (error) { setMsg('Mã KH đã tồn tại'); setLoading(false); return }
    }
    setForm(empty)
    setEditing(null)
    setLoading(false)
    load()
  }

  async function del(id: string) {
    if (!confirm('Xóa khách hàng này?')) return
    await supabase.from('customers').delete().eq('id', id)
    load()
  }

  function edit(c: Customer) {
    setEditing(c.id)
    setForm({ code: c.code, name: c.name, address: c.address || '', tax_code: c.tax_code || '' })
  }

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Danh mục khách hàng</h1>

      <div className="bg-white rounded-xl border p-5 space-y-3">
        <h2 className="font-semibold text-gray-700">{editing ? 'Sửa khách hàng' : 'Thêm khách hàng'}</h2>
        <div className="grid grid-cols-2 gap-3">
          <input disabled={!!editing} placeholder="Mã KH *" value={form.code}
            onChange={e => setForm({ ...form, code: e.target.value.toUpperCase() })}
            className="border rounded px-3 py-2 text-sm disabled:bg-gray-100" />
          <input placeholder="Tên KH *" value={form.name}
            onChange={e => setForm({ ...form, name: e.target.value })}
            className="border rounded px-3 py-2 text-sm" />
          <input placeholder="Địa chỉ" value={form.address}
            onChange={e => setForm({ ...form, address: e.target.value })}
            className="border rounded px-3 py-2 text-sm col-span-2" />
          <input placeholder="Mã số thuế" value={form.tax_code}
            onChange={e => setForm({ ...form, tax_code: e.target.value })}
            className="border rounded px-3 py-2 text-sm" />
        </div>
        {msg && <p className="text-red-500 text-sm">{msg}</p>}
        <div className="flex gap-2">
          <button onClick={save} disabled={loading}
            className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
            {loading ? 'Đang lưu...' : editing ? 'Cập nhật' : 'Thêm mới'}
          </button>
          {editing && (
            <button onClick={() => { setEditing(null); setForm(empty) }}
              className="border px-4 py-2 rounded text-sm hover:bg-gray-50">
              Hủy
            </button>
          )}
        </div>
      </div>

      <div className="bg-white rounded-xl border overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 border-b">
            <tr>
              {['Mã KH', 'Tên khách hàng', 'Địa chỉ', 'MST', ''].map(h => (
                <th key={h} className="text-left px-4 py-3 font-medium text-gray-600">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {customers.map(c => (
              <tr key={c.id} className="border-b last:border-0 hover:bg-gray-50">
                <td className="px-4 py-3 font-mono font-medium">{c.code}</td>
                <td className="px-4 py-3">{c.name}</td>
                <td className="px-4 py-3 text-gray-500">{c.address}</td>
                <td className="px-4 py-3 text-gray-500">{c.tax_code}</td>
                <td className="px-4 py-3 flex gap-2 justify-end">
                  <button onClick={() => edit(c)} className="text-blue-600 hover:underline text-xs">Sửa</button>
                  <button onClick={() => del(c.id)} className="text-red-500 hover:underline text-xs">Xóa</button>
                </td>
              </tr>
            ))}
            {customers.length === 0 && (
              <tr><td colSpan={5} className="px-4 py-8 text-center text-gray-400">Chưa có khách hàng nào</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}
