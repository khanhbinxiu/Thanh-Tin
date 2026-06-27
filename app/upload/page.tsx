'use client'
import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'

function fmtDate(d: string) { const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}` }
import { supabase } from '@/lib/supabase'
import crypto from 'crypto'

type ParsedRow = {
  input_key: string
  delivery_date: string
  customer_code: string | null
  location: string
  output_file_name: string | null
  output_sheet_name: string | null
  pic: string
  b45_delivered: number
  b45_returned: number
  b12_delivered: number
  b12_returned: number
  gas_delivered: number
  gas_returned: number
  gas_paid: number
  unit_price: number
  total_amount: number
  note: string
  month: number
  year: number
  dedup_hash: string
  status?: 'new' | 'duplicate' | 'unmapped'
}

function n(v: unknown): number {
  if (v === '' || v === null || v === undefined) return 0
  const x = parseFloat(String(v).replace(/[^\d.-]/g, ''))
  return isNaN(x) ? 0 : x
}

function buildHash(row: ParsedRow) {
  const str = `${row.delivery_date}|${row.input_key}|${row.gas_delivered}|${row.total_amount}`
  return crypto.createHash('md5').update(str).digest('hex')
}

export default function UploadPage() {
  const [rows, setRows] = useState<ParsedRow[]>([])
  const [status, setStatus] = useState('')
  const [importing, setImporting] = useState(false)
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())

  const handleFile = useCallback(async (file: File) => {
    setStatus('Đang đọc file...')
    setRows([])
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })

    // Find INPUT_GIAO_HANG sheet or first sheet
    const sheetName = wb.SheetNames.find(s => s.includes('INPUT') || s.includes('GIAO')) || wb.SheetNames[0]
    const ws = wb.Sheets[sheetName]
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][]

    // Find header row
    let headerRow = -1
    for (let i = 0; i < Math.min(10, raw.length); i++) {
      const r = raw[i] as string[]
      if (r.some(c => String(c).includes('Ngày') || String(c).includes('Mã KH'))) {
        headerRow = i; break
      }
    }
    if (headerRow === -1) { setStatus('Không tìm thấy header. Kiểm tra lại file.'); return }

    const headers = (raw[headerRow] as string[]).map(h => String(h).trim())
    const col = (name: string) => headers.findIndex(h => h.includes(name))

    const iDay = col('Ngày')
    const iKey = col('Mã KH')
    const iPic = col('PIC')
    // B45 giao = first "Bình Giao", B45 trả = first "Trả vỏ", B12 giao = second "Bình Giao", B12 trả = second "Trả vỏ"
    const allBinhGiao = headers.reduce((acc, h, i) => h.includes('Bình Giao') ? [...acc, i] : acc, [] as number[])
    const allTraVo = headers.reduce((acc, h, i) => h.includes('Trả vỏ') ? [...acc, i] : acc, [] as number[])
    const iB45d = allBinhGiao[0] ?? -1
    const iB45r = allTraVo[0] ?? -1
    const iB12d = allBinhGiao[1] ?? -1
    const iB12r = allTraVo[1] ?? -1
    const iGasD = col('Gas giao')
    const iGasR = col('Gas trả')
    const iGasPaid = col('Gas thanh toán')
    const iPrice = col('Đơn giá')
    const iTotal = col('Thành Tiền')
    const iNote = col('Ghi chú')

    // Skip 2 header rows (header + unit row)
    const parsed: ParsedRow[] = []
    for (let i = headerRow + 2; i < raw.length; i++) {
      const r = raw[i] as unknown[]
      const dayVal = r[iDay]
      const inputKey = String(r[iKey] || '').trim().toLowerCase()
      if (!inputKey || !dayVal) continue
      const day = parseInt(String(dayVal))
      if (isNaN(day) || day < 1 || day > 31) continue

      const date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
      const row: ParsedRow = {
        input_key: inputKey,
        delivery_date: date,
        customer_code: null,
        location: inputKey,
        output_file_name: null,
        output_sheet_name: null,
        pic: String(r[iPic] || '').trim(),
        b45_delivered: n(r[iB45d]),
        b45_returned: n(r[iB45r]),
        b12_delivered: n(r[iB12d]),
        b12_returned: n(r[iB12r]),
        gas_delivered: n(r[iGasD]),
        gas_returned: n(r[iGasR]),
        gas_paid: n(r[iGasPaid]),
        unit_price: n(r[iPrice]),
        total_amount: n(r[iTotal]),
        note: String(r[iNote] || '').trim(),
        month,
        year,
        dedup_hash: '',
      }
      row.dedup_hash = buildHash(row)
      parsed.push(row)
    }

    if (parsed.length === 0) { setStatus('Không tìm thấy dữ liệu hợp lệ.'); return }

    // Load mappings
    setStatus('Đang tra cứu mapping...')
    const { data: mappings } = await supabase.from('location_mappings').select('*')
    const mappingMap = new Map((mappings || []).map(m => [m.input_key, m]))

    // Apply mappings
    for (const row of parsed) {
      const m = mappingMap.get(row.input_key)
      if (m) {
        row.customer_code = m.customer_code
        row.location = m.output_location_name
        row.output_file_name = m.output_file_name
        row.output_sheet_name = m.output_sheet_name
      }
    }

    // Check duplicates
    setStatus('Kiểm tra trùng lặp...')
    const hashes = parsed.map(r => r.dedup_hash)
    const { data: existing } = await supabase.from('transactions').select('dedup_hash').in('dedup_hash', hashes)
    const existingSet = new Set((existing || []).map(e => e.dedup_hash))

    const withStatus = parsed.map(r => ({
      ...r,
      status: existingSet.has(r.dedup_hash) ? 'duplicate' as const
            : !r.customer_code ? 'unmapped' as const
            : 'new' as const
    }))

    setRows(withStatus)
    const newCount = withStatus.filter(r => r.status === 'new').length
    const unmappedCount = withStatus.filter(r => r.status === 'unmapped').length
    const dupCount = withStatus.filter(r => r.status === 'duplicate').length
    setStatus(`${parsed.length} dòng: ${newCount} mới ✓ | ${dupCount} trùng | ${unmappedCount} chưa có mapping`)
  }, [month, year])

  async function doImport() {
    const newRows = rows.filter(r => r.status === 'new')
    if (newRows.length === 0) { setStatus('Không có dữ liệu mới để import'); return }
    setImporting(true)
    const toInsert = newRows.map(({ status: _s, ...r }) => r)
    const { error } = await supabase.from('transactions').insert(toInsert)
    if (error) {
      setStatus('Lỗi: ' + error.message)
    } else {
      setStatus(`Đã import ${newRows.length} dòng thành công!`)
      setRows(rows.map(r => r.status === 'new' ? { ...r, status: 'duplicate' as const } : r))
    }
    setImporting(false)
  }

  const unmappedRows = rows.filter(r => r.status === 'unmapped')
  const unmappedKeys = [...new Set(unmappedRows.map(r => r.input_key))]
  const unmappedCount = (key: string) => unmappedRows.filter(r => r.input_key === key).length

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Upload file Excel tổng</h1>

      <div className="bg-white rounded-xl border p-4 flex gap-4 items-end">
        <div>
          <label className="block text-xs text-gray-500 mb-1">Tháng *</label>
          <select value={month} onChange={e => setMonth(Number(e.target.value))}
            className="border rounded px-3 py-2 text-sm">
            {Array.from({length: 12}, (_, i) => i + 1).map(m =>
              <option key={m} value={m}>Tháng {m}</option>
            )}
          </select>
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Năm *</label>
          <input type="number" value={year} onChange={e => setYear(Number(e.target.value))}
            className="border rounded px-3 py-2 text-sm w-24" />
        </div>
      </div>

      <div
        className="bg-white rounded-xl border-2 border-dashed border-gray-300 p-10 text-center cursor-pointer hover:border-blue-400 transition-colors"
        onDragOver={e => e.preventDefault()}
        onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFile(f) }}
        onClick={() => document.getElementById('fileInput')?.click()}
      >
        <input id="fileInput" type="file" accept=".xlsx,.xls" className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f) }} />
        <p className="text-gray-500">Kéo thả file Excel vào đây hoặc <span className="text-blue-600 font-medium">bấm để chọn</span></p>
        <p className="text-xs text-gray-400 mt-1">Sheet INPUT_GIAO_HANG hoặc sheet đầu tiên</p>
      </div>

      {status && (
        <div className="bg-blue-50 border border-blue-200 rounded-lg px-4 py-3 text-sm text-blue-800">{status}</div>
      )}

      {unmappedKeys.length > 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-lg p-4">
          <p className="font-medium text-amber-800 mb-2">Các Mã KH chưa có mapping ({unmappedKeys.length}):</p>
          <div className="flex flex-wrap gap-2">
            {unmappedKeys.map(k => (
              <span key={k} className="bg-amber-100 text-amber-800 px-2 py-0.5 rounded text-xs font-mono flex items-center gap-1.5">
                {k}
                <span className="bg-amber-300 text-amber-900 rounded px-1">{unmappedCount(k)}</span>
              </span>
            ))}
          </div>
          <a href="/mappings" className="text-blue-600 text-sm mt-2 block hover:underline">→ Vào trang Mapping để thêm</a>
        </div>
      )}

      {rows.length > 0 && (
        <>
          <div className="flex items-center justify-between">
            <span className="text-sm text-gray-600">{rows.filter(r => r.status === 'new').length} dòng sẽ được import</span>
            <button onClick={doImport} disabled={importing || !rows.some(r => r.status === 'new')}
              className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
              {importing ? 'Đang import...' : 'Import dữ liệu mới'}
            </button>
          </div>

          <div className="bg-white rounded-xl border overflow-auto max-h-[60vh]">
            <table className="w-full text-xs">
              <thead className="bg-gray-50 border-b sticky top-0">
                <tr>
                  {['', 'Ngày', 'Mã KH (input)', 'KH map', 'Địa điểm output', 'B45↓', 'B45↑', 'B12↓', 'B12↑', 'Gas giao', 'Gas TT', 'Đơn giá', 'Thành tiền'].map(h => (
                    <th key={h} className="text-left px-3 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={i} className={`border-b last:border-0 ${
                    r.status === 'duplicate' ? 'bg-gray-50 text-gray-400' :
                    r.status === 'unmapped' ? 'bg-amber-50' : 'hover:bg-blue-50'
                  }`}>
                    <td className="px-3 py-2">
                      <span className={`text-xs px-1.5 py-0.5 rounded ${
                        r.status === 'new' ? 'bg-green-100 text-green-700' :
                        r.status === 'unmapped' ? 'bg-amber-100 text-amber-700' :
                        'bg-gray-200 text-gray-500'
                      }`}>
                        {r.status === 'new' ? 'Mới' : r.status === 'unmapped' ? 'Map?' : 'Trùng'}
                      </span>
                    </td>
                    <td className="px-3 py-2 whitespace-nowrap">{fmtDate(r.delivery_date)}</td>
                    <td className="px-3 py-2 font-mono">{r.input_key}</td>
                    <td className="px-3 py-2">{r.customer_code || <span className="text-amber-500">—</span>}</td>
                    <td className="px-3 py-2">{r.location !== r.input_key ? r.location : <span className="text-gray-400">—</span>}</td>
                    <td className="px-3 py-2 text-right">{r.b45_delivered || ''}</td>
                    <td className="px-3 py-2 text-right">{r.b45_returned || ''}</td>
                    <td className="px-3 py-2 text-right">{r.b12_delivered || ''}</td>
                    <td className="px-3 py-2 text-right">{r.b12_returned || ''}</td>
                    <td className="px-3 py-2 text-right">{r.gas_delivered || ''}</td>
                    <td className="px-3 py-2 text-right">{r.gas_paid || ''}</td>
                    <td className="px-3 py-2 text-right">{r.unit_price ? r.unit_price.toLocaleString('vi-VN') : ''}</td>
                    <td className="px-3 py-2 text-right">{r.total_amount ? r.total_amount.toLocaleString('vi-VN') : ''}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  )
}
