'use client'
import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import { parseExcelDate } from '@/lib/utils'
import crypto from 'crypto'

type ParsedRow = {
  delivery_date: string
  customer_code: string
  location: string
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
  dedup_hash: string
  status?: 'new' | 'duplicate'
}

function buildHash(row: ParsedRow) {
  const str = `${row.delivery_date}|${row.customer_code}|${row.location}|${row.gas_delivered}|${row.total_amount}`
  return crypto.createHash('md5').update(str).digest('hex')
}

function n(v: unknown): number {
  const x = parseFloat(String(v).replace(/[^\d.-]/g, ''))
  return isNaN(x) ? 0 : x
}

export default function UploadPage() {
  const [rows, setRows] = useState<ParsedRow[]>([])
  const [status, setStatus] = useState('')
  const [importing, setImporting] = useState(false)

  const handleFile = useCallback(async (file: File) => {
    setStatus('Đang đọc file...')
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array', cellDates: true })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][]

    // Find header row (contains "Ngày giao" or "Mã KH")
    let headerRow = -1
    for (let i = 0; i < Math.min(10, raw.length); i++) {
      const row = raw[i] as string[]
      if (row.some(c => String(c).includes('Ngày') || String(c).includes('Mã KH'))) {
        headerRow = i
        break
      }
    }
    if (headerRow === -1) { setStatus('Không tìm thấy header. Kiểm tra lại file.'); return }

    const headers = (raw[headerRow] as string[]).map(h => String(h).trim())
    const colIdx = (name: string) => headers.findIndex(h => h.includes(name))

    const iDate = colIdx('Ngày')
    const iCode = colIdx('Mã KH')
    const iLoc = colIdx('Nội Dung')
    const iPic = colIdx('PIC')
    const iB45d = headers.findIndex((h, i) => h.includes('Bình Giao') && i < headers.findIndex((hh, ii) => hh.includes('Bình Giao') && ii > 0))
    const iB45r = headers.findIndex((h, i) => h.includes('Trả vỏ') && i < headers.findIndex((hh, ii) => hh.includes('Trả vỏ') && ii > 0))

    // Find columns by position pattern: B45 giao, B45 trả, B12 giao, B12 trả
    const b45dIdx = headers.findIndex(h => h.includes('Bình Giao'))
    const b45rIdx = headers.indexOf('Trả vỏ', b45dIdx)
    const b12dIdx = headers.indexOf('Bình Giao', b45dIdx + 1)
    const b12rIdx = headers.indexOf('Trả vỏ', b45rIdx + 1)
    const iGasD = colIdx('Gas giao')
    const iGasR = colIdx('Gas trả')
    const iGasPaid = colIdx('Gas thanh toán')
    const iPrice = colIdx('Đơn giá')
    const iTotal = colIdx('Thành Tiền')
    const iNote = colIdx('Ghi chú')

    const parsed: ParsedRow[] = []
    for (let i = headerRow + 2; i < raw.length; i++) {
      const r = raw[i] as unknown[]
      const dateVal = r[iDate]
      const code = String(r[iCode] || '').trim()
      if (!code || !dateVal) continue
      const dateStr = parseExcelDate(dateVal)
      if (!dateStr) continue

      const row: ParsedRow = {
        delivery_date: dateStr,
        customer_code: code,
        location: String(r[iLoc] || '').trim(),
        pic: String(r[iPic] || '').trim(),
        b45_delivered: n(r[b45dIdx]),
        b45_returned: n(r[b45rIdx]),
        b12_delivered: n(r[b12dIdx]),
        b12_returned: n(r[b12rIdx]),
        gas_delivered: n(r[iGasD]),
        gas_returned: n(r[iGasR]),
        gas_paid: n(r[iGasPaid]),
        unit_price: n(r[iPrice]),
        total_amount: n(r[iTotal]),
        note: String(r[iNote] || '').trim(),
        dedup_hash: '',
      }
      row.dedup_hash = buildHash(row)
      parsed.push(row)
    }

    if (parsed.length === 0) { setStatus('Không tìm thấy dữ liệu hợp lệ trong file.'); return }

    // Check duplicates
    setStatus('Kiểm tra trùng lặp...')
    const hashes = parsed.map(r => r.dedup_hash)
    const { data: existing } = await supabase.from('transactions').select('dedup_hash').in('dedup_hash', hashes)
    const existingSet = new Set((existing || []).map(e => e.dedup_hash))
    const withStatus = parsed.map(r => ({ ...r, status: existingSet.has(r.dedup_hash) ? 'duplicate' as const : 'new' as const }))
    setRows(withStatus)
    const newCount = withStatus.filter(r => r.status === 'new').length
    setStatus(`Tìm thấy ${parsed.length} dòng: ${newCount} mới, ${parsed.length - newCount} đã tồn tại`)
  }, [])

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
      setRows(rows.map(r => ({ ...r, status: 'duplicate' as const })))
    }
    setImporting(false)
  }

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Upload file Excel tổng</h1>

      <div
        className="bg-white rounded-xl border-2 border-dashed border-gray-300 p-10 text-center cursor-pointer hover:border-blue-400 transition-colors"
        onDragOver={e => e.preventDefault()}
        onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFile(f) }}
        onClick={() => document.getElementById('fileInput')?.click()}
      >
        <input id="fileInput" type="file" accept=".xlsx,.xls" className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f) }} />
        <p className="text-gray-500">Kéo thả file Excel vào đây hoặc <span className="text-blue-600 font-medium">bấm để chọn file</span></p>
        <p className="text-xs text-gray-400 mt-1">Hỗ trợ .xlsx, .xls</p>
      </div>

      {status && (
        <div className="bg-blue-50 border border-blue-200 rounded-lg px-4 py-3 text-sm text-blue-800">{status}</div>
      )}

      {rows.length > 0 && (
        <>
          <div className="flex items-center justify-between">
            <span className="text-sm text-gray-600">{rows.filter(r => r.status === 'new').length} dòng sẽ được import</span>
            <button onClick={doImport} disabled={importing || rows.every(r => r.status === 'duplicate')}
              className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
              {importing ? 'Đang import...' : 'Import dữ liệu mới'}
            </button>
          </div>

          <div className="bg-white rounded-xl border overflow-auto">
            <table className="w-full text-xs">
              <thead className="bg-gray-50 border-b sticky top-0">
                <tr>
                  {['', 'Ngày', 'Mã KH', 'Địa điểm', 'B45 Giao', 'B45 Trả', 'B12 Giao', 'B12 Trả', 'Gas giao', 'Gas TT', 'Đơn giá', 'Thành tiền'].map(h => (
                    <th key={h} className="text-left px-3 py-2 font-medium text-gray-600 whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={i} className={`border-b last:border-0 ${r.status === 'duplicate' ? 'bg-gray-50 text-gray-400' : 'hover:bg-blue-50'}`}>
                    <td className="px-3 py-2">
                      <span className={`text-xs px-1.5 py-0.5 rounded ${r.status === 'new' ? 'bg-green-100 text-green-700' : 'bg-gray-200 text-gray-500'}`}>
                        {r.status === 'new' ? 'Mới' : 'Trùng'}
                      </span>
                    </td>
                    <td className="px-3 py-2 whitespace-nowrap">{r.delivery_date}</td>
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
