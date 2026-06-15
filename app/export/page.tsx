'use client'
import { useEffect, useState } from 'react'
import { supabase, Customer, Transaction } from '@/lib/supabase'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

const COMPANY_NAME = 'CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ THÀNH TÍN LBG'
const COMPANY_ADDRESS = '115/22/60 BIS Đường Nguyễn Du, Phường 7, Quận Bình Thạnh, TP HCM'

function lastDayOfMonth(y: number, m: number) {
  return new Date(y, m, 0).getDate()
}

export default function ExportPage() {
  const [customers, setCustomers] = useState<Customer[]>([])
  const [selectedCode, setSelectedCode] = useState('')
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [customFrom, setCustomFrom] = useState('')
  const [customTo, setCustomTo] = useState('')
  const [useCustom, setUseCustom] = useState(false)
  const [preview, setPreview] = useState<{ location: string; rows: Transaction[] }[]>([])
  const [loading, setLoading] = useState(false)
  const [exporting, setExporting] = useState(false)

  useEffect(() => {
    supabase.from('customers').select('*').order('code').then(({ data }) => setCustomers(data || []))
  }, [])

  function getPeriod() {
    if (useCustom && customFrom && customTo) return { from: customFrom, to: customTo }
    const from = `${year}-${String(month).padStart(2, '0')}-01`
    const to = `${year}-${String(month).padStart(2, '0')}-${lastDayOfMonth(year, month)}`
    return { from, to }
  }

  async function loadPreview() {
    if (!selectedCode) return
    const { from, to } = getPeriod()
    setLoading(true)
    const { data } = await supabase.from('transactions')
      .select('*')
      .eq('customer_code', selectedCode)
      .gte('delivery_date', from)
      .lte('delivery_date', to)
      .order('delivery_date')
    const txs = data || []
    const locations = [...new Set(txs.map(t => t.location))]
    setPreview(locations.map(loc => ({ location: loc, rows: txs.filter(t => t.location === loc) })))
    setLoading(false)
  }

  async function exportExcel() {
    if (!selectedCode || preview.length === 0) return
    const customer = customers.find(c => c.code === selectedCode)!
    const { from, to } = getPeriod()
    const fromDate = new Date(from)
    const toDate = new Date(to)
    const monthLabel = useCustom ? `${from}_${to}` : `T${String(month).padStart(2, '0')}`

    setExporting(true)
    const wb = new ExcelJS.Workbook()

    for (const group of preview) {
      const ws = wb.addWorksheet(group.location.substring(0, 31))
      const totals = {
        b45d: group.rows.reduce((s, r) => s + r.b45_delivered, 0),
        b45r: group.rows.reduce((s, r) => s + r.b45_returned, 0),
        b12d: group.rows.reduce((s, r) => s + r.b12_delivered, 0),
        b12r: group.rows.reduce((s, r) => s + r.b12_returned, 0),
        gasD: group.rows.reduce((s, r) => s + r.gas_delivered, 0),
        gasR: group.rows.reduce((s, r) => s + r.gas_returned, 0),
        gasPaid: group.rows.reduce((s, r) => s + r.gas_paid, 0),
        total: group.rows.reduce((s, r) => s + r.total_amount, 0),
        vat: group.rows.reduce((s, r) => s + r.total_amount * 0.08, 0),
      }

      // Header
      ws.mergeCells('A1:L1')
      ws.getCell('A1').value = COMPANY_NAME
      ws.getCell('A1').font = { bold: true, size: 11 }

      ws.mergeCells('A2:L2')
      ws.getCell('A2').value = COMPANY_ADDRESS
      ws.getCell('A2').font = { size: 10 }

      ws.getCell('M1').value = 'Tháng:'
      ws.getCell('N1').value = month
      ws.getCell('M2').value = 'Năm:'
      ws.getCell('N2').value = year

      ws.mergeCells('A4:L4')
      ws.getCell('A4').value = 'BIÊN BẢN ĐỐI CHIẾU CÔNG NỢ'
      ws.getCell('A4').font = { bold: true, size: 13 }
      ws.getCell('A4').alignment = { horizontal: 'center' }

      ws.mergeCells('A5:L5')
      ws.getCell('A5').value = `(THÁNG ${String(month).padStart(2, '0')}/${year})`
      ws.getCell('A5').alignment = { horizontal: 'center' }

      ws.mergeCells('A7:L7')
      ws.getCell('A7').value = `KHÁCH HÀNG: ${customer.name}`
      ws.getCell('A7').font = { bold: true }

      ws.mergeCells('A8:L8')
      ws.getCell('A8').value = `ĐỊA CHỈ: ${customer.address || ''}`

      ws.mergeCells('A9:L9')
      ws.getCell('A9').value = `MST: ${customer.tax_code || ''}`

      ws.mergeCells('A11:L11')
      ws.getCell('A11').value = 'Hai bên thống nhất số lượng bên B mua của bên A như sau:'

      ws.mergeCells('A12:L12')
      ws.getCell('A12').value = `1. Thời gian giao nhận: Từ ${fromDate.getDate()}/${fromDate.getMonth() + 1}/${fromDate.getFullYear()} đến ${toDate.getDate()}/${toDate.getMonth() + 1}/${toDate.getFullYear()}`

      ws.mergeCells('A13:L13')
      ws.getCell('A13').value = '2. Khối lượng hàng hóa theo bảng kê chi tiết:'

      // Table header row 1
      const hdr1 = ws.getRow(14)
      ;['STT', 'Ngày giao', 'Nội Dung', 'Bình Giao', 'Trả vỏ', 'Bình Giao', 'Trả vỏ', 'Gas giao', 'Gas trả', 'Gas thanh toán', 'Đơn giá chưa VAT (vnđ/kg)', 'Thành Tiền', 'Ghi chú'].forEach((v, i) => {
        hdr1.getCell(i + 1).value = v
        hdr1.getCell(i + 1).font = { bold: true }
        hdr1.getCell(i + 1).alignment = { horizontal: 'center', wrapText: true }
        hdr1.getCell(i + 1).border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
      })

      // Table header row 2 (units)
      const hdr2 = ws.getRow(15)
      ;['', '', '', 'B45kg', 'B45kg', 'B12kg', 'B12kg', 'Kg', 'Kg', 'Kg', '', '', ''].forEach((v, i) => {
        hdr2.getCell(i + 1).value = v
        hdr2.getCell(i + 1).alignment = { horizontal: 'center' }
        hdr2.getCell(i + 1).border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
      })

      // Data rows
      group.rows.forEach((r, idx) => {
        const row = ws.getRow(16 + idx)
        const vals = [
          idx + 1,
          r.delivery_date,
          r.location,
          r.b45_delivered || '',
          r.b45_returned || '',
          r.b12_delivered || '',
          r.b12_returned || '',
          r.gas_delivered || '',
          r.gas_returned || '',
          r.gas_paid || '',
          r.unit_price || '',
          r.total_amount || '',
          r.note || '',
        ]
        vals.forEach((v, i) => {
          row.getCell(i + 1).value = v as ExcelJS.CellValue
          row.getCell(i + 1).border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
          if (i >= 3) row.getCell(i + 1).alignment = { horizontal: 'right' }
        })
      })

      // Totals row
      const totRow = ws.getRow(16 + group.rows.length)
      const totVals = ['', 'TỔNG', '', totals.b45d, totals.b45r, totals.b12d, totals.b12r, totals.gasD, totals.gasR, totals.gasPaid, '', totals.total, '']
      totVals.forEach((v, i) => {
        totRow.getCell(i + 1).value = v as ExcelJS.CellValue
        totRow.getCell(i + 1).font = { bold: true }
        totRow.getCell(i + 1).border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } }
        if (i >= 3) totRow.getCell(i + 1).alignment = { horizontal: 'right' }
      })

      const afterTotal = 17 + group.rows.length
      ws.getRow(afterTotal).getCell(10).value = 'Thuế GTGT (8%)'
      ws.getRow(afterTotal).getCell(12).value = totals.vat
      ws.getRow(afterTotal + 1).getCell(10).value = 'Tổng tiền hàng (bao gồm GTGT 8%) là:'
      ws.getRow(afterTotal + 1).getCell(12).value = totals.total + totals.vat
      ws.getRow(afterTotal + 2).getCell(10).value = `Tổng số tiền nợ tính từ ${fromDate.getDate()}/${fromDate.getMonth() + 1}/${fromDate.getFullYear()} đến ${toDate.getDate()}/${toDate.getMonth() + 1}/${toDate.getFullYear()}`
      ws.getRow(afterTotal + 2).getCell(12).value = totals.total + totals.vat

      // Signatures
      const sigRow = afterTotal + 5
      ws.getRow(sigRow).getCell(10).value = `TP Hồ Chí Minh, Ngày ${toDate.getDate()} tháng ${toDate.getMonth() + 1} năm ${toDate.getFullYear()}`
      ws.getRow(sigRow + 1).getCell(1).value = 'Xác nhận của khách hàng'
      ws.getRow(sigRow + 1).getCell(10).value = 'Người lập'

      // Column widths
      ws.columns = [
        { width: 5 }, { width: 14 }, { width: 20 }, { width: 8 }, { width: 8 },
        { width: 8 }, { width: 8 }, { width: 10 }, { width: 10 }, { width: 16 },
        { width: 18 }, { width: 16 }, { width: 20 },
      ]
    }

    const buf = await wb.xlsx.writeBuffer()
    const customerName = customer.name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 40)
    saveAs(new Blob([buf]), `ThanhTin_${monthLabel}_${customerName}.xlsx`)
    setExporting(false)
  }

  const customer = customers.find(c => c.code === selectedCode)
  const { from, to } = getPeriod()

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Xuất biên bản đối chiếu công nợ</h1>

      <div className="bg-white rounded-xl border p-5 space-y-4">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div>
            <label className="block text-xs text-gray-500 mb-1">Khách hàng *</label>
            <select value={selectedCode} onChange={e => setSelectedCode(e.target.value)}
              className="border rounded px-3 py-2 text-sm w-full">
              <option value="">-- Chọn KH --</option>
              {customers.map(c => <option key={c.code} value={c.code}>{c.code} – {c.name}</option>)}
            </select>
          </div>

          {!useCustom && (
            <>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Tháng</label>
                <select value={month} onChange={e => setMonth(Number(e.target.value))}
                  className="border rounded px-3 py-2 text-sm w-full">
                  {Array.from({ length: 12 }, (_, i) => i + 1).map(m => (
                    <option key={m} value={m}>Tháng {m}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Năm</label>
                <input type="number" value={year} onChange={e => setYear(Number(e.target.value))}
                  className="border rounded px-3 py-2 text-sm w-full" />
              </div>
            </>
          )}

          {useCustom && (
            <>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Từ ngày</label>
                <input type="date" value={customFrom} onChange={e => setCustomFrom(e.target.value)}
                  className="border rounded px-3 py-2 text-sm w-full" />
              </div>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Đến ngày</label>
                <input type="date" value={customTo} onChange={e => setCustomTo(e.target.value)}
                  className="border rounded px-3 py-2 text-sm w-full" />
              </div>
            </>
          )}
        </div>

        <label className="flex items-center gap-2 text-sm text-gray-600 cursor-pointer">
          <input type="checkbox" checked={useCustom} onChange={e => setUseCustom(e.target.checked)} />
          Tùy chỉnh kỳ (không theo tháng chuẩn)
        </label>

        <div className="flex gap-3">
          <button onClick={loadPreview} disabled={!selectedCode || loading}
            className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
            {loading ? 'Đang tải...' : 'Xem trước'}
          </button>
          <button onClick={exportExcel} disabled={!selectedCode || preview.length === 0 || exporting}
            className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
            {exporting ? 'Đang xuất...' : 'Xuất Excel'}
          </button>
        </div>
      </div>

      {preview.length > 0 && customer && (
        <div className="space-y-6">
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-sm text-yellow-800">
            <strong>{customer.name}</strong> — Kỳ: {from} đến {to} — {preview.length} địa điểm
          </div>

          {preview.map(group => {
            const totalAmount = group.rows.reduce((s, r) => s + r.total_amount, 0)
            const vat = totalAmount * 0.08
            return (
              <div key={group.location} className="bg-white rounded-xl border overflow-hidden">
                <div className="bg-gray-50 border-b px-4 py-3 font-semibold text-gray-700 flex justify-between">
                  <span>{group.location}</span>
                  <span className="text-blue-600">{(totalAmount + vat).toLocaleString('vi-VN')} đ (incl. VAT)</span>
                </div>
                <table className="w-full text-xs">
                  <thead className="bg-gray-50 border-b">
                    <tr>
                      {['STT', 'Ngày giao', 'Nội Dung', 'B45↓', 'B45↑', 'B12↓', 'B12↑', 'Gas giao', 'Gas trả', 'Gas TT', 'Đơn giá', 'Thành tiền', 'Ghi chú'].map(h => (
                        <th key={h} className="text-center px-2 py-2 font-medium text-gray-600 border-r last:border-0">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {group.rows.map((r, i) => (
                      <tr key={r.id} className="border-b last:border-0 hover:bg-gray-50">
                        <td className="px-2 py-1.5 text-center">{i + 1}</td>
                        <td className="px-2 py-1.5 whitespace-nowrap">{r.delivery_date}</td>
                        <td className="px-2 py-1.5">{r.location}</td>
                        <td className="px-2 py-1.5 text-right">{r.b45_delivered || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.b45_returned || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.b12_delivered || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.b12_returned || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.gas_delivered || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.gas_returned || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.gas_paid || ''}</td>
                        <td className="px-2 py-1.5 text-right">{r.unit_price ? r.unit_price.toLocaleString('vi-VN') : ''}</td>
                        <td className="px-2 py-1.5 text-right font-medium">{r.total_amount ? r.total_amount.toLocaleString('vi-VN') : ''}</td>
                        <td className="px-2 py-1.5 text-gray-500">{r.note}</td>
                      </tr>
                    ))}
                    <tr className="bg-gray-50 font-semibold border-t-2">
                      <td colSpan={3} className="px-2 py-2 text-center">TỔNG</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.b45_delivered, 0) || ''}</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.b45_returned, 0) || ''}</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.b12_delivered, 0) || ''}</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.b12_returned, 0) || ''}</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.gas_delivered, 0)}</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.gas_returned, 0) || ''}</td>
                      <td className="px-2 py-2 text-right">{group.rows.reduce((s, r) => s + r.gas_paid, 0)}</td>
                      <td colSpan={3} className="px-2 py-2 text-right">{totalAmount.toLocaleString('vi-VN')}</td>
                    </tr>
                  </tbody>
                </table>
                <div className="px-4 py-3 text-xs text-right space-y-1 border-t">
                  <div>Thuế GTGT (8%): <span className="font-medium">{vat.toLocaleString('vi-VN')} đ</span></div>
                  <div className="font-semibold text-sm">Tổng cộng (incl. VAT): {(totalAmount + vat).toLocaleString('vi-VN')} đ</div>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
