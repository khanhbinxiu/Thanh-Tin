'use client'
import { useEffect, useState } from 'react'
import { supabase, Customer, Transaction } from '@/lib/supabase'
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

const COMPANY_NAME = 'CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ THÀNH TÍN LBG'
const COMPANY_ADDRESS = '115/22/60 BIS Đường Nguyễn Du, Phường 7, Quận Bình Thạnh, TP HCM'

function lastDay(y: number, m: number) { return new Date(y, m, 0).getDate() }

type SheetGroup = { sheetName: string; locationName: string; rows: Transaction[] }
type FileGroup = { fileName: string; customerCode: string; sheets: SheetGroup[] }

export default function ExportPage() {
  const [customers, setCustomers] = useState<Customer[]>([])
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [customFrom, setCustomFrom] = useState('')
  const [customTo, setCustomTo] = useState('')
  const [useCustom, setUseCustom] = useState(false)
  const [fileGroups, setFileGroups] = useState<FileGroup[]>([])
  const [loading, setLoading] = useState(false)
  const [exporting, setExporting] = useState<string | null>(null)

  useEffect(() => {
    supabase.from('customers').select('*').order('code').then(({ data }) => setCustomers(data || []))
  }, [])

  function getPeriod() {
    if (useCustom && customFrom && customTo) return { from: customFrom, to: customTo }
    return {
      from: `${year}-${String(month).padStart(2, '0')}-01`,
      to: `${year}-${String(month).padStart(2, '0')}-${lastDay(year, month)}`,
    }
  }

  async function loadPreview() {
    const { from, to } = getPeriod()
    setLoading(true)
    const { data } = await supabase.from('transactions')
      .select('*')
      .gte('delivery_date', from)
      .lte('delivery_date', to)
      .not('output_file_name', 'is', null)
      .order('delivery_date')
    const txs: Transaction[] = data || []

    // Group by output_file_name → output_sheet_name
    const fileMap = new Map<string, FileGroup>()
    for (const tx of txs) {
      const fname = tx.output_file_name || 'UNMAPPED'
      if (!fileMap.has(fname)) {
        fileMap.set(fname, { fileName: fname, customerCode: tx.customer_code, sheets: [] })
      }
      const fg = fileMap.get(fname)!
      let sg = fg.sheets.find(s => s.sheetName === tx.location)
      if (!sg) {
        sg = { sheetName: tx.location, locationName: tx.location, rows: [] }
        fg.sheets.push(sg)
      }
      sg.rows.push(tx)
    }
    setFileGroups([...fileMap.values()])
    setLoading(false)
  }

  async function exportFile(fg: FileGroup) {
    const { from, to } = getPeriod()
    const fromDate = new Date(from)
    const toDate = new Date(to)
    const customer = customers.find(c => c.code === fg.customerCode)
    setExporting(fg.fileName)

    const wb = new ExcelJS.Workbook()

    for (const sg of fg.sheets) {
      const ws = wb.addWorksheet(sg.sheetName.substring(0, 31))
      const totals = {
        b45d: sg.rows.reduce((s, r) => s + r.b45_delivered, 0),
        b45r: sg.rows.reduce((s, r) => s + r.b45_returned, 0),
        b12d: sg.rows.reduce((s, r) => s + r.b12_delivered, 0),
        b12r: sg.rows.reduce((s, r) => s + r.b12_returned, 0),
        gasD: sg.rows.reduce((s, r) => s + r.gas_delivered, 0),
        gasR: sg.rows.reduce((s, r) => s + r.gas_returned, 0),
        gasPaid: sg.rows.reduce((s, r) => s + r.gas_paid, 0),
        total: sg.rows.reduce((s, r) => s + r.total_amount, 0),
      }
      const vat = totals.total * 0.08

      // Row 1: Company header + Tháng
      ws.mergeCells('A1:L1')
      ws.getCell('A1').value = `${COMPANY_NAME}\n${COMPANY_ADDRESS}`
      ws.getCell('A1').font = { bold: true, size: 10 }
      ws.getCell('A1').alignment = { wrapText: true }
      ws.getRow(1).height = 30
      ws.getCell('M1').value = 'Tháng:'
      ws.getCell('N1').value = useCustom ? '' : month

      // Row 2: Title + Năm
      ws.mergeCells('A2:L2')
      ws.getCell('A2').value = 'BIÊN BẢN ĐỐI CHIẾU CÔNG NỢ'
      ws.getCell('A2').font = { bold: true, size: 13 }
      ws.getCell('A2').alignment = { horizontal: 'center' }
      ws.getCell('M2').value = 'Năm:'
      ws.getCell('N2').value = year

      // Row 3
      ws.mergeCells('A3:L3')
      ws.getCell('A3').value = useCustom
        ? `(${fromDate.getDate()}/${fromDate.getMonth()+1}/${fromDate.getFullYear()} - ${toDate.getDate()}/${toDate.getMonth()+1}/${toDate.getFullYear()})`
        : `(THÁNG ${String(month).padStart(2,'0')}/${year})`
      ws.getCell('A3').alignment = { horizontal: 'center' }

      // Row 5: Customer info
      ws.mergeCells('A5:N5')
      ws.getCell('A5').value = `KHÁCH HÀNG: ${customer?.name || fg.customerCode}\nĐỊA CHỈ: ${customer?.address || ''}\nMST: ${customer?.tax_code || ''}`
      ws.getCell('A5').font = { bold: true }
      ws.getCell('A5').alignment = { wrapText: true }
      ws.getRow(5).height = 45

      ws.mergeCells('A7:N7')
      ws.getCell('A7').value = 'Hai bên thống nhất số lượng bên B mua của bên A như sau:'

      ws.mergeCells('A8:N8')
      ws.getCell('A8').value = `1. Thời gian giao nhận: Từ ${fromDate.getDate()}/${fromDate.getMonth()+1}/${fromDate.getFullYear()} đến ${toDate.getDate()}/${toDate.getMonth()+1}/${toDate.getFullYear()}`

      ws.mergeCells('A9:N9')
      ws.getCell('A9').value = '2. Khối lượng hàng hóa theo bảng kê chi tiết:'

      // Table headers row 10 & 11
      const hdrs = ['STT','Ngày giao','Nội Dung','Bình Giao','Trả vỏ','Bình Giao','Trả vỏ','Gas giao','Gas trả','Gas thanh toán','Đơn giá chưa VAT (vnđ/kg)','Thành Tiền','Ghi chú']
      const units = ['','','','B45kg','B45kg','B12kg','B12kg','Kg','Kg','Kg','','','']
      ;[hdrs, units].forEach((arr, ri) => {
        arr.forEach((v, ci) => {
          const cell = ws.getRow(10 + ri).getCell(ci + 1)
          cell.value = v
          cell.font = { bold: ri === 0 }
          cell.alignment = { horizontal: 'center', wrapText: true }
          cell.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} }
        })
      })

      // Data rows from 12
      const firstDataRow = 12
      const yellowFill: ExcelJS.FillPattern = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }
      const thinBorder: Partial<ExcelJS.Borders> = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} }

      sg.rows.forEach((r, idx) => {
        const rowNum = firstDataRow + idx
        const dr = ws.getRow(rowNum)
        const noPrice = !r.unit_price || r.unit_price === 0

        const vals: (string|number|ExcelJS.CellFormulaValue)[] = [
          idx+1, r.delivery_date, sg.locationName,
          r.b45_delivered||0, r.b45_returned||0,
          r.b12_delivered||0, r.b12_returned||0,
          r.gas_delivered||0, r.gas_returned||0,
          { formula: `D${rowNum}*45+F${rowNum}*12-I${rowNum}` } as ExcelJS.CellFormulaValue,
          r.unit_price||'',
          noPrice ? '' : { formula: `J${rowNum}*K${rowNum}` } as ExcelJS.CellFormulaValue,
          r.note||''
        ]
        vals.forEach((v, ci) => {
          const cell = dr.getCell(ci + 1)
          cell.value = v as ExcelJS.CellValue
          cell.border = thinBorder
          if (ci >= 3) cell.alignment = { horizontal: 'right' }
          if (noPrice && (ci === 10 || ci === 11)) cell.fill = yellowFill
        })
      })

      // Totals with SUM formulas
      const lastDataRow = firstDataRow + sg.rows.length - 1
      const tRowNum = lastDataRow + 1
      const tRow = ws.getRow(tRowNum)
      const sumCols = ['D','E','F','G','H','I','J','L']
      const tVals: (string|ExcelJS.CellFormulaValue)[] = [
        '','TỔNG','',
        ...(['D','E','F','G','H','I','J'].map(c => ({ formula: `SUM(${c}${firstDataRow}:${c}${lastDataRow})` } as ExcelJS.CellFormulaValue))),
        '',
        { formula: `SUM(L${firstDataRow}:L${lastDataRow})` } as ExcelJS.CellFormulaValue,
        ''
      ]
      tVals.forEach((v, ci) => {
        const cell = tRow.getCell(ci+1)
        cell.value = v as ExcelJS.CellValue
        cell.font = { bold: true }
        cell.border = { top:{style:'thin'}, bottom:{style:'double'}, left:{style:'thin'}, right:{style:'thin'} }
        if (ci >= 3) cell.alignment = { horizontal: 'right' }
      })

      // Thuế, tổng tiền hàng with formulas
      const after = tRowNum + 1
      const lightYellow: ExcelJS.FillPattern = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFCC' } }

      ws.getRow(after).getCell(9).value = 'Thuế GTGT (8%)'
      ws.getRow(after).getCell(9).font = { bold: true }
      ws.getRow(after).getCell(9).fill = lightYellow
      ws.getRow(after).getCell(12).value = { formula: `L${tRowNum}*0.08` } as ExcelJS.CellValue
      ws.getRow(after).getCell(12).fill = lightYellow

      ws.getRow(after+1).getCell(9).value = 'Tổng tiền hàng (bao gồm GTGT 8%) là:'
      ws.getRow(after+1).getCell(9).font = { bold: true }
      ws.getRow(after+1).getCell(9).fill = lightYellow
      ws.getRow(after+1).getCell(12).value = { formula: `L${tRowNum}+L${after}` } as ExcelJS.CellValue
      ws.getRow(after+1).getCell(12).fill = lightYellow

      ws.getRow(after+2).getCell(9).value = `Tổng số tiền nợ tính từ ${fromDate.getDate()}/${fromDate.getMonth()+1}/${fromDate.getFullYear()} đến ${toDate.getDate()}/${toDate.getMonth()+1}/${toDate.getFullYear()}`
      ws.getRow(after+2).getCell(9).font = { bold: true }
      ws.getRow(after+2).getCell(9).fill = lightYellow
      ws.getRow(after+2).getCell(12).value = { formula: `L${after+1}` } as ExcelJS.CellValue
      ws.getRow(after+2).getCell(12).fill = lightYellow

      const sigRow = after + 5
      ws.getRow(sigRow).getCell(10).value = `TP Hồ Chí Minh, Ngày ${toDate.getDate()} tháng ${toDate.getMonth()+1} năm ${toDate.getFullYear()}`
      ws.getRow(sigRow+1).getCell(2).value = 'Xác nhận của khách hàng'
      ws.getRow(sigRow+1).getCell(10).value = 'Người lập'

      ws.columns = [
        {width:5},{width:14},{width:22},{width:8},{width:8},
        {width:8},{width:8},{width:10},{width:10},{width:16},
        {width:18},{width:16},{width:20},
      ]
    }

    const buf = await wb.xlsx.writeBuffer()
    const monthLabel = useCustom ? `${from}_${to}` : `T${String(month).padStart(2,'0')}`
    saveAs(new Blob([buf]), `ThanhTin_${monthLabel}_${fg.fileName}.xlsx`)
    setExporting(null)
  }

  async function exportAll() {
    for (const fg of fileGroups) {
      await exportFile(fg)
    }
  }

  const { from, to } = getPeriod()

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Xuất biên bản đối chiếu công nợ</h1>

      <div className="bg-white rounded-xl border p-5 space-y-4">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          {!useCustom && (
            <>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Tháng</label>
                <select value={month} onChange={e => setMonth(Number(e.target.value))}
                  className="border rounded px-3 py-2 text-sm w-full">
                  {Array.from({length:12},(_,i)=>i+1).map(m=><option key={m} value={m}>Tháng {m}</option>)}
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
                <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)} className="border rounded px-3 py-2 text-sm w-full"/>
              </div>
              <div>
                <label className="block text-xs text-gray-500 mb-1">Đến ngày</label>
                <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)} className="border rounded px-3 py-2 text-sm w-full"/>
              </div>
            </>
          )}
        </div>
        <label className="flex items-center gap-2 text-sm text-gray-600 cursor-pointer">
          <input type="checkbox" checked={useCustom} onChange={e=>setUseCustom(e.target.checked)}/>
          Tùy chỉnh kỳ
        </label>
        <div className="flex gap-3">
          <button onClick={loadPreview} disabled={loading}
            className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
            {loading ? 'Đang tải...' : 'Xem trước'}
          </button>
          {fileGroups.length > 0 && (
            <button onClick={exportAll} disabled={!!exporting}
              className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
              {exporting ? `Đang xuất ${exporting}...` : `Xuất tất cả (${fileGroups.length} file)`}
            </button>
          )}
        </div>
      </div>

      {fileGroups.length > 0 && (
        <div className="space-y-4">
          <p className="text-sm text-gray-600">Kỳ: <strong>{from}</strong> → <strong>{to}</strong> — {fileGroups.length} file</p>
          {fileGroups.map(fg => {
            const customer = customers.find(c => c.code === fg.customerCode)
            const totalRows = fg.sheets.reduce((s, sh) => s + sh.rows.length, 0)
            const totalAmt = fg.sheets.reduce((s, sh) => s + sh.rows.reduce((ss, r) => ss + r.total_amount, 0), 0)
            return (
              <div key={fg.fileName} className="bg-white rounded-xl border overflow-hidden">
                <div className="bg-gray-50 border-b px-4 py-3 flex items-center justify-between">
                  <div>
                    <span className="font-semibold text-gray-800">{fg.fileName}.xlsx</span>
                    <span className="ml-3 text-sm text-gray-500">{customer?.name || fg.customerCode}</span>
                    <span className="ml-2 text-xs text-gray-400">{fg.sheets.length} sheet · {totalRows} dòng</span>
                  </div>
                  <div className="flex items-center gap-4">
                    <span className="text-blue-600 font-medium text-sm">{(totalAmt * 1.08).toLocaleString('vi-VN')} đ</span>
                    <button onClick={() => exportFile(fg)} disabled={!!exporting}
                      className="bg-green-600 text-white px-3 py-1.5 rounded text-xs hover:bg-green-700 disabled:opacity-50">
                      {exporting === fg.fileName ? 'Đang xuất...' : 'Xuất Excel'}
                    </button>
                  </div>
                </div>
                {fg.sheets.map(sg => (
                  <div key={sg.sheetName} className="border-b last:border-0">
                    <div className="px-4 py-2 bg-blue-50 text-xs font-medium text-blue-700">
                      Sheet: {sg.sheetName} — {sg.locationName} ({sg.rows.length} dòng)
                    </div>
                    <table className="w-full text-xs">
                      <thead className="bg-gray-50">
                        <tr>
                          {['STT','Ngày','Nội Dung','B45↓','B45↑','B12↓','B12↑','Gas giao','Gas TT','Đơn giá','Thành tiền'].map(h=>(
                            <th key={h} className="text-center px-2 py-1.5 font-medium text-gray-600 border-r last:border-0">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sg.rows.map((r,i)=>(
                          <tr key={r.id} className="border-t hover:bg-gray-50">
                            <td className="px-2 py-1.5 text-center">{i+1}</td>
                            <td className="px-2 py-1.5 whitespace-nowrap">{r.delivery_date}</td>
                            <td className="px-2 py-1.5">{sg.locationName}</td>
                            <td className="px-2 py-1.5 text-right">{r.b45_delivered||''}</td>
                            <td className="px-2 py-1.5 text-right">{r.b45_returned||''}</td>
                            <td className="px-2 py-1.5 text-right">{r.b12_delivered||''}</td>
                            <td className="px-2 py-1.5 text-right">{r.b12_returned||''}</td>
                            <td className="px-2 py-1.5 text-right">{r.gas_delivered||''}</td>
                            <td className="px-2 py-1.5 text-right">{r.gas_paid||''}</td>
                            <td className="px-2 py-1.5 text-right">{r.unit_price?r.unit_price.toLocaleString('vi-VN'):''}</td>
                            <td className="px-2 py-1.5 text-right font-medium">{r.total_amount?r.total_amount.toLocaleString('vi-VN'):''}</td>
                          </tr>
                        ))}
                        <tr className="border-t-2 bg-gray-50 font-semibold">
                          <td colSpan={3} className="px-2 py-1.5 text-center">TỔNG</td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.b45_delivered,0)||''}</td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.b45_returned,0)||''}</td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.b12_delivered,0)||''}</td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.b12_returned,0)||''}</td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.gas_delivered,0)}</td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.gas_paid,0)}</td>
                          <td></td>
                          <td className="px-2 py-1.5 text-right">{sg.rows.reduce((s,r)=>s+r.total_amount,0).toLocaleString('vi-VN')}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                ))}
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
