'use client'
import { useEffect, useState } from 'react'
import { supabase, Customer } from '@/lib/supabase'
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx'
import { saveAs } from 'file-saver'

function formatPrice(n: number) { return n.toLocaleString('vi-VN') }
async function loadImage(url: string): Promise<ArrayBuffer> { return (await fetch(url)).arrayBuffer() }

type PriceRow = { customer_code: string; customer_name: string; location: string; output_file_name: string; unit_price: number; count: number }

export default function PricesPage() {
  const [tab, setTab] = useState<'prices' | 'quote'>('prices')
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [customers, setCustomers] = useState<Customer[]>([])

  // Prices state
  const [rows, setRows] = useState<PriceRow[]>([])
  const [loading, setLoading] = useState(false)
  const [saving, setSaving] = useState(false)
  const [msg, setMsg] = useState('')
  const [editPrices, setEditPrices] = useState<Record<string, string>>({})

  // Quote state
  const [customerCode, setCustomerCode] = useState('')
  const [quotePrice, setQuotePrice] = useState('')
  const [generating, setGenerating] = useState(false)

  useEffect(() => {
    supabase.from('customers').select('*').order('code').then(({ data }) => setCustomers(data || []))
  }, [])

  const monthStr = String(month).padStart(2, '0')
  const lastDay = new Date(year, month, 0).getDate()
  const from = `${year}-${monthStr}-01`
  const to = `${year}-${monthStr}-${lastDay}`

  async function loadPrices() {
    setLoading(true); setMsg('')
    const { data: txs } = await supabase.from('transactions')
      .select('customer_code, location, output_file_name, unit_price')
      .gte('delivery_date', from).lte('delivery_date', to).order('customer_code')
    const custMap = new Map(customers.map(c => [c.code, c.name]))
    const grouped = new Map<string, PriceRow>()
    for (const tx of (txs || [])) {
      const key = `${tx.customer_code}||${tx.location}`
      if (!grouped.has(key)) grouped.set(key, { customer_code: tx.customer_code, customer_name: custMap.get(tx.customer_code) || tx.customer_code, location: tx.location, output_file_name: tx.output_file_name || '', unit_price: tx.unit_price || 0, count: 0 })
      const row = grouped.get(key)!
      row.count++
      if (tx.unit_price && !row.unit_price) row.unit_price = tx.unit_price
    }
    const result = [...grouped.values()].sort((a, b) => a.output_file_name.localeCompare(b.output_file_name) || a.location.localeCompare(b.location))
    setRows(result)
    const prices: Record<string, string> = {}
    for (const r of result) prices[`${r.customer_code}||${r.location}`] = r.unit_price ? String(r.unit_price) : ''
    setEditPrices(prices)
    setLoading(false)
  }

  async function savePrices() {
    setSaving(true); setMsg('')
    let updated = 0
    for (const row of rows) {
      const key = `${row.customer_code}||${row.location}`
      const newPrice = parseFloat(editPrices[key] || '0')
      if (newPrice !== row.unit_price) {
        const { data: txs } = await supabase.from('transactions')
          .select('id, b45_delivered, b12_delivered, gas_returned')
          .eq('customer_code', row.customer_code).eq('location', row.location)
          .gte('delivery_date', from).lte('delivery_date', to)
        for (const tx of (txs || [])) {
          const gp = tx.b45_delivered * 45 + tx.b12_delivered * 12 - (tx.gas_returned || 0)
          await supabase.from('transactions').update({ unit_price: newPrice, total_amount: gp * newPrice }).eq('id', tx.id)
        }
        updated++
      }
    }
    setMsg(`Đã cập nhật ${updated} nhóm giá!`)
    await loadPrices()
    setSaving(false)
  }

  function applyToFile(fileName: string, price: string) {
    const np = { ...editPrices }
    for (const r of rows) { if (r.output_file_name === fileName) np[`${r.customer_code}||${r.location}`] = price }
    setEditPrices(np)
  }

  async function generateQuote() {
    const customer = customers.find(c => c.code === customerCode)
    if (!customer || !quotePrice) return
    setGenerating(true)
    const priceNum = parseFloat(quotePrice)
    const pf = formatPrice(priceNum)
    const [logoData, stampData, sigData] = await Promise.all([loadImage('/sig_4.jpg'), loadImage('/sig_1.png'), loadImage('/sig_2.png')])
    const font = 'Times New Roman'; const sz = 24
    const ns = BorderStyle.NONE
    const noBorders = { top:{style:ns}, bottom:{style:ns}, left:{style:ns}, right:{style:ns}, insideHorizontal:{style:ns}, insideVertical:{style:ns} }

    const doc = new Document({ sections: [{ properties: { page: { margin: { top: 720, bottom: 720, left: 1080, right: 1080 } } },
      headers: { default: new Header({ children: [
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: noBorders, rows: [new TableRow({ children: [
          new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, verticalAlign: 'center' as never, children: [new Paragraph({ children: [new ImageRun({ data: logoData, transformation: { width: 80, height: 80 }, type: 'jpg' })] })] }),
          new TableCell({ width: { size: 80, type: WidthType.PERCENTAGE }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'CÔNG TY TNHH THÀNH TÍN LBG', bold: true, size: 28, font })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Trụ sở: 115/22/60 Bis Nguyễn Du, Phường Bến Thành, TP.HCM', size: 20, font })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Email: thanhtinlpg@gmail.com Điện thoại: 092.555.84.84', size: 20, font })] }),
          ] }),
        ] })] }),
        new Paragraph({ spacing: { after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' } }, children: [] }),
      ] }) },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 }, children: [new TextRun({ text: 'THƯ CHÀO GIÁ', bold: true, size: 32, font })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 200 }, children: [new TextRun({ text: `TPHCM, ngày 01 tháng ${monthStr} năm ${year}`, italics: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Kính gửi: QUÝ CÔNG TY KHÁCH HÀNG ', bold: true, size: sz, font }), new TextRun({ text: customer.name.toUpperCase() + '.', bold: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, children: [] }),
        new Paragraph({ spacing: { after: 200 }, indent: { firstLine: 720 }, children: [new TextRun({ text: 'Trước tiên, Công ty TNHH TM DV THÀNH TÍN LBG (Thành Tín) chân thành cảm ơn Quý Khách hàng đã quan tâm đến sản phẩm LPG (Liquefied Petrolium Gas) của chúng tôi.', size: sz, font })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Thành Tín là đại diện phân phối các sản phẩm Khí dầu mỏ hóa lỏng (LPG) của Tập đoàn dầu khí Quốc Gia Việt Nam (PetroVietnam).', size: sz, font })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Với đội ngũ kỹ sư dày dạn kinh nghiệm chúng tôi tự hào mang đến sản phẩm LPG và theo đó các dịch vụ thi công, lắp đặt, bảo trì hệ thống đường ống cung cấp Gas (LPG) cho các nhà máy, cơ sở sản xuất, trường học…', size: sz, font })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Theo như yêu cầu, Thành Tín trân trọng gửi đến Quý Khách hàng thông tin liên quan đến việc cung cấp LPG đóng bình loại Bình 45kg, cụ thể như sau:', size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: '1. Đơn giá: ', bold: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: '- Đơn giá gas đóng Bình 45kg, 12kg giao, lắp đặt cho Quý Khách hàng là: ', size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, indent: { left: 360 }, children: [new TextRun({ text: '- KV TPHCM & Miền Đông Nam Bộ:  ', size: sz, font }), new TextRun({ text: `${pf} VNĐ/kg.`, bold: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, indent: { left: 360 }, children: [new TextRun({ text: `- Đơn giá trên (chưa bao gồm thuế VAT 8%) và áp dụng trong tháng ${monthStr} năm ${year} cho đến khi có thông báo giá mới.`, italics: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, indent: { left: 360 }, children: [new TextRun({ text: '- Đơn giá trên sẽ thay đổi (tăng hoặc giảm) theo giá CP thế giới hàng tháng. (Thành Tín sẽ báo giá áp dụng trong tháng từ ngày 01 đến 03 hàng tháng)', italics: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, indent: { left: 360 }, children: [new TextRun({ text: '- Giá trên đã bao gồm phí vận chuyển theo yêu cầu của Quý khách.', italics: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, indent: { left: 360 }, children: [new TextRun({ text: '- Đầu tư lắp đặt hệ thống mới, bảo trì bảo dưỡng hệ thống hiện hữu theo tháng/Quý.', bold: true, italics: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 200 }, indent: { left: 360 }, children: [new TextRun({ text: '- Công nợ chốt vào ngày cuối tháng và được thanh toán vào ngày 25-30 của tháng tiếp theo.', italics: true, size: sz, font })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: '2. Phương thức thanh toán', bold: true, size: sz, font }), new TextRun({ text: ': Thanh toán bằng tiền mặt hoặc chuyển khoản theo thoả thuận giữa hai Bên.', size: sz, font })] }),
        new Paragraph({ spacing: { after: 100 }, children: [] }),
        new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: 'Rất mong nhận được sự hợp tác và ủng hộ nhiệt tình của Quý khách hàng.', size: sz, font })] }),
        new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: 'Trân trọng cảm ơn!', size: sz, font })] }),
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: noBorders, rows: [new TableRow({ children: [
          new TableCell({ width: { size: 45, type: WidthType.PERCENTAGE }, children: [
            new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: 'Nơi nhận:', bold: true, italics: true, size: 20, font })] }),
            new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: '-   Như trên;', size: 20, font })] }),
            new Paragraph({ children: [new TextRun({ text: '-   Lưu VT, KHKD, TCKT. HH01.', size: 20, font })] }),
          ] }),
          new TableCell({ width: { size: 55, type: WidthType.PERCENTAGE }, children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [new TextRun({ text: 'GIÁM ĐỐC', bold: true, size: 24, font })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: stampData, transformation: { width: 120, height: 80 }, type: 'png' }), new ImageRun({ data: sigData, transformation: { width: 180, height: 140 }, type: 'png' })] }),
          ] }),
        ] })] }),
      ],
    }] })

    // Auto-update prices
    const { data: txs } = await supabase.from('transactions').select('id, b45_delivered, b12_delivered, gas_returned')
      .eq('customer_code', customerCode).gte('delivery_date', from).lte('delivery_date', to)
    let updated = 0
    for (const tx of (txs || [])) {
      const gp = tx.b45_delivered * 45 + tx.b12_delivered * 12 - (tx.gas_returned || 0)
      await supabase.from('transactions').update({ unit_price: priceNum, total_amount: gp * priceNum }).eq('id', tx.id)
      updated++
    }

    const blob = await Packer.toBlob(doc)
    const cn = customer.name.replace(/[^a-zA-Z0-9À-ỹ\s]/g, '').replace(/\s+/g, '_').toUpperCase()
    saveAs(blob, `THANH_TIN_${cn}_${monthStr}.${year}.docx`)
    setGenerating(false)
    if (updated > 0) setMsg(`Đã xuất báo giá + cập nhật giá ${formatPrice(priceNum)} cho ${updated} giao dịch`)
  }

  const noPrice = rows.filter(r => !r.unit_price)
  const fileGroups = [...new Set(rows.map(r => r.output_file_name))].sort()
  const customer = customers.find(c => c.code === customerCode)

  return (
    <div className="space-y-4">
      <h1 className="text-lg md:text-xl font-bold text-gray-800">Giá & Báo giá</h1>

      {/* Tabs */}
      <div className="flex gap-1 bg-gray-100 rounded-lg p-1 w-fit">
        <button onClick={() => setTab('prices')} className={`px-4 py-2 rounded-md text-sm font-medium transition ${tab === 'prices' ? 'bg-white shadow text-blue-700' : 'text-gray-500 hover:text-gray-700'}`}>
          Nhập giá theo tháng
        </button>
        <button onClick={() => setTab('quote')} className={`px-4 py-2 rounded-md text-sm font-medium transition ${tab === 'quote' ? 'bg-white shadow text-blue-700' : 'text-gray-500 hover:text-gray-700'}`}>
          Xuất báo giá
        </button>
      </div>

      {/* Shared month/year */}
      <div className="bg-white rounded-xl border p-3 flex flex-wrap gap-3 items-end">
        <div>
          <label className="block text-xs text-gray-500 mb-1">Tháng</label>
          <select value={month} onChange={e => setMonth(Number(e.target.value))} className="border rounded px-3 py-2 text-sm">
            {Array.from({ length: 12 }, (_, i) => i + 1).map(m => <option key={m} value={m}>Tháng {m}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Năm</label>
          <input type="number" value={year} onChange={e => setYear(Number(e.target.value))} className="border rounded px-3 py-2 text-sm w-24" />
        </div>
        {tab === 'prices' && (
          <>
            <button onClick={loadPrices} disabled={loading} className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
              {loading ? 'Đang tải...' : 'Tải dữ liệu'}
            </button>
            {rows.length > 0 && (
              <button onClick={savePrices} disabled={saving} className="bg-green-600 text-white px-4 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
                {saving ? 'Đang lưu...' : 'Lưu giá'}
              </button>
            )}
          </>
        )}
      </div>

      {msg && <div className="bg-green-50 border border-green-200 rounded-lg px-4 py-3 text-sm text-green-800">{msg}</div>}

      {/* Tab: Nhập giá */}
      {tab === 'prices' && (
        <>
          {noPrice.length > 0 && (
            <div className="bg-amber-50 border border-amber-200 rounded-lg px-3 py-2.5 text-sm text-amber-800">{noPrice.length} địa điểm chưa có giá</div>
          )}
          {rows.length > 0 && (
            <div className="space-y-4">
              {fileGroups.map(fname => {
                const fileRows = rows.filter(r => r.output_file_name === fname)
                return (
                  <div key={fname} className="bg-white rounded-xl border overflow-hidden">
                    <div className="bg-gray-50 border-b px-4 py-3 flex flex-wrap items-center justify-between gap-2">
                      <div>
                        <span className="font-semibold text-gray-800">{fname || 'UNMAPPED'}</span>
                        <span className="ml-2 text-xs text-gray-400">{fileRows.length} địa điểm · {fileRows.reduce((s, r) => s + r.count, 0)} dòng</span>
                      </div>
                      {fileRows.length > 1 && (
                        <div className="flex items-center gap-2">
                          <span className="text-xs text-gray-500">Giá chung:</span>
                          <input type="number" placeholder="Nhập" className="border rounded px-2 py-1 text-sm w-24"
                            onKeyDown={e => { if (e.key === 'Enter') applyToFile(fname, (e.target as HTMLInputElement).value) }}
                            onBlur={e => { if (e.target.value) applyToFile(fname, e.target.value) }} />
                        </div>
                      )}
                    </div>
                    <div className="divide-y">
                      {fileRows.map(r => {
                        const key = `${r.customer_code}||${r.location}`
                        const hasPrice = editPrices[key] && parseFloat(editPrices[key]) > 0
                        return (
                          <div key={key} className={`px-4 py-2.5 flex flex-wrap items-center justify-between gap-2 ${!hasPrice ? 'bg-amber-50' : ''}`}>
                            <div>
                              <span className="font-medium text-sm">{r.location}</span>
                              <span className="ml-2 text-xs text-gray-400">{r.customer_name} · {r.count} dòng</span>
                            </div>
                            <input type="number" value={editPrices[key] || ''} onChange={e => setEditPrices({ ...editPrices, [key]: e.target.value })}
                              placeholder="Giá" className={`border rounded px-3 py-1.5 text-sm w-32 text-right ${!hasPrice ? 'border-amber-400 bg-amber-50' : ''}`} />
                          </div>
                        )
                      })}
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </>
      )}

      {/* Tab: Xuất báo giá */}
      {tab === 'quote' && (
        <div className="bg-white rounded-xl border p-5 space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs text-gray-500 mb-1">Khách hàng *</label>
              <select value={customerCode} onChange={e => setCustomerCode(e.target.value)} className="border rounded px-3 py-2 text-sm w-full">
                <option value="">-- Chọn KH --</option>
                {customers.map(c => <option key={c.code} value={c.code}>{c.code} – {c.name}</option>)}
              </select>
            </div>
            <div>
              <label className="block text-xs text-gray-500 mb-1">Đơn giá (VNĐ/kg, chưa VAT) *</label>
              <input type="number" value={quotePrice} onChange={e => setQuotePrice(e.target.value)} placeholder="vd: 41989" className="border rounded px-3 py-2 text-sm w-full" />
              {quotePrice && <span className="text-xs text-gray-400 mt-1 block">{formatPrice(parseFloat(quotePrice))} VNĐ/kg</span>}
            </div>
          </div>
          <button onClick={generateQuote} disabled={generating || !customerCode || !quotePrice}
            className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
            {generating ? 'Đang tạo...' : 'Xuất DOCX'}
          </button>

          {customer && quotePrice && (
            <div className="border rounded-lg p-4 md:p-6 max-w-2xl mx-auto text-sm leading-relaxed mt-4" style={{ fontFamily: 'Times New Roman, serif' }}>
              <div className="flex items-center gap-4 mb-4 pb-3 border-b">
                <img src="/sig_4.jpg" alt="Logo" className="h-12 md:h-16" />
                <div className="text-center flex-1">
                  <p className="font-bold text-sm md:text-base">CÔNG TY TNHH THÀNH TÍN LBG</p>
                  <p className="text-xs text-gray-600">Trụ sở: 115/22/60 Bis Nguyễn Du, Phường Bến Thành, TP.HCM</p>
                </div>
              </div>
              <p className="text-center font-bold text-lg mb-3">THƯ CHÀO GIÁ</p>
              <p className="text-right italic mb-3 text-xs">TPHCM, ngày 01 tháng {monthStr} năm {year}</p>
              <p className="mb-3"><strong>Kính gửi: {customer.name.toUpperCase()}.</strong></p>
              <p className="font-bold mb-1 text-xs">1. Đơn giá:</p>
              <p className="ml-4 mb-1 text-xs">- KV TPHCM & ĐNB: <strong>{formatPrice(parseFloat(quotePrice))} VNĐ/kg</strong> (chưa VAT, tháng {monthStr}/{year})</p>
              <p className="font-bold mt-3 mb-1 text-xs">2. Thanh toán: <span className="font-normal">Tiền mặt hoặc chuyển khoản.</span></p>
              <div className="flex justify-between items-start mt-4">
                <div className="text-xs"><p className="font-bold italic">Nơi nhận:</p><p>- Như trên</p></div>
                <div className="text-center">
                  <p className="font-bold text-xs mb-1">GIÁM ĐỐC</p>
                  <img src="/sig_1.png" alt="Dấu" className="h-12 inline-block" />
                  <img src="/sig_2.png" alt="Ký" className="h-16 inline-block -ml-6" />
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
