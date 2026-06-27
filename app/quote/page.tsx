'use client'
import { useEffect, useState } from 'react'
import { supabase, Customer } from '@/lib/supabase'
import { Document, Packer, Paragraph, TextRun, AlignmentType, TabStopPosition, TabStopType, ImageRun } from 'docx'
import { saveAs } from 'file-saver'

function numberToWords(n: number): string {
  const units = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín']
  const tens = ['', 'mười', 'hai mươi', 'ba mươi', 'bốn mươi', 'năm mươi', 'sáu mươi', 'bảy mươi', 'tám mươi', 'chín mươi']
  if (n < 10) return units[n]
  if (n < 100) {
    const t = Math.floor(n / 10)
    const u = n % 10
    return tens[t] + (u ? ' ' + units[u] : '')
  }
  if (n < 1000) {
    const h = Math.floor(n / 100)
    const r = n % 100
    return units[h] + ' trăm' + (r ? ' ' + numberToWords(r) : '')
  }
  if (n < 1000000) {
    const k = Math.floor(n / 1000)
    const r = n % 1000
    return numberToWords(k) + ' nghìn' + (r ? ' ' + numberToWords(r) : '')
  }
  return String(n)
}

function formatPrice(n: number) {
  return n.toLocaleString('vi-VN')
}

export default function QuotePage() {
  const [customers, setCustomers] = useState<Customer[]>([])
  const [customerCode, setCustomerCode] = useState('')
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [year, setYear] = useState(new Date().getFullYear())
  const [price, setPrice] = useState('')
  const [generating, setGenerating] = useState(false)

  useEffect(() => {
    supabase.from('customers').select('*').order('name').then(({ data }) => setCustomers(data || []))
  }, [])

  async function generate(format: 'docx' | 'pdf') {
    const customer = customers.find(c => c.code === customerCode)
    if (!customer || !price) return

    setGenerating(true)
    const priceNum = parseFloat(price)
    const priceFormatted = formatPrice(priceNum)
    const monthStr = String(month).padStart(2, '0')

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 1080, right: 1080 },
          },
        },
        children: [
          // Title
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'THƯ CHÀO GIÁ', bold: true, size: 32, font: 'Times New Roman' }),
            ],
          }),

          // Date
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 200 },
            children: [
              new TextRun({ text: `TPHCM, ngày 01 tháng ${monthStr} năm ${year}`, italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          // Customer name
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Kính gửi: QUÝ CÔNG TY KHÁCH HÀNG ', bold: true, size: 24, font: 'Times New Roman' }),
              new TextRun({ text: customer.name.toUpperCase(), bold: true, size: 24, font: 'Times New Roman', highlight: 'yellow' }),
              new TextRun({ text: '.', bold: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({ spacing: { after: 100 }, children: [] }),

          // Intro paragraph
          new Paragraph({
            spacing: { after: 200 },
            indent: { firstLine: 720 },
            children: [
              new TextRun({ text: 'Trước tiên, Công ty TNHH TM DV THÀNH TÍN LBG (Thành Tín) chân thành cảm ơn Quý Khách hàng đã quan tâm đến sản phẩm LPG (Liquefied Petrolium Gas) của chúng tôi.', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Thành Tín là đại diện phân phối các sản phẩm Khí dầu mỏ hóa lỏng (LPG) của Tập đoàn dầu khí Quốc Gia Việt Nam (PetroVietnam).', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Với đội ngũ kỹ sư dày dạn kinh nghiệm chúng tôi tự hào mang đến sản phẩm LPG và theo đó các dịch vụ thi công, lắp đặt, bảo trì hệ thống đường ống cung cấp Gas (LPG) cho các nhà máy, cơ sở sản xuất, trường học…', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Theo như yêu cầu, Thành Tín trân trọng gửi đến Quý Khách hàng thông tin liên quan đến việc cung cấp LPG đóng bình loại Bình 45kg, cụ thể như sau:', size: 24, font: 'Times New Roman' }),
            ],
          }),

          // Section 1: Đơn giá
          new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({ text: '1. Đơn giá: ', bold: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({ text: '- Đơn giá gas đóng Bình 45kg, 12kg giao, lắp đặt cho Quý Khách hàng là: ', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: `- KV TPHCM & Miền Đông Nam Bộ:  `, size: 24, font: 'Times New Roman' }),
              new TextRun({ text: `${priceFormatted} VNĐ/kg.`, bold: true, size: 24, font: 'Times New Roman', highlight: 'yellow' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: `- Đơn giá trên (chưa bao gồm thuế VAT 8%) và áp dụng trong `, italics: true, size: 24, font: 'Times New Roman' }),
              new TextRun({ text: `tháng ${monthStr} năm ${year}`, italics: true, bold: true, size: 24, font: 'Times New Roman', highlight: 'yellow' }),
              new TextRun({ text: ` cho đến khi có thông báo giá mới.`, italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Đơn giá trên sẽ thay đổi (tăng hoặc giảm) theo giá CP thế giới hàng tháng. (Thành Tín sẽ báo giá áp dụng trong tháng từ ngày 01 đến 03 hàng tháng)', italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Giá trên đã bao gồm phí vận chuyển theo yêu cầu của Quý khách.', italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Đầu tư lắp đặt hệ thống mới. bảo trì bảo dưỡng hệ thống hiện hữu theo tháng/Quý', bold: true, italics: true, size: 24, font: 'Times New Roman' }),
              new TextRun({ text: '.', italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Công nợ chốt vào ngày cuối tháng và được thanh toán vào ngày 25-30 của tháng tiếp theo.', italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          // Section 2: Phương thức thanh toán
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: '2. Phương thức thanh toán', bold: true, size: 24, font: 'Times New Roman' }),
              new TextRun({ text: ': Thanh toán bằng tiền mặt hoặc chuyển khoản theo thoả thuận giữa hai Bên.', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({ spacing: { after: 100 }, children: [] }),

          // Closing
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Rất mong nhận được sự hợp tác và ủng hộ nhiệt tình của Quý khách hàng.', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Trân trọng cảm ơn!', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({ spacing: { after: 100 }, children: [] }),

          // Footer
          new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({ text: 'Nơi nhận:', bold: true, italics: true, size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 50 },
            children: [
              new TextRun({ text: '-   Như trên;', size: 24, font: 'Times New Roman' }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: '-   Lưu VT,  KHKD, TCKT. HH01.', size: 24, font: 'Times New Roman' }),
            ],
          }),
        ],
      }],
    })

    const blob = await Packer.toBlob(doc)
    const customerName = customer.name.replace(/[^a-zA-Z0-9À-ỹ\s]/g, '').replace(/\s+/g, '_').toUpperCase()
    const fileName = `THANH_TIN_${customerName}_${monthStr}.${year}`

    if (format === 'docx') {
      saveAs(blob, `${fileName}.docx`)
    } else {
      saveAs(blob, `${fileName}.docx`)
      alert('PDF: Mở file .docx vừa tải → File → Save As → chọn PDF')
    }

    setGenerating(false)
  }

  const customer = customers.find(c => c.code === customerCode)

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Xuất báo giá</h1>

      <div className="bg-white rounded-xl border p-5 space-y-4">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="col-span-2">
            <label className="block text-xs text-gray-500 mb-1">Khách hàng *</label>
            <select value={customerCode} onChange={e => setCustomerCode(e.target.value)}
              className="border rounded px-3 py-2 text-sm w-full">
              <option value="">-- Chọn KH --</option>
              {customers.map(c => <option key={c.code} value={c.code}>{c.code} – {c.name}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Tháng</label>
            <select value={month} onChange={e => setMonth(Number(e.target.value))}
              className="border rounded px-3 py-2 text-sm w-full">
              {Array.from({ length: 12 }, (_, i) => i + 1).map(m =>
                <option key={m} value={m}>Tháng {m}</option>
              )}
            </select>
          </div>
          <div>
            <label className="block text-xs text-gray-500 mb-1">Năm</label>
            <input type="number" value={year} onChange={e => setYear(Number(e.target.value))}
              className="border rounded px-3 py-2 text-sm w-full" />
          </div>
          <div className="col-span-2">
            <label className="block text-xs text-gray-500 mb-1">Đơn giá (VNĐ/kg, chưa VAT) *</label>
            <input type="number" value={price} onChange={e => setPrice(e.target.value)}
              placeholder="vd: 41989"
              className="border rounded px-3 py-2 text-sm w-full" />
            {price && <span className="text-xs text-gray-400 mt-1 block">{formatPrice(parseFloat(price))} VNĐ/kg</span>}
          </div>
        </div>

        <div className="flex gap-3">
          <button onClick={() => generate('docx')} disabled={generating || !customerCode || !price}
            className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
            {generating ? 'Đang tạo...' : 'Xuất DOCX'}
          </button>
          <button onClick={() => generate('pdf')} disabled={generating || !customerCode || !price}
            className="bg-green-600 text-white px-5 py-2 rounded text-sm hover:bg-green-700 disabled:opacity-50">
            Xuất PDF
          </button>
        </div>
      </div>

      {customer && price && (
        <div className="bg-white rounded-xl border p-5 space-y-3">
          <h2 className="font-semibold text-gray-700">Xem trước</h2>
          <div className="border rounded-lg p-6 max-w-2xl mx-auto text-sm leading-relaxed" style={{ fontFamily: 'Times New Roman, serif' }}>
            <p className="text-center font-bold text-lg mb-4">THƯ CHÀO GIÁ</p>
            <p className="text-right italic mb-4">TPHCM, ngày 01 tháng {String(month).padStart(2, '0')} năm {year}</p>
            <p className="mb-4">
              <strong>Kính gửi: QUÝ CÔNG TY KHÁCH HÀNG <span className="bg-yellow-200">{customer.name.toUpperCase()}</span>.</strong>
            </p>
            <p className="indent-8 mb-3">Trước tiên, Công ty TNHH TM DV THÀNH TÍN LBG (Thành Tín) chân thành cảm ơn Quý Khách hàng đã quan tâm đến sản phẩm LPG của chúng tôi.</p>
            <p className="mb-3">Thành Tín là đại diện phân phối các sản phẩm Khí dầu mỏ hóa lỏng (LPG) của Tập đoàn dầu khí Quốc Gia Việt Nam (PetroVietnam).</p>
            <p className="mb-4">Theo như yêu cầu, Thành Tín trân trọng gửi đến Quý Khách hàng thông tin cung cấp LPG đóng bình loại Bình 45kg:</p>
            <p className="font-bold mb-2">1. Đơn giá:</p>
            <p className="ml-4 mb-1">- Đơn giá gas đóng Bình 45kg, 12kg giao, lắp đặt cho Quý Khách hàng là:</p>
            <p className="ml-8 mb-1">- KV TPHCM & Miền Đông Nam Bộ: <strong className="bg-yellow-200">{formatPrice(parseFloat(price))} VNĐ/kg.</strong></p>
            <p className="ml-8 italic mb-1">- Đơn giá trên (chưa bao gồm thuế VAT 8%) và áp dụng trong <strong className="bg-yellow-200">tháng {String(month).padStart(2, '0')} năm {year}</strong> cho đến khi có thông báo giá mới.</p>
            <p className="font-bold mt-4 mb-2">2. Phương thức thanh toán: <span className="font-normal">Thanh toán bằng tiền mặt hoặc chuyển khoản.</span></p>
            <p className="mt-4">Trân trọng cảm ơn!</p>
          </div>
        </div>
      )}
    </div>
  )
}
