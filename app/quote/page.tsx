'use client'
import { useEffect, useState } from 'react'
import { supabase, Customer } from '@/lib/supabase'
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, Header, Table, TableRow, TableCell, WidthType, BorderStyle } from 'docx'
import { saveAs } from 'file-saver'

function formatPrice(n: number) {
  return n.toLocaleString('vi-VN')
}

async function loadImage(url: string): Promise<ArrayBuffer> {
  const res = await fetch(url)
  return res.arrayBuffer()
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

  async function generate() {
    const customer = customers.find(c => c.code === customerCode)
    if (!customer || !price) return

    setGenerating(true)
    const priceNum = parseFloat(price)
    const priceFormatted = formatPrice(priceNum)
    const monthStr = String(month).padStart(2, '0')

    const [logoData, stampData, sigData] = await Promise.all([
      loadImage('/sig_4.jpg'),
      loadImage('/sig_1.png'),
      loadImage('/sig_2.png'),
    ])

    const font = 'Times New Roman'
    const sz = 24

    const doc = new Document({
      sections: [{
        properties: {
          page: { margin: { top: 720, bottom: 720, left: 1080, right: 1080 } },
        },
        headers: {
          default: new Header({
            children: [
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE}, insideHorizontal: {style:BorderStyle.NONE}, insideVertical: {style:BorderStyle.NONE} },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        verticalAlign: 'center' as never,
                        children: [
                          new Paragraph({
                            children: [
                              new ImageRun({ data: logoData, transformation: { width: 80, height: 80 }, type: 'jpg' }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        width: { size: 80, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: 'CÔNG TY TNHH THÀNH TÍN LBG', bold: true, size: 28, font }),
                            ],
                          }),
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: 'Trụ sở: 115/22/60 Bis Nguyễn Du, Phường Bến Thành, TP.HCM', size: 20, font }),
                            ],
                          }),
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: 'Email: thanhtinlpg@gmail.com', size: 20, font }),
                              new TextRun({ text: ' Điện thoại: 092.555.84.84', size: 20, font }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new Paragraph({
                spacing: { after: 200 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
                children: [],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
            children: [
              new TextRun({ text: 'THƯ CHÀO GIÁ', bold: true, size: 32, font }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 200 },
            children: [
              new TextRun({ text: `TPHCM, ngày 01 tháng ${monthStr} năm ${year}`, italics: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Kính gửi: QUÝ CÔNG TY KHÁCH HÀNG ', bold: true, size: sz, font }),
              new TextRun({ text: customer.name.toUpperCase(), bold: true, size: sz, font }),
              new TextRun({ text: '.', bold: true, size: sz, font }),
            ],
          }),

          new Paragraph({ spacing: { after: 100 }, children: [] }),

          new Paragraph({
            spacing: { after: 200 },
            indent: { firstLine: 720 },
            children: [
              new TextRun({ text: 'Trước tiên, Công ty TNHH TM DV THÀNH TÍN LBG (Thành Tín) chân thành cảm ơn Quý Khách hàng đã quan tâm đến sản phẩm LPG (Liquefied Petrolium Gas) của chúng tôi.', size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Thành Tín là đại diện phân phối các sản phẩm Khí dầu mỏ hóa lỏng (LPG) của Tập đoàn dầu khí Quốc Gia Việt Nam (PetroVietnam).', size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Với đội ngũ kỹ sư dày dạn kinh nghiệm chúng tôi tự hào mang đến sản phẩm LPG và theo đó các dịch vụ thi công, lắp đặt, bảo trì hệ thống đường ống cung cấp Gas (LPG) cho các nhà máy, cơ sở sản xuất, trường học…', size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: 'Theo như yêu cầu, Thành Tín trân trọng gửi đến Quý Khách hàng thông tin liên quan đến việc cung cấp LPG đóng bình loại Bình 45kg, cụ thể như sau:', size: sz, font }),
            ],
          }),

          // 1. Đơn giá
          new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({ text: '1. Đơn giá: ', bold: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({ text: '- Đơn giá gas đóng Bình 45kg, 12kg giao, lắp đặt cho Quý Khách hàng là: ', size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: `- KV TPHCM & Miền Đông Nam Bộ:  `, size: sz, font }),
              new TextRun({ text: `${priceFormatted} VNĐ/kg.`, bold: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: `- Đơn giá trên (chưa bao gồm thuế VAT 8%) và áp dụng trong tháng ${monthStr} năm ${year} cho đến khi có thông báo giá mới.`, italics: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Đơn giá trên sẽ thay đổi (tăng hoặc giảm) theo giá CP thế giới hàng tháng. (Thành Tín sẽ báo giá áp dụng trong tháng từ ngày 01 đến 03 hàng tháng)', italics: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Giá trên đã bao gồm phí vận chuyển theo yêu cầu của Quý khách.', italics: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 100 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Đầu tư lắp đặt hệ thống mới, bảo trì bảo dưỡng hệ thống hiện hữu theo tháng/Quý.', bold: true, italics: true, size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            indent: { left: 360 },
            children: [
              new TextRun({ text: '- Công nợ chốt vào ngày cuối tháng và được thanh toán vào ngày 25-30 của tháng tiếp theo.', italics: true, size: sz, font }),
            ],
          }),

          // 2. Phương thức thanh toán
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({ text: '2. Phương thức thanh toán', bold: true, size: sz, font }),
              new TextRun({ text: ': Thanh toán bằng tiền mặt hoặc chuyển khoản theo thoả thuận giữa hai Bên.', size: sz, font }),
            ],
          }),

          new Paragraph({ spacing: { after: 100 }, children: [] }),

          new Paragraph({
            spacing: { after: 100 },
            children: [
              new TextRun({ text: 'Rất mong nhận được sự hợp tác và ủng hộ nhiệt tình của Quý khách hàng.', size: sz, font }),
            ],
          }),

          new Paragraph({
            spacing: { after: 300 },
            children: [
              new TextRun({ text: 'Trân trọng cảm ơn!', size: sz, font }),
            ],
          }),

          // Nơi nhận + Chữ ký side by side
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideHorizontal:{style:BorderStyle.NONE}, insideVertical:{style:BorderStyle.NONE} },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 45, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        spacing: { after: 50 },
                        children: [
                          new TextRun({ text: 'Nơi nhận:', bold: true, italics: true, size: 20, font }),
                        ],
                      }),
                      new Paragraph({
                        spacing: { after: 30 },
                        children: [
                          new TextRun({ text: '-   Như trên;', size: 20, font }),
                        ],
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({ text: '-   Lưu VT, KHKD, TCKT. HH01.', size: 20, font }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 55, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 50 },
                        children: [
                          new TextRun({ text: 'GIÁM ĐỐC', bold: true, size: 24, font }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new ImageRun({ data: stampData, transformation: { width: 120, height: 80 }, type: 'png' }),
                          new ImageRun({ data: sigData, transformation: { width: 180, height: 140 }, type: 'png' }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }],
    })

    const blob = await Packer.toBlob(doc)
    const customerName = customer.name.replace(/[^a-zA-Z0-9À-ỹ\s]/g, '').replace(/\s+/g, '_').toUpperCase()
    saveAs(blob, `THANH_TIN_${customerName}_${monthStr}.${year}.docx`)
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

        <button onClick={generate} disabled={generating || !customerCode || !price}
          className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
          {generating ? 'Đang tạo...' : 'Xuất DOCX'}
        </button>
      </div>

      {customer && price && (
        <div className="bg-white rounded-xl border p-5 space-y-3">
          <h2 className="font-semibold text-gray-700">Xem trước</h2>
          <div className="border rounded-lg p-6 max-w-2xl mx-auto text-sm leading-relaxed" style={{ fontFamily: 'Times New Roman, serif' }}>
            <div className="flex items-center gap-4 mb-4 pb-3 border-b">
              <img src="/sig_4.jpg" alt="Logo" className="h-16" />
              <div className="text-center flex-1">
                <p className="font-bold text-base">CÔNG TY TNHH THÀNH TÍN LBG</p>
                <p className="text-xs text-gray-600">Trụ sở: 115/22/60 Bis Nguyễn Du, Phường Bến Thành, TP.HCM</p>
                <p className="text-xs text-gray-600">Email: thanhtinlpg@gmail.com Điện thoại: 092.555.84.84</p>
              </div>
            </div>
            <p className="text-center font-bold text-lg mb-4">THƯ CHÀO GIÁ</p>
            <p className="text-right italic mb-4">TPHCM, ngày 01 tháng {String(month).padStart(2, '0')} năm {year}</p>
            <p className="mb-4">
              <strong>Kính gửi: QUÝ CÔNG TY KHÁCH HÀNG {customer.name.toUpperCase()}.</strong>
            </p>
            <p className="indent-8 mb-3">Trước tiên, Công ty TNHH TM DV THÀNH TÍN LBG (Thành Tín) chân thành cảm ơn Quý Khách hàng đã quan tâm đến sản phẩm LPG của chúng tôi.</p>
            <p className="mb-4">Theo yêu cầu, Thành Tín trân trọng gửi thông tin cung cấp LPG đóng bình loại Bình 45kg:</p>
            <p className="font-bold mb-2">1. Đơn giá:</p>
            <p className="ml-4 mb-1">- Đơn giá gas đóng Bình 45kg, 12kg:</p>
            <p className="ml-8 mb-1">- KV TPHCM & Miền Đông Nam Bộ: <strong>{formatPrice(parseFloat(price))} VNĐ/kg.</strong></p>
            <p className="ml-8 italic mb-1 text-xs text-gray-600">Chưa bao gồm thuế VAT 8%, áp dụng tháng {String(month).padStart(2, '0')}/{year}</p>
            <p className="font-bold mt-4 mb-2">2. Phương thức thanh toán: <span className="font-normal">Tiền mặt hoặc chuyển khoản.</span></p>
            <p className="mt-4 mb-6">Trân trọng cảm ơn!</p>
            <div className="flex justify-between items-start">
              <div className="text-xs">
                <p className="font-bold italic">Nơi nhận:</p>
                <p>- Như trên</p>
                <p>- Lưu VT, KHKD</p>
              </div>
              <div className="text-center">
                <p className="font-bold text-sm mb-1">GIÁM ĐỐC</p>
                <div className="relative">
                  <img src="/sig_1.png" alt="Con dấu" className="h-20 inline-block" />
                  <img src="/sig_2.png" alt="Chữ ký" className="h-24 inline-block -ml-8" />
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
