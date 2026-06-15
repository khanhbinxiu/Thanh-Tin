import Link from 'next/link'

export default function Home() {
  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold text-gray-800">Quản lý công nợ – Thành Tín LBG</h1>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        {[
          { href: '/customers', title: 'Khách hàng', desc: 'Thêm / sửa danh mục khách hàng' },
          { href: '/upload', title: 'Upload Excel', desc: 'Nhập dữ liệu giao dịch từ file tổng' },
          { href: '/export', title: 'Xuất biên bản', desc: 'Tạo biên bản đối chiếu công nợ theo kỳ' },
          { href: '/review', title: 'Xem dữ liệu', desc: 'Tra cứu & xuất dữ liệu thô' },
        ].map((c) => (
          <Link
            key={c.href}
            href={c.href}
            className="bg-white rounded-xl border border-gray-200 p-5 hover:shadow-md transition-shadow"
          >
            <div className="font-semibold text-blue-700 mb-1">{c.title}</div>
            <div className="text-sm text-gray-500">{c.desc}</div>
          </Link>
        ))}
      </div>
    </div>
  )
}
