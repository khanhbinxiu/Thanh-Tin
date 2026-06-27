'use client'
import Link from 'next/link'
import { usePathname } from 'next/navigation'

const links = [
  { href: '/', label: 'Tổng quan' },
  { href: '/customers', label: 'Khách hàng' },
  { href: '/mappings', label: 'Mapping' },
  { href: '/upload', label: 'Upload Excel' },
  { href: '/manual', label: 'Nhập tay' },
  { href: '/prices', label: 'Nhập giá' },
  { href: '/export', label: 'Xuất biên bản' },
  { href: '/quote', label: 'Báo giá' },
  { href: '/review', label: 'Xem dữ liệu' },
]

export default function Nav() {
  const pathname = usePathname()
  return (
    <nav className="bg-blue-700 text-white shadow">
      <div className="max-w-7xl mx-auto px-4 flex items-center gap-1 h-14">
        <span className="font-bold text-lg mr-6">Thành Tín LBG</span>
        {links.map((l) => (
          <Link
            key={l.href}
            href={l.href}
            className={`px-3 py-1.5 rounded text-sm font-medium transition-colors ${
              pathname === l.href ? 'bg-blue-900' : 'hover:bg-blue-600'
            }`}
          >
            {l.label}
          </Link>
        ))}
      </div>
    </nav>
  )
}
