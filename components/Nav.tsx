'use client'
import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { useState } from 'react'

const links = [
  { href: '/', label: 'Tổng quan', icon: '📊' },
  { href: '/customers', label: 'Khách hàng', icon: '👥' },
  { href: '/mappings', label: 'Mapping', icon: '🔗' },
  { href: '/manual', label: 'Nhập tay', icon: '✏️' },
  { href: '/preorder', label: 'Đặt trước', icon: '📦' },
  { href: '/prices', label: 'Giá & Báo giá', icon: '💰' },
  { href: '/export', label: 'Xuất biên bản', icon: '📄' },
  { href: '/inventory', label: 'Tồn kho', icon: '🏭' },
  { href: '/review', label: 'Xem dữ liệu', icon: '🔍' },
]

export default function Nav() {
  const pathname = usePathname()
  const [open, setOpen] = useState(false)

  return (
    <nav className="bg-blue-700 text-white shadow sticky top-0 z-50">
      <div className="max-w-7xl mx-auto px-4 flex items-center justify-between h-14">
        <Link href="/" className="font-bold text-lg">Thành Tín LBG</Link>

        {/* Desktop nav */}
        <div className="hidden md:flex items-center gap-1">
          {links.map((l) => (
            <Link key={l.href} href={l.href}
              className={`px-3 py-1.5 rounded text-sm font-medium transition-colors ${
                pathname === l.href ? 'bg-blue-900' : 'hover:bg-blue-600'
              }`}>
              {l.label}
            </Link>
          ))}
        </div>

        {/* Mobile hamburger */}
        <button onClick={() => setOpen(!open)} className="md:hidden p-2 rounded hover:bg-blue-600">
          <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            {open
              ? <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
              : <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 12h16M4 18h16" />
            }
          </svg>
        </button>
      </div>

      {/* Mobile menu */}
      {open && (
        <div className="md:hidden border-t border-blue-600 pb-2">
          {links.map((l) => (
            <Link key={l.href} href={l.href} onClick={() => setOpen(false)}
              className={`block px-4 py-3 text-sm font-medium ${
                pathname === l.href ? 'bg-blue-900' : 'hover:bg-blue-600'
              }`}>
              {l.icon} {l.label}
            </Link>
          ))}
        </div>
      )}
    </nav>
  )
}
