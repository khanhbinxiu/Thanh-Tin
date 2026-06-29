'use client'
import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { useState } from 'react'

const groups = [
  {
    label: 'Quản lý',
    color: 'bg-blue-600',
    mobileColor: 'border-blue-400',
    links: [
      { href: '/', label: 'Tổng quan', icon: '📊' },
      { href: '/customers', label: 'Khách hàng', icon: '👥' },
      { href: '/mappings', label: 'Mapping', icon: '🔗' },
    ],
  },
  {
    label: 'Nhập liệu',
    color: 'bg-emerald-600',
    mobileColor: 'border-emerald-400',
    links: [
      { href: '/preorder', label: 'Đặt trước', icon: '📦' },
      { href: '/manual', label: 'Nhập tay', icon: '✏️' },
      { href: '/prices', label: 'Giá & Báo giá', icon: '💰' },
    ],
  },
  {
    label: 'Xuất & Báo cáo',
    color: 'bg-violet-600',
    mobileColor: 'border-violet-400',
    links: [
      { href: '/export', label: 'Xuất biên bản', icon: '📄' },
      { href: '/inventory', label: 'Tồn kho', icon: '🏭' },
      { href: '/review', label: 'Xem dữ liệu', icon: '🔍' },
    ],
  },
]

export default function Nav() {
  const pathname = usePathname()
  const [open, setOpen] = useState(false)

  return (
    <nav className="bg-gray-900 text-white shadow sticky top-0 z-50">
      <div className="max-w-7xl mx-auto px-4 flex items-center justify-between h-14">
        <Link href="/" className="font-bold text-lg">Thành Tín LBG</Link>

        {/* Desktop nav */}
        <div className="hidden lg:flex items-center gap-0.5">
          {groups.map((g) => (
            <div key={g.label} className="flex items-center">
              {g.links.map((l) => (
                <Link key={l.href} href={l.href}
                  className={`px-2.5 py-1.5 rounded text-xs font-medium transition-colors ${
                    pathname === l.href ? g.color + ' text-white' : 'text-gray-300 hover:text-white hover:bg-gray-700'
                  }`}>
                  {l.label}
                </Link>
              ))}
              <div className="w-px h-6 bg-gray-700 mx-1" />
            </div>
          ))}
        </div>

        {/* Mobile hamburger */}
        <button onClick={() => setOpen(!open)} className="lg:hidden p-2 rounded hover:bg-gray-700">
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
        <div className="lg:hidden border-t border-gray-700 pb-2">
          {groups.map((g) => (
            <div key={g.label}>
              <div className="px-4 py-2 text-xs text-gray-400 uppercase tracking-wider border-l-4 border-transparent">
                {g.label}
              </div>
              {g.links.map((l) => (
                <Link key={l.href} href={l.href} onClick={() => setOpen(false)}
                  className={`block px-4 py-3 text-sm font-medium border-l-4 ${
                    pathname === l.href ? g.color + ' ' + g.mobileColor : 'border-transparent hover:bg-gray-800'
                  }`}>
                  {l.icon} {l.label}
                </Link>
              ))}
            </div>
          ))}
        </div>
      )}
    </nav>
  )
}
