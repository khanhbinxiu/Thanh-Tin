'use client'
import { useEffect, useState, useRef } from 'react'
import { supabase } from '@/lib/supabase'

type MonthRow = {
  label: string
  deliveries: number
  gasDelivered: number
  gasPaid: number
  revenue: number
  revenueVat: number
}

export default function Home() {
  const [dateFrom, setDateFrom] = useState(`${new Date().getFullYear()}-01-01`)
  const [dateTo, setDateTo] = useState(new Date().toISOString().slice(0, 10))
  const [data, setData] = useState<MonthRow[]>([])
  const [loading, setLoading] = useState(false)
  const canvasRef = useRef<HTMLCanvasElement>(null)

  async function load() {
    setLoading(true)
    const { data: txs } = await supabase.from('transactions')
      .select('delivery_date, gas_delivered, gas_paid, total_amount')
      .gte('delivery_date', dateFrom)
      .lte('delivery_date', dateTo)

    const grouped = new Map<string, MonthRow>()
    for (const tx of (txs || [])) {
      const [y, m] = tx.delivery_date.split('-')
      const key = `${y}-${m}`
      const label = `T${parseInt(m)}/${y}`
      if (!grouped.has(key)) {
        grouped.set(key, { label, deliveries: 0, gasDelivered: 0, gasPaid: 0, revenue: 0, revenueVat: 0 })
      }
      const row = grouped.get(key)!
      row.deliveries++
      row.gasDelivered += tx.gas_delivered || 0
      row.gasPaid += tx.gas_paid || 0
      row.revenue += tx.total_amount || 0
      row.revenueVat += (tx.total_amount || 0) * 1.08
    }

    const sorted = [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(e => e[1])
    setData(sorted)
    setLoading(false)
  }

  useEffect(() => { load() }, [])

  useEffect(() => {
    if (!canvasRef.current || data.length === 0) return
    const canvas = canvasRef.current
    const ctx = canvas.getContext('2d')
    if (!ctx) return

    const dpr = window.devicePixelRatio || 1
    const w = canvas.clientWidth
    const h = canvas.clientHeight
    canvas.width = w * dpr
    canvas.height = h * dpr
    ctx.scale(dpr, dpr)
    ctx.clearRect(0, 0, w, h)

    const pad = { top: 30, right: 20, bottom: 50, left: 80 }
    const chartW = w - pad.left - pad.right
    const chartH = h - pad.top - pad.bottom
    const maxVal = Math.max(...data.map(d => d.revenueVat), 1)
    const barW = Math.min(60, (chartW / data.length) * 0.6)
    const gap = chartW / data.length

    // Grid lines
    ctx.strokeStyle = '#e5e7eb'
    ctx.lineWidth = 1
    const gridLines = 5
    for (let i = 0; i <= gridLines; i++) {
      const y = pad.top + chartH - (chartH * i / gridLines)
      ctx.beginPath()
      ctx.moveTo(pad.left, y)
      ctx.lineTo(w - pad.right, y)
      ctx.stroke()

      const val = (maxVal * i / gridLines)
      ctx.fillStyle = '#9ca3af'
      ctx.font = '11px sans-serif'
      ctx.textAlign = 'right'
      ctx.fillText(val >= 1e6 ? (val / 1e6).toFixed(1) + 'M' : val >= 1e3 ? (val / 1e3).toFixed(0) + 'K' : String(Math.round(val)), pad.left - 8, y + 4)
    }

    // Bars
    data.forEach((d, i) => {
      const x = pad.left + gap * i + (gap - barW) / 2
      const barH = (d.revenueVat / maxVal) * chartH
      const y = pad.top + chartH - barH

      // Gradient
      const grad = ctx.createLinearGradient(x, y, x, y + barH)
      grad.addColorStop(0, '#3b82f6')
      grad.addColorStop(1, '#1d4ed8')
      ctx.fillStyle = grad
      ctx.beginPath()
      ctx.roundRect(x, y, barW, barH, [4, 4, 0, 0])
      ctx.fill()

      // Value on top
      if (d.revenueVat > 0) {
        ctx.fillStyle = '#1e40af'
        ctx.font = 'bold 10px sans-serif'
        ctx.textAlign = 'center'
        const valText = d.revenueVat >= 1e6 ? (d.revenueVat / 1e6).toFixed(1) + 'M' : (d.revenueVat / 1e3).toFixed(0) + 'K'
        ctx.fillText(valText, x + barW / 2, y - 6)
      }

      // Label
      ctx.fillStyle = '#374151'
      ctx.font = '11px sans-serif'
      ctx.textAlign = 'center'
      ctx.fillText(d.label, x + barW / 2, pad.top + chartH + 20)
    })
  }, [data])

  const totals = data.reduce((acc, d) => ({
    deliveries: acc.deliveries + d.deliveries,
    gasDelivered: acc.gasDelivered + d.gasDelivered,
    gasPaid: acc.gasPaid + d.gasPaid,
    revenue: acc.revenue + d.revenue,
    revenueVat: acc.revenueVat + d.revenueVat,
  }), { deliveries: 0, gasDelivered: 0, gasPaid: 0, revenue: 0, revenueVat: 0 })

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-gray-800">Tổng hợp doanh thu & sản lượng</h1>

      <div className="bg-white rounded-xl border p-4 flex flex-wrap gap-3 items-end">
        <div>
          <label className="block text-xs text-gray-500 mb-1">Từ ngày</label>
          <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)}
            className="border rounded px-3 py-2 text-sm" />
        </div>
        <div>
          <label className="block text-xs text-gray-500 mb-1">Đến ngày</label>
          <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)}
            className="border rounded px-3 py-2 text-sm" />
        </div>
        <button onClick={load} disabled={loading}
          className="bg-blue-600 text-white px-5 py-2 rounded text-sm hover:bg-blue-700 disabled:opacity-50">
          {loading ? 'Đang tải...' : 'Xem'}
        </button>
      </div>

      {data.length > 0 && (
        <>
          {/* Summary cards */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border p-4">
              <div className="text-xs text-gray-500">Tổng lần giao</div>
              <div className="text-2xl font-bold text-gray-800">{totals.deliveries.toLocaleString('vi-VN')}</div>
            </div>
            <div className="bg-white rounded-xl border p-4">
              <div className="text-xs text-gray-500">Gas thanh toán (kg)</div>
              <div className="text-2xl font-bold text-gray-800">{totals.gasPaid.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</div>
            </div>
            <div className="bg-white rounded-xl border p-4">
              <div className="text-xs text-gray-500">Doanh thu chưa VAT</div>
              <div className="text-2xl font-bold text-blue-700">{totals.revenue.toLocaleString('vi-VN')} đ</div>
            </div>
            <div className="bg-white rounded-xl border p-4">
              <div className="text-xs text-gray-500">Doanh thu có VAT 8%</div>
              <div className="text-2xl font-bold text-green-700">{Math.round(totals.revenueVat).toLocaleString('vi-VN')} đ</div>
            </div>
          </div>

          {/* Chart */}
          <div className="bg-white rounded-xl border p-5">
            <h2 className="text-sm font-semibold text-gray-700 mb-3">Doanh thu theo tháng (có VAT)</h2>
            <canvas ref={canvasRef} className="w-full" style={{ height: 300 }} />
          </div>

          {/* Table */}
          <div className="bg-white rounded-xl border overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b">
                <tr>
                  {['Tháng', 'Số lần giao', 'Gas giao (kg)', 'Gas TT (kg)', 'Doanh thu chưa VAT', 'Doanh thu có VAT'].map(h => (
                    <th key={h} className="text-left px-4 py-3 font-medium text-gray-600">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.map(d => (
                  <tr key={d.label} className="border-b last:border-0 hover:bg-gray-50">
                    <td className="px-4 py-2.5 font-medium">{d.label}</td>
                    <td className="px-4 py-2.5 text-right">{d.deliveries}</td>
                    <td className="px-4 py-2.5 text-right">{d.gasDelivered.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</td>
                    <td className="px-4 py-2.5 text-right">{d.gasPaid.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</td>
                    <td className="px-4 py-2.5 text-right">{d.revenue.toLocaleString('vi-VN')}</td>
                    <td className="px-4 py-2.5 text-right font-medium text-green-700">{Math.round(d.revenueVat).toLocaleString('vi-VN')}</td>
                  </tr>
                ))}
                <tr className="bg-gray-50 font-bold border-t-2">
                  <td className="px-4 py-2.5">TỔNG</td>
                  <td className="px-4 py-2.5 text-right">{totals.deliveries}</td>
                  <td className="px-4 py-2.5 text-right">{totals.gasDelivered.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</td>
                  <td className="px-4 py-2.5 text-right">{totals.gasPaid.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</td>
                  <td className="px-4 py-2.5 text-right">{totals.revenue.toLocaleString('vi-VN')}</td>
                  <td className="px-4 py-2.5 text-right text-green-700">{Math.round(totals.revenueVat).toLocaleString('vi-VN')}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  )
}
