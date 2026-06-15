import crypto from 'crypto'

export function buildDedupHash(row: {
  delivery_date: string
  customer_code: string
  location: string
  gas_delivered: number
  total_amount: number
}): string {
  const str = `${row.delivery_date}|${row.customer_code}|${row.location}|${row.gas_delivered}|${row.total_amount}`
  return crypto.createHash('md5').update(str).digest('hex')
}

export function formatCurrency(amount: number): string {
  return new Intl.NumberFormat('vi-VN').format(amount)
}

export function formatDate(dateStr: string): string {
  const d = new Date(dateStr)
  return `${d.getDate()}-thg ${d.getMonth() + 1}-${String(d.getFullYear()).slice(2)}`
}

export function parseExcelDate(val: unknown): string | null {
  if (!val) return null
  if (typeof val === 'number') {
    // Excel serial date
    const date = new Date((val - 25569) * 86400 * 1000)
    return date.toISOString().split('T')[0]
  }
  if (typeof val === 'string') {
    const d = new Date(val)
    if (!isNaN(d.getTime())) return d.toISOString().split('T')[0]
  }
  if (val instanceof Date) return val.toISOString().split('T')[0]
  return null
}
