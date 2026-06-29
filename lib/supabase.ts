import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!

export const supabase = createClient(supabaseUrl, supabaseAnonKey)

export type Customer = {
  id: string
  code: string
  name: string
  address: string | null
  tax_code: string | null
  created_at: string
}

export type Transaction = {
  id: string
  delivery_date: string
  customer_code: string
  location: string
  input_key: string | null
  output_file_name: string | null
  output_sheet_name: string | null
  month: number | null
  year: number | null
  pic: string | null
  b45_delivered: number
  b45_returned: number
  b12_delivered: number
  b12_returned: number
  gas_delivered: number
  gas_returned: number
  gas_paid: number
  unit_price: number
  total_amount: number
  note: string | null
  period_start: string | null
  period_end: string | null
  dedup_hash: string
  trang_thai: string | null
  created_at: string
}
