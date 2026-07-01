export type Status = 'Completed' | 'Draft' | 'In Review' | 'Rejected'
export type Vertical = 'Marketplace' | 'OMP' | 'EPR'
export type Category = 'Brand' | 'Aggregator' | 'Recycler' | 'PRO' | 'Dismantler' | 'Refurbisher'
export type VendorType = 'Enterprise' | 'SME' | 'Individual'

export interface OnboardingRecord {
  id: string
  name: string
  vertical: Vertical
  category: Category
  vendorType: VendorType
  status: Status
  submittedDate: string
  completedDate?: string
  tatDays?: number
  rm: string
  city: string
  gstin?: string
}

export interface KPI {
  label: string
  value: number | string
  subLabel?: string
  trend?: number
  trendLabel?: string
  color?: string
}

export interface VerticalSummary {
  name: Vertical
  total: number
  completed: number
  draft: number
  inReview: number
  rejected: number
  completionPct: number
  avgTat: number
}

const NAMES = [
  'Tata Steel', 'Reliance Ind.', 'HDFC Bank', 'Infosys Ltd', 'Wipro Ltd',
  'Maruti Suzuki', 'Bajaj Auto', 'ITC Limited', 'Hindustan Unilever', 'ONGC',
  'Coal India', 'NTPC Limited', 'Power Grid', 'BPCL', 'Ultratech Cement',
  'Asian Paints', 'Sun Pharma', 'Dr Reddys', 'Cipla Ltd', 'Divi Labs',
  'Tech Mahindra', 'HCL Tech', 'TCS Ltd', 'L&T Ltd', 'Grasim Ind',
  'Hindalco', 'JSW Steel', 'SBI Life', 'Axis Bank', 'Kotak Mahindra',
  'GreenCycle Co', 'EcoWaste Pvt', 'RecyclePro', 'CleanEarth', 'PlanetSafe',
  'ZeroWaste Hub', 'CircleBack', 'RenewMat', 'Upcycle Inc', 'EcoVend',
]

const CITIES = ['Mumbai', 'Delhi', 'Bengaluru', 'Hyderabad', 'Chennai', 'Pune', 'Ahmedabad', 'Kolkata', 'Surat', 'Jaipur']
const RMS = ['Ankit Sharma', 'Priya Mehta', 'Rohan Gupta', 'Sneha Patel', 'Vikram Singh', 'Neha Joshi', 'Amit Kumar', 'Divya Rao']

const STATUSES: Status[] = ['Completed', 'Draft', 'In Review', 'Rejected']
const CATEGORIES: Category[] = ['Brand', 'Aggregator', 'Recycler', 'PRO', 'Dismantler', 'Refurbisher']
const VENDOR_TYPES: VendorType[] = ['Enterprise', 'SME', 'Individual']
const VERTICALS: Vertical[] = ['Marketplace', 'OMP', 'EPR']

function rng(seed: number) {
  let s = seed
  return () => { s = (s * 9301 + 49297) % 233280; return s / 233280 }
}

function pick<T>(arr: T[], r: () => number): T {
  return arr[Math.floor(r() * arr.length)]
}

function addDays(base: Date, days: number): string {
  const d = new Date(base)
  d.setDate(d.getDate() + days)
  return d.toISOString().slice(0, 10)
}

function genRecords(): OnboardingRecord[] {
  const r = rng(42)
  const base = new Date('2024-01-01')
  const records: OnboardingRecord[] = []

  for (let i = 0; i < 160; i++) {
    const status = pick(STATUSES.flatMap(s => {
      if (s === 'Completed') return Array(5).fill(s)
      if (s === 'In Review') return Array(2).fill(s)
      if (s === 'Draft') return Array(2).fill(s)
      return [s]
    }), r) as Status

    const dayOffset = Math.floor(r() * 540)
    const submittedDate = addDays(base, dayOffset)
    const tat = status === 'Completed' ? Math.floor(r() * 25) + 3 : undefined
    const completedDate = tat ? addDays(new Date(submittedDate), tat) : undefined

    records.push({
      id: `RK-${String(i + 1).padStart(4, '0')}`,
      name: NAMES[i % NAMES.length] + (i >= NAMES.length ? ` ${Math.floor(i / NAMES.length) + 1}` : ''),
      vertical: pick(VERTICALS.flatMap(v => v === 'Marketplace' ? Array(3).fill(v) : [v]), r) as Vertical,
      category: pick(CATEGORIES, r),
      vendorType: pick(VENDOR_TYPES, r),
      status,
      submittedDate,
      completedDate,
      tatDays: tat,
      rm: pick(RMS, r),
      city: pick(CITIES, r),
      gstin: r() > 0.3 ? `27${String(Math.floor(r() * 1e12)).padStart(12, '0')}Z` : undefined,
    })
  }
  return records
}

export const RECORDS = genRecords()

export function computeKPIs(records: OnboardingRecord[]) {
  const total = records.length
  const completed = records.filter(r => r.status === 'Completed').length
  const inReview = records.filter(r => r.status === 'In Review').length
  const draft = records.filter(r => r.status === 'Draft').length
  const rejected = records.filter(r => r.status === 'Rejected').length
  const withTat = records.filter(r => r.tatDays != null)
  const avgTat = withTat.length ? Math.round(withTat.reduce((s, r) => s + (r.tatDays ?? 0), 0) / withTat.length) : 0
  const completionPct = total ? Math.round((completed / total) * 100) : 0

  return { total, completed, inReview, draft, rejected, avgTat, completionPct }
}

export function computeVerticals(records: OnboardingRecord[]): VerticalSummary[] {
  return VERTICALS.map(v => {
    const sub = records.filter(r => r.vertical === v)
    const total = sub.length
    const completed = sub.filter(r => r.status === 'Completed').length
    const draft = sub.filter(r => r.status === 'Draft').length
    const inReview = sub.filter(r => r.status === 'In Review').length
    const rejected = sub.filter(r => r.status === 'Rejected').length
    const withTat = sub.filter(r => r.tatDays != null)
    const avgTat = withTat.length ? Math.round(withTat.reduce((s, r) => s + (r.tatDays ?? 0), 0) / withTat.length) : 0
    const completionPct = total ? Math.round((completed / total) * 100) : 0
    return { name: v, total, completed, draft, inReview, rejected, completionPct, avgTat }
  })
}

export function computeStatusSplit(records: OnboardingRecord[]) {
  const kpi = computeKPIs(records)
  return [
    { name: 'Completed', value: kpi.completed, fill: '#16A34A' },
    { name: 'In Review', value: kpi.inReview, fill: '#2563EB' },
    { name: 'Draft', value: kpi.draft, fill: '#9333EA' },
    { name: 'Rejected', value: kpi.rejected, fill: '#DC2626' },
  ]
}

export function computeCategorySplit(records: OnboardingRecord[]) {
  const fills = ['#2D9248', '#2563EB', '#9333EA', '#DC2626', '#D97706', '#0891B2']
  return CATEGORIES.map((cat, i) => ({
    name: cat,
    value: records.filter(r => r.category === cat).length,
    fill: fills[i % fills.length],
  })).filter(c => c.value > 0)
}

export function computeMonthlyTrend(records: OnboardingRecord[]) {
  const months: Record<string, { submitted: number; completed: number }> = {}
  records.forEach(r => {
    const m = r.submittedDate.slice(0, 7)
    if (!months[m]) months[m] = { submitted: 0, completed: 0 }
    months[m].submitted++
    if (r.completedDate) {
      const cm = r.completedDate.slice(0, 7)
      if (!months[cm]) months[cm] = { submitted: 0, completed: 0 }
      months[cm].completed++
    }
  })
  return Object.entries(months)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([month, v]) => ({ month: month.slice(0, 7), ...v }))
}
