import { Users, CheckCircle, Clock, AlertTriangle, BarChart2 } from 'lucide-react'
import KPICard from './KPICard'

interface Props {
  total: number
  completed: number
  inReview: number
  draft: number
  rejected: number
  avgTat: number
  completionPct: number
}

export default function ExecutiveSummary({ total, completed, inReview, draft, rejected, avgTat, completionPct }: Props) {
  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: 16 }}>
      <KPICard
        label="Total Applications"
        value={total}
        trendLabel="All time"
        icon={<Users size={16} />}
        accent="var(--primary)"
      />
      <KPICard
        label="Completed"
        value={completed}
        subLabel={`${completionPct}%`}
        trend={5}
        trendLabel={`${completionPct}% rate`}
        icon={<CheckCircle size={16} />}
        accent="#16A34A"
      />
      <KPICard
        label="In Review"
        value={inReview}
        trendLabel="Pending review"
        icon={<Clock size={16} />}
        accent="#2563EB"
      />
      <KPICard
        label="Draft"
        value={draft}
        trendLabel="Awaiting submission"
        icon={<AlertTriangle size={16} />}
        accent="#9333EA"
      />
      <KPICard
        label="Rejected"
        value={rejected}
        trendLabel="Need resubmission"
        icon={<AlertTriangle size={16} />}
        accent="#DC2626"
      />
      <KPICard
        label="Avg. TAT"
        value={avgTat}
        subLabel="days"
        trendLabel="Time to completion"
        icon={<BarChart2 size={16} />}
        accent="var(--chart-5)"
      />
    </div>
  )
}
