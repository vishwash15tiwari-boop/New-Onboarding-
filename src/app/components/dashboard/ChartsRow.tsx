import StatusDonut from './StatusDonut'
import CategoryBar from './CategoryBar'
import TrendLine from './TrendLine'
import type { OnboardingRecord } from '../../data/onboarding'
import { computeStatusSplit, computeCategorySplit, computeMonthlyTrend } from '../../data/onboarding'

interface Props { records: OnboardingRecord[] }

export default function ChartsRow({ records }: Props) {
  const statusData = computeStatusSplit(records)
  const categoryData = computeCategorySplit(records)
  const trendData = computeMonthlyTrend(records)

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.6fr', gap: 16 }}>
        <StatusDonut data={statusData} title="Status Distribution" />
        <CategoryBar data={categoryData} title="Applications by Category" />
      </div>
      <TrendLine data={trendData} title="Monthly Trend — Submitted vs Completed" />
    </div>
  )
}
