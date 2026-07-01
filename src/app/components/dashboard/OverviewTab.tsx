import ExecutiveSummary from './ExecutiveSummary'
import BusinessVerticals from './BusinessVerticals'
import ChartsRow from './ChartsRow'
import type { OnboardingRecord } from '../../data/onboarding'
import { computeKPIs, computeVerticals } from '../../data/onboarding'

interface Props { records: OnboardingRecord[] }

export default function OverviewTab({ records }: Props) {
  const kpi = computeKPIs(records)
  const verticals = computeVerticals(records)

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 28 }}>
      <ExecutiveSummary {...kpi} />
      <BusinessVerticals verticals={verticals} />
      <ChartsRow records={records} />
    </div>
  )
}
