import VerticalCard from './VerticalCard'
import type { VerticalSummary } from '../../data/onboarding'

interface Props { verticals: VerticalSummary[] }

export default function BusinessVerticals({ verticals }: Props) {
  return (
    <div>
      <h3 style={{ fontSize: 14, fontWeight: 600, color: 'var(--txt)', marginBottom: 14 }}>
        Business Verticals
      </h3>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
        {verticals.map(v => <VerticalCard key={v.name} summary={v} />)}
      </div>
    </div>
  )
}
