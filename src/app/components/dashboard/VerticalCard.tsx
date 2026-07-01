import type { VerticalSummary } from '../../data/onboarding'

const VERTICAL_ICON: Record<string, string> = {
  Marketplace: '🏪',
  OMP: '🌐',
  EPR: '♻️',
}

const VERTICAL_COLOR: Record<string, string> = {
  Marketplace: '#2D9248',
  OMP: '#2563EB',
  EPR: '#7C3AED',
}

interface Props { summary: VerticalSummary }

export default function VerticalCard({ summary }: Props) {
  const accent = VERTICAL_COLOR[summary.name] ?? 'var(--primary)'
  const pct = summary.completionPct

  return (
    <div style={{
      background: 'var(--card)',
      border: '1px solid var(--border)',
      borderRadius: 'var(--r)',
      padding: '20px',
      boxShadow: 'var(--shadow-sm)',
      display: 'flex',
      flexDirection: 'column',
      gap: 14,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <span style={{
          width: 36, height: 36, borderRadius: 10,
          background: `${accent}18`,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: 18,
        }}>
          {VERTICAL_ICON[summary.name]}
        </span>
        <div>
          <div style={{ fontWeight: 600, fontSize: 14, color: 'var(--txt)' }}>{summary.name}</div>
          <div style={{ fontSize: 12, color: 'var(--txt3)' }}>{summary.total} total</div>
        </div>
        <div style={{ marginLeft: 'auto', textAlign: 'right' }}>
          <div style={{ fontSize: 22, fontWeight: 700, color: accent }}>{pct}%</div>
          <div style={{ fontSize: 11, color: 'var(--txt3)' }}>completion</div>
        </div>
      </div>

      {/* progress bar */}
      <div style={{ height: 6, background: 'var(--border)', borderRadius: 99, overflow: 'hidden' }}>
        <div style={{ width: `${pct}%`, height: '100%', background: accent, borderRadius: 99, transition: 'width .4s ease' }} />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 4 }}>
        {[
          { label: 'Done', value: summary.completed, color: '#16A34A' },
          { label: 'Review', value: summary.inReview, color: '#2563EB' },
          { label: 'Draft', value: summary.draft, color: '#9333EA' },
          { label: 'Rejected', value: summary.rejected, color: '#DC2626' },
        ].map(item => (
          <div key={item.label} style={{ textAlign: 'center', padding: '8px 4px', background: 'var(--bg)', borderRadius: 8 }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: item.color }}>{item.value}</div>
            <div style={{ fontSize: 10, color: 'var(--txt3)', marginTop: 1 }}>{item.label}</div>
          </div>
        ))}
      </div>

      {summary.avgTat > 0 && (
        <div style={{ fontSize: 12, color: 'var(--txt2)', display: 'flex', justifyContent: 'space-between' }}>
          <span>Avg. TAT</span>
          <span style={{ fontWeight: 600, color: 'var(--txt)' }}>{summary.avgTat} days</span>
        </div>
      )}
    </div>
  )
}
