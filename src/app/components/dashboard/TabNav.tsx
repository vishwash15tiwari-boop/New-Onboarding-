
export type Tab = 'overview' | 'table'

interface Props {
  active: Tab
  onChange: (t: Tab) => void
}

const TABS: { id: Tab; label: string }[] = [
  { id: 'overview', label: 'Overview' },
  { id: 'table', label: 'Detailed Table' },
]

export default function TabNav({ active, onChange }: Props) {
  return (
    <div style={{ display: 'flex', gap: 4, borderBottom: '2px solid var(--border)', marginBottom: 24 }}>
      {TABS.map(t => (
        <button
          key={t.id}
          onClick={() => onChange(t.id)}
          style={{
            padding: '10px 18px',
            fontSize: 13,
            fontWeight: active === t.id ? 600 : 400,
            color: active === t.id ? 'var(--primary)' : 'var(--txt2)',
            background: 'transparent',
            border: 'none',
            borderBottom: active === t.id ? '2px solid var(--primary)' : '2px solid transparent',
            marginBottom: -2,
            transition: 'color .15s, border-color .15s',
          }}
        >
          {t.label}
        </button>
      ))}
    </div>
  )
}
