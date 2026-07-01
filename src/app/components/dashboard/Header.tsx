import { RefreshCw } from 'lucide-react'

interface Props {
  lastUpdated: Date
  onRefresh: () => void
}

export default function Header({ lastUpdated, onRefresh }: Props) {
  const fmt = lastUpdated.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })

  return (
    <header style={{
      background: 'var(--card)',
      borderBottom: '1px solid var(--border)',
      padding: '0 24px',
      height: 58,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      position: 'sticky',
      top: 0,
      zIndex: 10,
      boxShadow: '0 1px 0 var(--border)',
    }}>
      {/* brand */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <div style={{
          width: 30, height: 30, borderRadius: 8,
          background: 'var(--primary)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: 16,
        }}>
          ♻️
        </div>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: 'var(--txt)', lineHeight: 1.2 }}>
            Recykal
          </div>
          <div style={{ fontSize: 11, color: 'var(--txt3)', lineHeight: 1.2 }}>
            Onboarding Operations
          </div>
        </div>
      </div>

      {/* right */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
        <span style={{ fontSize: 12, color: 'var(--txt3)' }}>
          Updated {fmt}
        </span>
        <button
          onClick={onRefresh}
          style={{
            display: 'flex', alignItems: 'center', gap: 6,
            background: 'var(--primary)', color: '#fff',
            border: 'none', borderRadius: 8,
            padding: '7px 14px', fontSize: 12, fontWeight: 500,
          }}
        >
          <RefreshCw size={13} />
          Refresh
        </button>
      </div>
    </header>
  )
}
