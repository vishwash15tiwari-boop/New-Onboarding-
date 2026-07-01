import { type ReactNode } from 'react'
import { TrendingUp, TrendingDown, Minus } from 'lucide-react'

interface Props {
  label: string
  value: number | string
  subLabel?: string
  trend?: number
  trendLabel?: string
  accent?: string
  icon?: ReactNode
}

export default function KPICard({ label, value, subLabel, trend, trendLabel, accent = 'var(--primary)', icon }: Props) {
  const trendIcon = trend == null ? null : trend > 0
    ? <TrendingUp size={13} />
    : trend < 0
    ? <TrendingDown size={13} />
    : <Minus size={13} />

  const trendColor = trend == null ? '' : trend > 0 ? '#16A34A' : trend < 0 ? '#DC2626' : 'var(--txt3)'

  return (
    <div style={{
      background: 'var(--card)',
      border: '1px solid var(--border)',
      borderRadius: 'var(--r)',
      padding: '20px 22px',
      boxShadow: 'var(--shadow-sm)',
      display: 'flex',
      flexDirection: 'column',
      gap: 8,
      minWidth: 0,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <span style={{ fontSize: 12, fontWeight: 500, color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.04em' }}>
          {label}
        </span>
        {icon && (
          <span style={{
            width: 32, height: 32, borderRadius: 8,
            background: `${accent}18`,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            color: accent, flexShrink: 0,
          }}>
            {icon}
          </span>
        )}
      </div>

      <div style={{ display: 'flex', alignItems: 'baseline', gap: 6 }}>
        <span style={{ fontSize: 28, fontWeight: 700, color: 'var(--txt)', lineHeight: 1 }}>
          {value}
        </span>
        {subLabel && (
          <span style={{ fontSize: 12, color: 'var(--txt3)' }}>{subLabel}</span>
        )}
      </div>

      {(trend != null || trendLabel) && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: trendColor || 'var(--txt3)' }}>
          {trendIcon}
          <span>{trendLabel ?? (trend != null ? `${Math.abs(trend)}%` : '')}</span>
        </div>
      )}
    </div>
  )
}
