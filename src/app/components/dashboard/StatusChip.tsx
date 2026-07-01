import type { Status } from '../../data/onboarding'

const CONFIG: Record<Status, { bg: string; color: string; label: string }> = {
  Completed: { bg: '#DCFCE7', color: '#15803D', label: 'Completed' },
  Draft:     { bg: '#F3E8FF', color: '#7E22CE', label: 'Draft' },
  'In Review': { bg: '#DBEAFE', color: '#1D4ED8', label: 'In Review' },
  Rejected:  { bg: '#FEE2E2', color: '#B91C1C', label: 'Rejected' },
}

interface Props { status: Status; size?: 'sm' | 'md' }

export default function StatusChip({ status, size = 'md' }: Props) {
  const c = CONFIG[status]
  return (
    <span style={{
      background: c.bg,
      color: c.color,
      borderRadius: 999,
      padding: size === 'sm' ? '2px 8px' : '3px 10px',
      fontSize: size === 'sm' ? 11 : 12,
      fontWeight: 600,
      whiteSpace: 'nowrap',
      display: 'inline-flex',
      alignItems: 'center',
    }}>
      {c.label}
    </span>
  )
}
