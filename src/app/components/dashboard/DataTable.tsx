import { useState } from 'react'
import { ChevronUp, ChevronDown, ChevronsUpDown } from 'lucide-react'
import StatusChip from './StatusChip'
import type { OnboardingRecord } from '../../data/onboarding'

type SortKey = keyof OnboardingRecord

interface Props { records: OnboardingRecord[] }

export default function DataTable({ records }: Props) {
  const [sortKey, setSortKey] = useState<SortKey>('submittedDate')
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('desc')
  const [page, setPage] = useState(0)
  const PAGE_SIZE = 15

  const handleSort = (key: SortKey) => {
    if (sortKey === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc')
    else { setSortKey(key); setSortDir('asc') }
    setPage(0)
  }

  const sorted = [...records].sort((a, b) => {
    const av = a[sortKey] ?? ''
    const bv = b[sortKey] ?? ''
    const cmp = String(av).localeCompare(String(bv), undefined, { numeric: true })
    return sortDir === 'asc' ? cmp : -cmp
  })

  const pageData = sorted.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE)
  const totalPages = Math.ceil(sorted.length / PAGE_SIZE)

  const SortIcon = ({ col }: { col: SortKey }) => {
    if (sortKey !== col) return <ChevronsUpDown size={12} style={{ opacity: .4 }} />
    return sortDir === 'asc' ? <ChevronUp size={12} /> : <ChevronDown size={12} />
  }

  const Th = ({ col, label, align = 'left' }: { col: SortKey; label: string; align?: string }) => (
    <th
      onClick={() => handleSort(col)}
      style={{
        padding: '10px 14px', fontSize: 11, fontWeight: 600, color: 'var(--txt2)',
        textTransform: 'uppercase', letterSpacing: '.04em', textAlign: align as 'left',
        cursor: 'pointer', userSelect: 'none', background: 'var(--bg)', whiteSpace: 'nowrap',
        borderBottom: '1px solid var(--border)',
      }}
    >
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
        {label} <SortIcon col={col} />
      </span>
    </th>
  )

  const VERTICAL_COLOR: Record<string, string> = {
    Marketplace: '#DCFCE7',
    OMP: '#DBEAFE',
    EPR: '#F3E8FF',
  }
  const VERTICAL_TEXT: Record<string, string> = {
    Marketplace: '#15803D',
    OMP: '#1D4ED8',
    EPR: '#7E22CE',
  }

  return (
    <div style={{
      background: 'var(--card)', border: '1px solid var(--border)',
      borderRadius: 'var(--r)', boxShadow: 'var(--shadow-sm)', overflow: 'hidden',
    }}>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr>
              <Th col="id" label="ID" />
              <Th col="name" label="Name" />
              <Th col="vertical" label="Vertical" />
              <Th col="category" label="Category" />
              <Th col="status" label="Status" />
              <Th col="submittedDate" label="Submitted" />
              <Th col="tatDays" label="TAT (d)" align="right" />
              <Th col="rm" label="RM" />
              <Th col="city" label="City" />
            </tr>
          </thead>
          <tbody>
            {pageData.map((r, i) => (
              <tr key={r.id} style={{
                background: i % 2 === 0 ? '#fff' : 'var(--bg)',
                borderBottom: '1px solid var(--border)',
              }}>
                <td style={{ padding: '10px 14px', fontSize: 12, fontWeight: 600, color: 'var(--txt3)', whiteSpace: 'nowrap' }}>{r.id}</td>
                <td style={{ padding: '10px 14px', fontSize: 13, fontWeight: 500, color: 'var(--txt)', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.name}</td>
                <td style={{ padding: '10px 14px' }}>
                  <span style={{
                    background: VERTICAL_COLOR[r.vertical] ?? '#f0f0f0',
                    color: VERTICAL_TEXT[r.vertical] ?? '#333',
                    borderRadius: 999, padding: '2px 8px', fontSize: 11, fontWeight: 600,
                  }}>
                    {r.vertical}
                  </span>
                </td>
                <td style={{ padding: '10px 14px', fontSize: 12, color: 'var(--txt2)' }}>{r.category}</td>
                <td style={{ padding: '10px 14px' }}><StatusChip status={r.status} size="sm" /></td>
                <td style={{ padding: '10px 14px', fontSize: 12, color: 'var(--txt2)', whiteSpace: 'nowrap' }}>{r.submittedDate}</td>
                <td style={{ padding: '10px 14px', fontSize: 12, color: r.tatDays != null ? 'var(--txt)' : 'var(--txt3)', textAlign: 'right' }}>
                  {r.tatDays ?? '—'}
                </td>
                <td style={{ padding: '10px 14px', fontSize: 12, color: 'var(--txt2)', whiteSpace: 'nowrap' }}>{r.rm}</td>
                <td style={{ padding: '10px 14px', fontSize: 12, color: 'var(--txt2)' }}>{r.city}</td>
              </tr>
            ))}
            {pageData.length === 0 && (
              <tr>
                <td colSpan={9} style={{ padding: '40px', textAlign: 'center', color: 'var(--txt3)', fontSize: 13 }}>
                  No records match the current filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* pagination */}
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        padding: '12px 16px', borderTop: '1px solid var(--border)',
        fontSize: 12, color: 'var(--txt2)',
      }}>
        <span>{records.length} records · page {page + 1} of {totalPages}</span>
        <div style={{ display: 'flex', gap: 6 }}>
          {[...Array(Math.min(totalPages, 8))].map((_, i) => (
            <button key={i} onClick={() => setPage(i)} style={{
              width: 28, height: 28, borderRadius: 6,
              border: i === page ? '1px solid var(--primary)' : '1px solid var(--border)',
              background: i === page ? 'var(--primary)' : 'transparent',
              color: i === page ? '#fff' : 'var(--txt2)',
              fontSize: 12, fontWeight: i === page ? 600 : 400,
            }}>
              {i + 1}
            </button>
          ))}
          {totalPages > 8 && <span style={{ alignSelf: 'center' }}>…{totalPages}</span>}
        </div>
      </div>
    </div>
  )
}
