import { type CSSProperties } from 'react'
import { Search, X } from 'lucide-react'
import type { Vertical, Status, Category } from '../../data/onboarding'

export interface Filters {
  vertical: Vertical | 'All'
  status: Status | 'All'
  category: Category | 'All'
  search: string
}

interface Props {
  filters: Filters
  onChange: (f: Filters) => void
}

const selectStyle: CSSProperties = {
  border: '1px solid var(--border)',
  borderRadius: 8,
  padding: '7px 10px',
  fontSize: 13,
  color: 'var(--txt)',
  background: 'var(--card)',
  outline: 'none',
  cursor: 'pointer',
  minWidth: 130,
}

export default function FilterBar({ filters, onChange }: Props) {
  const set = (patch: Partial<Filters>) => onChange({ ...filters, ...patch })
  const clear = () => onChange({ vertical: 'All', status: 'All', category: 'All', search: '' })
  const isDirty = filters.vertical !== 'All' || filters.status !== 'All' || filters.category !== 'All' || filters.search !== ''

  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
      background: 'var(--card)', border: '1px solid var(--border)',
      borderRadius: 'var(--r)', padding: '12px 16px',
      boxShadow: 'var(--shadow-sm)',
    }}>
      {/* search */}
      <div style={{ position: 'relative', flex: '1 1 200px', minWidth: 160 }}>
        <Search size={14} style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--txt3)' }} />
        <input
          value={filters.search}
          onChange={e => set({ search: e.target.value })}
          placeholder="Search by name, ID, city…"
          style={{
            width: '100%', paddingLeft: 32, paddingRight: 10, paddingTop: 7, paddingBottom: 7,
            border: '1px solid var(--border)', borderRadius: 8, fontSize: 13,
            color: 'var(--txt)', outline: 'none', background: 'var(--card)',
          }}
        />
      </div>

      <select style={selectStyle} value={filters.vertical} onChange={e => set({ vertical: e.target.value as Filters['vertical'] })}>
        <option value="All">All Verticals</option>
        <option value="Marketplace">Marketplace</option>
        <option value="OMP">OMP</option>
        <option value="EPR">EPR</option>
      </select>

      <select style={selectStyle} value={filters.status} onChange={e => set({ status: e.target.value as Filters['status'] })}>
        <option value="All">All Statuses</option>
        <option value="Completed">Completed</option>
        <option value="In Review">In Review</option>
        <option value="Draft">Draft</option>
        <option value="Rejected">Rejected</option>
      </select>

      <select style={selectStyle} value={filters.category} onChange={e => set({ category: e.target.value as Filters['category'] })}>
        <option value="All">All Categories</option>
        <option value="Brand">Brand</option>
        <option value="Aggregator">Aggregator</option>
        <option value="Recycler">Recycler</option>
        <option value="PRO">PRO</option>
        <option value="Dismantler">Dismantler</option>
        <option value="Refurbisher">Refurbisher</option>
      </select>

      {isDirty && (
        <button
          onClick={clear}
          style={{
            display: 'flex', alignItems: 'center', gap: 4,
            border: '1px solid var(--border)', borderRadius: 8,
            padding: '7px 12px', fontSize: 12, color: 'var(--txt2)',
            background: 'transparent',
          }}
        >
          <X size={13} /> Clear
        </button>
      )}
    </div>
  )
}
