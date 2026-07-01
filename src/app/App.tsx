import { useMemo, useState, useCallback } from 'react'
import Header from './components/dashboard/Header'
import TabNav, { type Tab } from './components/dashboard/TabNav'
import FilterBar, { type Filters } from './components/dashboard/FilterBar'
import OverviewTab from './components/dashboard/OverviewTab'
import DataTable from './components/dashboard/DataTable'
import { RECORDS } from './data/onboarding'

function applyFilters(filters: Filters) {
  return RECORDS.filter(r => {
    if (filters.vertical !== 'All' && r.vertical !== filters.vertical) return false
    if (filters.status !== 'All' && r.status !== filters.status) return false
    if (filters.category !== 'All' && r.category !== filters.category) return false
    if (filters.search) {
      const q = filters.search.toLowerCase()
      if (
        !r.name.toLowerCase().includes(q) &&
        !r.id.toLowerCase().includes(q) &&
        !r.city.toLowerCase().includes(q) &&
        !r.rm.toLowerCase().includes(q)
      ) return false
    }
    return true
  })
}

const DEFAULT_FILTERS: Filters = { vertical: 'All', status: 'All', category: 'All', search: '' }

export default function App() {
  const [tab, setTab] = useState<Tab>('overview')
  const [filters, setFilters] = useState<Filters>(DEFAULT_FILTERS)
  const [lastUpdated, setLastUpdated] = useState(() => new Date())

  const records = useMemo(() => applyFilters(filters), [filters])

  const handleRefresh = useCallback(() => {
    setLastUpdated(new Date())
  }, [])

  return (
    <div style={{ minHeight: '100vh', background: 'var(--bg)' }}>
      <Header lastUpdated={lastUpdated} onRefresh={handleRefresh} />

      <main style={{ maxWidth: 1360, margin: '0 auto', padding: '24px 24px 48px' }}>
        <div style={{ marginBottom: 20 }}>
          <FilterBar filters={filters} onChange={setFilters} />
        </div>

        <TabNav active={tab} onChange={setTab} />

        {tab === 'overview' && <OverviewTab records={records} />}
        {tab === 'table' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <h2 style={{ fontSize: 15, fontWeight: 600, color: 'var(--txt)' }}>
                All Applications
                <span style={{ fontSize: 12, fontWeight: 400, color: 'var(--txt3)', marginLeft: 8 }}>
                  {records.length} records
                </span>
              </h2>
            </div>
            <DataTable records={records} />
          </div>
        )}
      </main>
    </div>
  )
}
