import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from 'recharts'

interface Slice { name: string; value: number; fill: string }
interface Props { data: Slice[]; title: string }

const CustomTooltip = ({ active, payload }: { active?: boolean; payload?: Array<{ payload: Slice }> }) => {
  if (!active || !payload?.length) return null
  const d = payload[0].payload
  return (
    <div style={{
      background: '#fff', border: '1px solid var(--border)', borderRadius: 8,
      padding: '8px 12px', fontSize: 13, boxShadow: 'var(--shadow)',
    }}>
      <span style={{ color: d.fill, fontWeight: 600 }}>{d.name}</span>
      <span style={{ color: 'var(--txt2)', marginLeft: 6 }}>{d.value}</span>
    </div>
  )
}

export default function StatusDonut({ data, title }: Props) {
  const total = data.reduce((s, d) => s + d.value, 0)

  return (
    <div style={{
      background: 'var(--card)', border: '1px solid var(--border)',
      borderRadius: 'var(--r)', padding: '20px', boxShadow: 'var(--shadow-sm)',
      display: 'flex', flexDirection: 'column', gap: 12,
    }}>
      <h3 style={{ fontSize: 14, fontWeight: 600, color: 'var(--txt)' }}>{title}</h3>
      <div style={{ height: 200, position: 'relative' }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie data={data} cx="50%" cy="50%" innerRadius={58} outerRadius={80}
              paddingAngle={3} dataKey="value" stroke="none">
              {data.map((d, i) => <Cell key={i} fill={d.fill} />)}
            </Pie>
            <Tooltip content={<CustomTooltip />} />
          </PieChart>
        </ResponsiveContainer>
        {/* centre label */}
        <div style={{
          position: 'absolute', top: '50%', left: '50%',
          transform: 'translate(-50%, -50%)',
          textAlign: 'center', pointerEvents: 'none',
        }}>
          <div style={{ fontSize: 24, fontWeight: 700, color: 'var(--txt)', lineHeight: 1 }}>{total}</div>
          <div style={{ fontSize: 11, color: 'var(--txt3)', marginTop: 2 }}>total</div>
        </div>
      </div>
      {/* legend */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {data.map(d => (
          <div key={d.name} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12 }}>
            <span style={{ width: 10, height: 10, borderRadius: 3, background: d.fill, flexShrink: 0 }} />
            <span style={{ color: 'var(--txt2)', flex: 1 }}>{d.name}</span>
            <span style={{ fontWeight: 600, color: 'var(--txt)' }}>{d.value}</span>
            <span style={{ color: 'var(--txt3)', minWidth: 34, textAlign: 'right' }}>
              {total ? Math.round((d.value / total) * 100) : 0}%
            </span>
          </div>
        ))}
      </div>
    </div>
  )
}
