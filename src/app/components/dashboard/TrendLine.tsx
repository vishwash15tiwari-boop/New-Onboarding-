import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts'

interface DataPoint { month: string; submitted: number; completed: number }
interface Props { data: DataPoint[]; title: string }

const CustomTooltip = ({ active, payload, label }: {
  active?: boolean
  payload?: Array<{ name: string; value: number; color: string }>
  label?: string
}) => {
  if (!active || !payload?.length) return null
  return (
    <div style={{
      background: '#fff', border: '1px solid var(--border)', borderRadius: 8,
      padding: '8px 12px', fontSize: 13, boxShadow: 'var(--shadow)',
    }}>
      <div style={{ fontWeight: 600, color: 'var(--txt)', marginBottom: 4 }}>{label}</div>
      {payload.map(p => (
        <div key={p.name} style={{ display: 'flex', gap: 8, color: 'var(--txt2)' }}>
          <span style={{ color: p.color, fontWeight: 600 }}>{p.name}</span>
          <span>{p.value}</span>
        </div>
      ))}
    </div>
  )
}

export default function TrendLine({ data, title }: Props) {
  return (
    <div style={{
      background: 'var(--card)', border: '1px solid var(--border)',
      borderRadius: 'var(--r)', padding: '20px', boxShadow: 'var(--shadow-sm)',
      display: 'flex', flexDirection: 'column', gap: 12,
    }}>
      <h3 style={{ fontSize: 14, fontWeight: 600, color: 'var(--txt)' }}>{title}</h3>
      <div style={{ height: 220 }}>
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={data} margin={{ top: 4, right: 16, left: -20, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: 'var(--txt3)' }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: 'var(--txt3)' }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip />} />
            <Legend iconType="circle" iconSize={8}
              wrapperStyle={{ fontSize: 12, color: 'var(--txt2)' }} />
            <Line type="monotone" dataKey="submitted" name="Submitted"
              stroke="#2563EB" strokeWidth={2} dot={false} activeDot={{ r: 4 }} />
            <Line type="monotone" dataKey="completed" name="Completed"
              stroke="#16A34A" strokeWidth={2} dot={false} activeDot={{ r: 4 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  )
}
