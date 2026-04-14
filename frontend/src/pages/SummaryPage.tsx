import { useParams, Link } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { getProjectSummary } from '../lib/api'
import { ProjectSummary } from '../types'

function fmt(n: number) {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(n)
}
function pct(n: number) {
  return (n * 100).toFixed(1) + '%'
}

export default function SummaryPage() {
  const { projectId } = useParams<{ projectId: string }>()

  const { data: summary, isLoading, error } = useQuery<ProjectSummary>({
    queryKey: ['summary', projectId],
    queryFn: () => getProjectSummary(projectId!),
    enabled: !!projectId,
  })

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-64 text-gray-400">
        Loading summary...
      </div>
    )
  }

  if (error || !summary) {
    return (
      <div className="flex items-center justify-center h-64 text-red-400">
        Failed to load summary.
      </div>
    )
  }

  const { systems, totals, non_equipment_lines, settings } = summary

  return (
    <div className="p-6 max-w-7xl mx-auto space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <Link to={`/projects/${projectId}`} className="text-sm text-blue-400 hover:underline">
            ← Back to Project
          </Link>
          <h1 className="text-2xl font-bold mt-1">{summary.project_name}</h1>
          <p className="text-gray-400 text-sm">Cost Estimate Summary</p>
        </div>
        <div className="text-sm text-gray-400 space-y-1 text-right">
          <div>Discount: <span className="text-white font-medium">{pct(settings.discount_pct)}</span></div>
          <div>Contingency: <span className="text-white font-medium">{pct(settings.contingency_pct)}</span></div>
          {settings.rounding_variable !== 0 && (
            <div>Rounding: <span className="text-white font-medium">Nearest {Math.abs(10 ** (-settings.rounding_variable))}</span></div>
          )}
        </div>
      </div>

      {/* Non-equipment lines (labor breakout) */}
      {non_equipment_lines.some(l => l.mult > 0) && (
        <div className="bg-gray-800 rounded-lg p-4">
          <h2 className="text-sm font-semibold text-gray-300 mb-3 uppercase tracking-wide">Non-Equipment Multipliers</h2>
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
            {non_equipment_lines.filter(l => l.mult > 0).map((line) => (
              <div key={line.label} className="bg-gray-700 rounded p-2 text-sm">
                <div className="text-gray-400 text-xs">{line.label}</div>
                <div className="text-white font-medium">{pct(line.mult)}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Per-system table */}
      <div className="bg-gray-800 rounded-lg overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-gray-700 text-gray-300 text-xs uppercase tracking-wide">
                <th className="text-left px-4 py-3">System</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Rooms</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Equip Subtotal</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Discount</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Disc. Equip</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Non-Equip</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Contingency</th>
                <th className="text-right px-4 py-3 whitespace-nowrap">Subtotal</th>
                <th className="text-right px-4 py-3 whitespace-nowrap font-semibold text-white">Extended Cost</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-700">
              {systems.map((row) => (
                <tr key={row.system_id} className="hover:bg-gray-700/50 transition-colors">
                  <td className="px-4 py-3">
                    <div className="font-medium text-white">{row.system_name}</div>
                    {row.room_info && (
                      <div className="text-xs text-gray-400 mt-0.5">{row.room_info}</div>
                    )}
                  </td>
                  <td className="text-right px-4 py-3 text-gray-300">{row.room_count}</td>
                  <td className="text-right px-4 py-3 text-gray-300">{fmt(row.equipment_subtotal)}</td>
                  <td className="text-right px-4 py-3 text-red-400">
                    {row.discount_amount > 0 ? `(${fmt(row.discount_amount)})` : '—'}
                  </td>
                  <td className="text-right px-4 py-3 text-gray-300">{fmt(row.discounted_equipment)}</td>
                  <td className="text-right px-4 py-3 text-gray-300">{fmt(row.non_equipment_subtotal)}</td>
                  <td className="text-right px-4 py-3 text-gray-300">
                    {row.contingency_amount > 0 ? fmt(row.contingency_amount) : '—'}
                  </td>
                  <td className="text-right px-4 py-3 text-gray-300">{fmt(row.system_subtotal)}</td>
                  <td className="text-right px-4 py-3 font-semibold text-white">{fmt(row.system_extended)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-gray-700 font-semibold text-white border-t-2 border-gray-500">
                <td className="px-4 py-3" colSpan={2}>PROJECT TOTALS</td>
                <td className="text-right px-4 py-3">{fmt(totals.total_equipment_subtotal)}</td>
                <td className="text-right px-4 py-3 text-red-400">
                  {totals.total_discount > 0 ? `(${fmt(totals.total_discount)})` : '—'}
                </td>
                <td className="text-right px-4 py-3">{fmt(totals.total_discounted_equipment)}</td>
                <td className="text-right px-4 py-3">{fmt(totals.total_non_equipment)}</td>
                <td className="text-right px-4 py-3">
                  {totals.total_contingency > 0 ? fmt(totals.total_contingency) : '—'}
                </td>
                <td className="px-4 py-3" />
                <td className="text-right px-4 py-3 text-lg">{fmt(totals.total_installed_cost)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Summary cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <SummaryCard label="Total Equipment" value={fmt(totals.total_equipment_subtotal)} />
        <SummaryCard
          label="Total Discount"
          value={totals.total_discount > 0 ? `(${fmt(totals.total_discount)})` : '$0'}
          valueClass="text-red-400"
        />
        <SummaryCard label="Labor & Services" value={fmt(totals.total_non_equipment)} />
        <SummaryCard label="Total Installed Cost" value={fmt(totals.total_installed_cost)} highlight />
      </div>
    </div>
  )
}

function SummaryCard({
  label,
  value,
  valueClass = 'text-white',
  highlight = false,
}: {
  label: string
  value: string
  valueClass?: string
  highlight?: boolean
}) {
  return (
    <div className={`rounded-lg p-4 ${highlight ? 'bg-blue-600' : 'bg-gray-800'}`}>
      <div className={`text-xs uppercase tracking-wide mb-1 ${highlight ? 'text-blue-200' : 'text-gray-400'}`}>
        {label}
      </div>
      <div className={`text-xl font-bold ${highlight ? 'text-white' : valueClass}`}>{value}</div>
    </div>
  )
}
