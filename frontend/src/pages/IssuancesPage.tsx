import { useParams, Link } from 'react-router-dom'
import { useQuery } from '@tanstack/react-query'
import { listIssuances, listRevisions } from '../lib/api'
import { Issuance, RevisionEntry } from '../types'

export default function IssuancesPage() {
  const { projectId } = useParams<{ projectId: string }>()

  const { data: issuances = [], isLoading: loadingIss } = useQuery<Issuance[]>({
    queryKey: ['issuances', projectId],
    queryFn: () => listIssuances(projectId!),
    enabled: !!projectId,
  })

  const { data: revisions = [], isLoading: loadingRevs } = useQuery<RevisionEntry[]>({
    queryKey: ['revisions', projectId],
    queryFn: () => listRevisions(projectId!),
    enabled: !!projectId,
  })

  const statusColor = (s?: string) => {
    if (s === 'increased') return 'text-red-400'
    if (s === 'decreased') return 'text-green-400'
    return 'text-gray-400'
  }

  const statusBadge = (s?: string) => {
    if (s === 'increased') return 'bg-red-900/50 text-red-300'
    if (s === 'decreased') return 'bg-green-900/50 text-green-300'
    return 'bg-gray-700 text-gray-300'
  }

  return (
    <div className="p-6 max-w-7xl mx-auto space-y-6">
      {/* Header */}
      <div>
        <Link to={`/projects/${projectId}`} className="text-sm text-blue-400 hover:underline">
          ← Back to Project
        </Link>
        <h1 className="text-2xl font-bold mt-1">Issuances & Revisions</h1>
        <p className="text-gray-400 text-sm">Track published BoM/Estimate exports and equipment quantity changes</p>
      </div>

      {/* Issuances list */}
      <section>
        <h2 className="text-lg font-semibold mb-3">Issuances</h2>
        {loadingIss ? (
          <div className="text-gray-400 text-sm">Loading…</div>
        ) : issuances.length === 0 ? (
          <div className="bg-gray-800 rounded-lg p-6 text-center text-gray-400 text-sm">
            No issuances yet. Publish a BoM or Estimate to create one.
          </div>
        ) : (
          <div className="space-y-3">
            {issuances.map((iss) => (
              <div key={iss.id} className="bg-gray-800 rounded-lg px-5 py-4 flex items-center justify-between">
                <div>
                  <div className="font-semibold text-white">{iss.name}</div>
                  <div className="text-xs text-gray-400 mt-0.5">
                    {iss.issue_date ? new Date(iss.issue_date).toLocaleDateString() : 'No date'} ·{' '}
                    Created {new Date(iss.created_at).toLocaleString()}
                  </div>
                  {iss.system_names.length > 0 && (
                    <div className="flex flex-wrap gap-1 mt-2">
                      {iss.system_names.map((name, i) => (
                        <span key={i} className="px-2 py-0.5 bg-gray-700 text-gray-300 text-xs rounded-full">
                          {name}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
                <div className="text-right text-sm text-gray-400">
                  {iss.system_names.length} system{iss.system_names.length !== 1 ? 's' : ''}
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      {/* Revision log */}
      <section>
        <h2 className="text-lg font-semibold mb-3">Revision Log</h2>
        {loadingRevs ? (
          <div className="text-gray-400 text-sm">Loading…</div>
        ) : revisions.length === 0 ? (
          <div className="bg-gray-800 rounded-lg p-6 text-center text-gray-400 text-sm">
            No revisions recorded yet. Changes to equipment quantities are logged when a BoM or Estimate is published.
          </div>
        ) : (
          <div className="bg-gray-800 rounded-lg overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-700 text-gray-300 text-xs uppercase tracking-wide">
                    <th className="text-left px-4 py-3">Date</th>
                    <th className="text-left px-4 py-3">System</th>
                    <th className="text-left px-4 py-3">Manufacturer</th>
                    <th className="text-left px-4 py-3">Model</th>
                    <th className="text-left px-4 py-3">Item ID</th>
                    <th className="text-right px-4 py-3">Old Qty</th>
                    <th className="text-right px-4 py-3">New Qty</th>
                    <th className="text-center px-4 py-3">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-700">
                  {revisions.map((rev) => (
                    <tr key={rev.id} className="hover:bg-gray-700/50">
                      <td className="px-4 py-2.5 text-gray-400 text-xs whitespace-nowrap">
                        {new Date(rev.created_at).toLocaleString()}
                      </td>
                      <td className="px-4 py-2.5 text-gray-300">{rev.system_name}</td>
                      <td className="px-4 py-2.5 text-white">{rev.mfr}</td>
                      <td className="px-4 py-2.5 text-white">{rev.model}</td>
                      <td className="px-4 py-2.5 text-gray-400 font-mono text-xs">{rev.item_id}</td>
                      <td className={`px-4 py-2.5 text-right ${statusColor(rev.status)}`}>
                        {rev.old_qty ?? '—'}
                      </td>
                      <td className={`px-4 py-2.5 text-right font-semibold ${statusColor(rev.status)}`}>
                        {rev.new_qty ?? '—'}
                      </td>
                      <td className="px-4 py-2.5 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-xs capitalize ${statusBadge(rev.status)}`}>
                          {rev.status || 'unknown'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </section>
    </div>
  )
}
