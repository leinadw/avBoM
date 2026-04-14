import { useState, useEffect } from 'react'
import { useParams, Link } from 'react-router-dom'
import { useQuery, useMutation } from '@tanstack/react-query'
import { listSystems, getEquipmentCount } from '../lib/api'
import { System, EquipmentCountRow } from '../types'

export default function EquipmentReportPage() {
  const { projectId } = useParams<{ projectId: string }>()
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())
  const [autoSelected, setAutoSelected] = useState(false)
  const [results, setResults] = useState<EquipmentCountRow[] | null>(null)

  const { data: systems = [], isLoading } = useQuery<System[]>({
    queryKey: ['systems', projectId],
    queryFn: () => listSystems(projectId!),
    enabled: !!projectId,
  })

  // Auto-select all systems on first load
  useEffect(() => {
    if (systems.length > 0 && !autoSelected) {
      setSelectedIds(new Set(systems.map((s) => s.id)))
      setAutoSelected(true)
    }
  }, [systems, autoSelected])

  const countMut = useMutation({
    mutationFn: (ids: string[]) => getEquipmentCount(projectId!, ids),
    onSuccess: (data) => setResults(data),
  })

  function toggleSystem(id: string) {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
    setResults(null)
  }

  function toggleAll() {
    if (selectedIds.size === systems.length) {
      setSelectedIds(new Set())
    } else {
      setSelectedIds(new Set(systems.map((s) => s.id)))
    }
    setResults(null)
  }

  function runReport() {
    if (selectedIds.size === 0) return
    countMut.mutate(Array.from(selectedIds))
  }

  const totalQty = results?.reduce((sum, r) => sum + r.total_qty, 0) ?? 0

  return (
    <div className="p-6 max-w-5xl mx-auto space-y-6">
      {/* Header */}
      <div>
        <Link to={`/projects/${projectId}`} className="text-sm text-blue-400 hover:underline">
          ← Back to Project
        </Link>
        <h1 className="text-2xl font-bold mt-1">Equipment Count Report</h1>
        <p className="text-gray-400 text-sm">
          Aggregates total quantities across rooms for selected systems
        </p>
      </div>

      {/* System selector */}
      <div className="bg-gray-800 rounded-lg p-4 space-y-3">
        <div className="flex items-center justify-between">
          <h2 className="text-sm font-semibold text-gray-300">Select Systems</h2>
          <button
            onClick={toggleAll}
            className="text-xs text-blue-400 hover:underline"
          >
            {selectedIds.size === systems.length ? 'Deselect All' : 'Select All'}
          </button>
        </div>
        {isLoading ? (
          <div className="text-gray-400 text-sm">Loading…</div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
            {systems.map((sys) => (
              <label
                key={sys.id}
                className={`flex items-center gap-3 p-3 rounded-lg border cursor-pointer transition-colors ${
                  selectedIds.has(sys.id)
                    ? 'border-blue-500 bg-blue-900/20'
                    : 'border-gray-700 hover:border-gray-600'
                }`}
              >
                <input
                  type="checkbox"
                  checked={selectedIds.has(sys.id)}
                  onChange={() => toggleSystem(sys.id)}
                  className="accent-blue-500"
                />
                <div className="min-w-0">
                  <div className="text-sm text-white truncate">{sys.name}</div>
                  {sys.room_info && (
                    <div className="text-xs text-gray-400 truncate">{sys.room_info}</div>
                  )}
                  <div className="text-xs text-gray-500">×{sys.room_count} room{sys.room_count !== 1 ? 's' : ''}</div>
                </div>
              </label>
            ))}
          </div>
        )}
        <div className="flex items-center justify-between pt-1">
          <span className="text-sm text-gray-400">
            {selectedIds.size} of {systems.length} system{systems.length !== 1 ? 's' : ''} selected
          </span>
          <button
            onClick={runReport}
            disabled={selectedIds.size === 0 || countMut.isPending}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-500 rounded text-sm font-medium disabled:opacity-50 transition-colors"
          >
            {countMut.isPending ? 'Calculating…' : 'Generate Report'}
          </button>
        </div>
      </div>

      {/* Results */}
      {results && (
        <div className="bg-gray-800 rounded-lg overflow-hidden">
          <div className="px-4 py-3 border-b border-gray-700 flex items-center justify-between">
            <h2 className="font-semibold">Equipment Count Results</h2>
            <span className="text-sm text-gray-400">{results.length} unique items · {totalQty} total units</span>
          </div>
          {results.length === 0 ? (
            <div className="p-8 text-center text-gray-400 text-sm">No equipment items found in selected systems.</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-700 text-gray-300 text-xs uppercase tracking-wide">
                    <th className="text-left px-4 py-3">Item ID</th>
                    <th className="text-left px-4 py-3">Manufacturer</th>
                    <th className="text-left px-4 py-3">Model</th>
                    <th className="text-right px-4 py-3">Total Qty</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-700">
                  {results.map((row) => (
                    <tr key={row.item_id} className="hover:bg-gray-700/50">
                      <td className="px-4 py-2.5 text-gray-400 font-mono text-xs">{row.item_id}</td>
                      <td className="px-4 py-2.5 text-white">{row.mfr}</td>
                      <td className="px-4 py-2.5 text-white">{row.model}</td>
                      <td className="px-4 py-2.5 text-right font-semibold text-white">{row.total_qty}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-gray-700 font-semibold text-white border-t border-gray-500">
                    <td className="px-4 py-3" colSpan={3}>Total</td>
                    <td className="px-4 py-3 text-right">{totalQty}</td>
                  </tr>
                </tfoot>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
