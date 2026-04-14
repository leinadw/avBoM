import { useState, useRef } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import { Plus, Upload, Trash2, FolderOpen, Settings } from 'lucide-react'
import { listProjects, createProject, deleteProject, importProjectXlsx } from '../lib/api'
import { Project } from '../types'
import { formatCurrency } from '../lib/utils'

export default function ProjectsPage() {
  const navigate = useNavigate()
  const qc = useQueryClient()
  const [newName, setNewName] = useState('')
  const [creating, setCreating] = useState(false)
  const importRef = useRef<HTMLInputElement>(null)

  const { data: projects = [], isLoading } = useQuery<Project[]>({
    queryKey: ['projects'],
    queryFn: listProjects,
  })

  const createMutation = useMutation({
    mutationFn: (name: string) => createProject({ name }),
    onSuccess: (proj) => {
      qc.invalidateQueries({ queryKey: ['projects'] })
      setCreating(false)
      setNewName('')
      navigate(`/projects/${proj.id}`)
    },
  })

  const deleteMutation = useMutation({
    mutationFn: deleteProject,
    onSuccess: () => qc.invalidateQueries({ queryKey: ['projects'] }),
  })

  const importMutation = useMutation({
    mutationFn: importProjectXlsx,
    onSuccess: (data) => {
      qc.invalidateQueries({ queryKey: ['projects'] })
      navigate(`/projects/${data.project_id}`)
    },
  })

  return (
    <div className="p-6">
      <div className="mb-6 flex items-center justify-between">
        <h1 className="text-2xl font-bold">Projects</h1>
        <div className="flex gap-2">
          <input
            ref={importRef}
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) importMutation.mutate(file)
            }}
          />
          <button
            onClick={() => importRef.current?.click()}
            className="flex items-center gap-2 rounded-lg border px-3 py-2 text-sm hover:bg-accent"
          >
            <Upload size={16} />
            Import .xlsx
          </button>
          <button
            onClick={() => setCreating(true)}
            className="flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm text-primary-foreground hover:bg-primary/90"
          >
            <Plus size={16} />
            New Project
          </button>
        </div>
      </div>

      {/* New project form */}
      {creating && (
        <div className="mb-4 flex items-center gap-2 rounded-lg border bg-card p-4">
          <input
            autoFocus
            placeholder="Project name…"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && newName.trim()) createMutation.mutate(newName.trim())
              if (e.key === 'Escape') setCreating(false)
            }}
            className="flex-1 rounded border px-3 py-1.5 text-sm outline-none focus:ring-2 focus:ring-primary"
          />
          <button
            onClick={() => newName.trim() && createMutation.mutate(newName.trim())}
            className="rounded bg-primary px-3 py-1.5 text-sm text-primary-foreground hover:bg-primary/90"
          >
            Create
          </button>
          <button
            onClick={() => setCreating(false)}
            className="rounded px-3 py-1.5 text-sm hover:bg-accent"
          >
            Cancel
          </button>
        </div>
      )}

      {isLoading ? (
        <div className="text-muted-foreground">Loading…</div>
      ) : projects.length === 0 ? (
        <div className="flex flex-col items-center gap-4 py-20 text-center text-muted-foreground">
          <FolderOpen size={48} strokeWidth={1} />
          <p>No projects yet. Create one or import an existing .xlsx file.</p>
        </div>
      ) : (
        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {projects.map((p) => (
            <div
              key={p.id}
              className="group relative cursor-pointer rounded-xl border bg-card p-5 shadow-sm hover:border-primary/40 hover:shadow-md transition-all"
              onClick={() => navigate(`/projects/${p.id}`)}
            >
              <div className="mb-2 flex items-start justify-between">
                <h2 className="font-semibold leading-tight">{p.name}</h2>
                <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                  <button
                    onClick={(e) => { e.stopPropagation(); navigate(`/projects/${p.id}`) }}
                    className="rounded p-1 hover:bg-accent"
                  >
                    <Settings size={14} />
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation()
                      if (confirm(`Delete "${p.name}"?`)) deleteMutation.mutate(p.id)
                    }}
                    className="rounded p-1 text-destructive hover:bg-destructive/10"
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
              </div>
              <div className="text-xs text-muted-foreground space-y-0.5">
                <div>Discount: {(p.discount_pct * 100).toFixed(1)}%</div>
                <div>Contingency: {(p.contingency_pct * 100).toFixed(1)}%</div>
                <div className="mt-2 text-[11px]">
                  Updated {new Date(p.updated_at).toLocaleDateString()}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
