import { useState } from 'react'
import { useParams, useNavigate, Link } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import {
  Plus, Trash2, Copy, ChevronRight, BarChart2,
  FileText, ClipboardList, Settings, ArrowLeft,
} from 'lucide-react'
import {
  getProject, updateProject, listSystems,
  createSystem, deleteSystem,
} from '../lib/api'
import { Project, System } from '../types'

export default function ProjectPage() {
  const { projectId } = useParams<{ projectId: string }>()
  const navigate = useNavigate()
  const qc = useQueryClient()

  const { data: project, isLoading: projLoading } = useQuery<Project>({
    queryKey: ['project', projectId],
    queryFn: () => getProject(projectId!),
  })

  const { data: systems = [], isLoading: sysLoading } = useQuery<System[]>({
    queryKey: ['systems', projectId],
    queryFn: () => listSystems(projectId!),
  })

  const [newSysName, setNewSysName] = useState('')
  const [addingSystem, setAddingSystem] = useState(false)
  const [copyFromId, setCopyFromId] = useState<string | undefined>()
  const [editingName, setEditingName] = useState(false)
  const [projName, setProjName] = useState('')

  const createSysMutation = useMutation({
    mutationFn: (data: object) => createSystem(projectId!, data),
    onSuccess: (sys) => {
      qc.invalidateQueries({ queryKey: ['systems', projectId] })
      setAddingSystem(false)
      setNewSysName('')
      setCopyFromId(undefined)
      navigate(`/projects/${projectId}/systems/${sys.id}`)
    },
  })

  const deleteSysMutation = useMutation({
    mutationFn: (sysId: string) => deleteSystem(projectId!, sysId),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['systems', projectId] }),
  })

  const updateProjMutation = useMutation({
    mutationFn: (data: object) => updateProject(projectId!, data),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['project', projectId] })
      setEditingName(false)
    },
  })

  if (projLoading) return <div className="p-6 text-muted-foreground">Loading…</div>
  if (!project) return <div className="p-6">Project not found</div>

  return (
    <div className="p-6">
      {/* Header */}
      <div className="mb-6 flex items-center gap-3">
        <button onClick={() => navigate('/projects')} className="rounded p-1 hover:bg-accent">
          <ArrowLeft size={18} />
        </button>
        {editingName ? (
          <input
            autoFocus
            value={projName}
            onChange={(e) => setProjName(e.target.value)}
            onBlur={() => updateProjMutation.mutate({ name: projName })}
            onKeyDown={(e) => {
              if (e.key === 'Enter') updateProjMutation.mutate({ name: projName })
              if (e.key === 'Escape') setEditingName(false)
            }}
            className="text-2xl font-bold border-b-2 border-primary outline-none bg-transparent"
          />
        ) : (
          <h1
            className="text-2xl font-bold cursor-pointer hover:text-primary"
            onClick={() => { setProjName(project.name); setEditingName(true) }}
          >
            {project.name}
          </h1>
        )}
      </div>

      {/* Quick nav */}
      <div className="mb-6 flex flex-wrap gap-2">
        <NavChip to={`/projects/${projectId}/summary`} icon={<BarChart2 size={14} />} label="Summary" />
        <NavChip to={`/projects/${projectId}/issuances`} icon={<ClipboardList size={14} />} label="Issuances & Revisions" />
        <NavChip to={`/projects/${projectId}/equipment-report`} icon={<FileText size={14} />} label="Equipment Report" />
        <NavChip to={`/projects/${projectId}`} icon={<Settings size={14} />} label="Project Settings" onClick={() => {}} />
      </div>

      {/* Project Settings Summary */}
      <div className="mb-6 rounded-xl border bg-card p-4">
        <h2 className="mb-3 font-semibold text-sm text-muted-foreground uppercase tracking-wide">Project Settings</h2>
        <div className="grid grid-cols-2 gap-x-8 gap-y-1 text-sm sm:grid-cols-4">
          <SettingRow label="Discount from MSRP" value={`${(project.discount_pct * 100).toFixed(2)}%`} />
          <SettingRow label="Contingency" value={`${(project.contingency_pct * 100).toFixed(2)}%`} />
          <SettingRow label="Rounding" value={project.rounding_variable === 0 ? 'None' : `10^${project.rounding_variable}`} />
          <SettingRow label="Total Non-Equip" value={`${((project.engineering_mult + project.pm_mult + project.preinstall_mult + project.installation_mult + project.programming_mult + project.tax_mult + project.ga_mult) * 100).toFixed(2)}%`} />
        </div>
        <button
          className="mt-3 text-xs text-primary hover:underline"
          onClick={() => navigate(`/projects/${projectId}/settings`)}
        >
          Edit settings →
        </button>
      </div>

      {/* Systems */}
      <div className="flex items-center justify-between mb-3">
        <h2 className="font-semibold">AV Systems</h2>
        <button
          onClick={() => setAddingSystem(true)}
          className="flex items-center gap-1 rounded-lg bg-primary px-3 py-1.5 text-sm text-primary-foreground hover:bg-primary/90"
        >
          <Plus size={14} />
          Add System
        </button>
      </div>

      {addingSystem && (
        <div className="mb-3 rounded-lg border bg-card p-4 space-y-3">
          <div>
            <label className="text-xs text-muted-foreground">System name</label>
            <input
              autoFocus
              value={newSysName}
              onChange={(e) => setNewSysName(e.target.value)}
              placeholder="e.g. Conference Room A"
              className="mt-1 w-full rounded border px-3 py-1.5 text-sm outline-none focus:ring-2 focus:ring-primary"
            />
          </div>
          {systems.length > 0 && (
            <div>
              <label className="text-xs text-muted-foreground">Copy from existing system (optional)</label>
              <select
                value={copyFromId || ''}
                onChange={(e) => setCopyFromId(e.target.value || undefined)}
                className="mt-1 w-full rounded border px-3 py-1.5 text-sm outline-none focus:ring-2 focus:ring-primary"
              >
                <option value="">— Start from template —</option>
                {systems.map((s) => (
                  <option key={s.id} value={s.id}>{s.name}</option>
                ))}
              </select>
            </div>
          )}
          <div className="flex gap-2">
            <button
              onClick={() => {
                if (!newSysName.trim()) return
                createSysMutation.mutate({
                  name: newSysName.trim(),
                  copy_from_system_id: copyFromId,
                })
              }}
              className="rounded bg-primary px-3 py-1.5 text-sm text-primary-foreground hover:bg-primary/90"
            >
              Create
            </button>
            <button
              onClick={() => { setAddingSystem(false); setNewSysName(''); setCopyFromId(undefined) }}
              className="rounded px-3 py-1.5 text-sm hover:bg-accent"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {sysLoading ? (
        <div className="text-muted-foreground text-sm">Loading systems…</div>
      ) : systems.length === 0 ? (
        <div className="rounded-xl border-2 border-dashed p-10 text-center text-muted-foreground">
          No systems yet. Add your first AV system above.
        </div>
      ) : (
        <div className="space-y-2">
          {systems.map((s) => (
            <div
              key={s.id}
              className="group flex items-center gap-3 rounded-lg border bg-card px-4 py-3 hover:border-primary/30 hover:shadow-sm transition-all cursor-pointer"
              onClick={() => navigate(`/projects/${projectId}/systems/${s.id}`)}
            >
              <div className="flex-1">
                <div className="font-medium">{s.name}</div>
                <div className="text-xs text-muted-foreground">
                  {s.system_type === 'room_numbers' ? `Rooms: ${s.room_info || '—'}` : `Count: ${s.room_info || '1'}`}
                  {' · '}
                  {s.items.filter(i => !i.is_section_header && !i.is_note_row).length} items
                </div>
              </div>
              <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                <button
                  onClick={(e) => {
                    e.stopPropagation()
                    if (confirm(`Delete system "${s.name}"?`)) deleteSysMutation.mutate(s.id)
                  }}
                  className="rounded p-1 text-destructive hover:bg-destructive/10"
                >
                  <Trash2 size={14} />
                </button>
              </div>
              <ChevronRight size={16} className="text-muted-foreground" />
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

function NavChip({ to, icon, label, onClick }: { to: string; icon: React.ReactNode; label: string; onClick?: () => void }) {
  return (
    <Link
      to={to}
      onClick={onClick}
      className="flex items-center gap-1.5 rounded-full border px-3 py-1 text-xs hover:bg-accent transition-colors"
    >
      {icon}
      {label}
    </Link>
  )
}

function SettingRow({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <div className="text-xs text-muted-foreground">{label}</div>
      <div className="font-medium">{value}</div>
    </div>
  )
}
