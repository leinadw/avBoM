import { useState, useEffect } from 'react'
import { useParams, Link, useNavigate } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { getProject, updateProject, listMembers, addMember, deleteProject } from '../lib/api'
import { Project, ProjectSettings } from '../types'

type FormValues = {
  name: string
  discount_pct: string
  contingency_pct: string
  rounding_variable: string
  engineering_mult: string
  engineering_label: string
  pm_mult: string
  pm_label: string
  preinstall_mult: string
  preinstall_label: string
  installation_mult: string
  installation_label: string
  programming_mult: string
  programming_label: string
  tax_mult: string
  tax_label: string
  ga_mult: string
  ga_label: string
  combine_equip_nonequip: boolean
  separate_license: boolean
  include_support: boolean
  ignore_hidden_tabs: boolean
}

function projectToForm(p: Project): FormValues {
  return {
    name: p.name,
    discount_pct: String(Number(p.discount_pct) * 100),
    contingency_pct: String(Number(p.contingency_pct) * 100),
    rounding_variable: String(p.rounding_variable),
    engineering_mult: String(Number(p.engineering_mult) * 100),
    engineering_label: p.engineering_label,
    pm_mult: String(Number(p.pm_mult) * 100),
    pm_label: p.pm_label,
    preinstall_mult: String(Number(p.preinstall_mult) * 100),
    preinstall_label: p.preinstall_label,
    installation_mult: String(Number(p.installation_mult) * 100),
    installation_label: p.installation_label,
    programming_mult: String(Number(p.programming_mult) * 100),
    programming_label: p.programming_label,
    tax_mult: String(Number(p.tax_mult) * 100),
    tax_label: p.tax_label,
    ga_mult: String(Number(p.ga_mult) * 100),
    ga_label: p.ga_label,
    combine_equip_nonequip: p.combine_equip_nonequip,
    separate_license: p.separate_license,
    include_support: p.include_support,
    ignore_hidden_tabs: p.ignore_hidden_tabs,
  }
}

function formToPayload(f: FormValues) {
  return {
    name: f.name,
    settings: {
      discount_pct: (parseFloat(f.discount_pct) || 0) / 100,
      contingency_pct: (parseFloat(f.contingency_pct) || 0) / 100,
      rounding_variable: parseInt(f.rounding_variable) || 0,
      engineering_mult: (parseFloat(f.engineering_mult) || 0) / 100,
      engineering_label: f.engineering_label,
      pm_mult: (parseFloat(f.pm_mult) || 0) / 100,
      pm_label: f.pm_label,
      preinstall_mult: (parseFloat(f.preinstall_mult) || 0) / 100,
      preinstall_label: f.preinstall_label,
      installation_mult: (parseFloat(f.installation_mult) || 0) / 100,
      installation_label: f.installation_label,
      programming_mult: (parseFloat(f.programming_mult) || 0) / 100,
      programming_label: f.programming_label,
      tax_mult: (parseFloat(f.tax_mult) || 0) / 100,
      tax_label: f.tax_label,
      ga_mult: (parseFloat(f.ga_mult) || 0) / 100,
      ga_label: f.ga_label,
      combine_equip_nonequip: f.combine_equip_nonequip,
      separate_license: f.separate_license,
      include_support: f.include_support,
      ignore_hidden_tabs: f.ignore_hidden_tabs,
    },
  }
}

interface MemberRow {
  user_id: string
  email: string
  display_name?: string
  role: string
}

export default function ProjectSettingsPage() {
  const { projectId } = useParams<{ projectId: string }>()
  const navigate = useNavigate()
  const qc = useQueryClient()

  const [form, setForm] = useState<FormValues | null>(null)
  const [saveMsg, setSaveMsg] = useState('')
  const [saveError, setSaveError] = useState('')
  const [newMemberEmail, setNewMemberEmail] = useState('')
  const [newMemberRole, setNewMemberRole] = useState('editor')
  const [memberError, setMemberError] = useState('')

  const { data: project } = useQuery<Project>({
    queryKey: ['project', projectId],
    queryFn: () => getProject(projectId!),
    enabled: !!projectId,
  })

  const { data: members = [] } = useQuery<MemberRow[]>({
    queryKey: ['members', projectId],
    queryFn: () => listMembers(projectId!),
    enabled: !!projectId,
  })

  useEffect(() => {
    if (project && !form) setForm(projectToForm(project))
  }, [project])

  const updateMut = useMutation({
    mutationFn: (data: object) => updateProject(projectId!, data),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['project', projectId] })
      setSaveMsg('Settings saved.')
      setSaveError('')
      setTimeout(() => setSaveMsg(''), 3000)
    },
    onError: (e: any) => setSaveError(e.response?.data?.detail || 'Failed to save'),
  })

  const addMemberMut = useMutation({
    mutationFn: ({ email, role }: { email: string; role: string }) =>
      addMember(projectId!, email, role),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['members', projectId] })
      setNewMemberEmail('')
      setMemberError('')
    },
    onError: (e: any) => setMemberError(e.response?.data?.detail || 'Failed to add member'),
  })

  const deleteMut = useMutation({
    mutationFn: () => deleteProject(projectId!),
    onSuccess: () => navigate('/projects'),
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!form) return
    updateMut.mutate(formToPayload(form))
  }

  function setField<K extends keyof FormValues>(key: K, value: FormValues[K]) {
    setForm((f) => f ? { ...f, [key]: value } : f)
  }

  function handleAddMember(e: React.FormEvent) {
    e.preventDefault()
    if (!newMemberEmail.trim()) return
    addMemberMut.mutate({ email: newMemberEmail.trim(), role: newMemberRole })
  }

  if (!form) {
    return <div className="p-6 text-gray-400">Loading…</div>
  }

  return (
    <div className="p-6 max-w-3xl mx-auto space-y-8">
      {/* Header */}
      <div>
        <Link to={`/projects/${projectId}`} className="text-sm text-blue-400 hover:underline">
          ← Back to Project
        </Link>
        <h1 className="text-2xl font-bold mt-1">Project Settings</h1>
      </div>

      <form onSubmit={handleSubmit} className="space-y-6">
        {/* General */}
        <Section title="General">
          <Field label="Project Name">
            <input
              type="text"
              value={form.name}
              onChange={(e) => setField('name', e.target.value)}
              className="input"
              required
            />
          </Field>
        </Section>

        {/* Cost settings */}
        <Section title="Cost Settings">
          <div className="grid grid-cols-2 gap-4">
            <Field label="Discount (%)">
              <input type="number" value={form.discount_pct} step="0.1" min="0" max="100"
                onChange={(e) => setField('discount_pct', e.target.value)} className="input" />
            </Field>
            <Field label="Contingency (%)">
              <input type="number" value={form.contingency_pct} step="0.1" min="0" max="100"
                onChange={(e) => setField('contingency_pct', e.target.value)} className="input" />
            </Field>
            <Field label="Rounding Variable" hint="0 = none, -1 = tens, -2 = hundreds, -3 = thousands">
              <select value={form.rounding_variable} onChange={(e) => setField('rounding_variable', e.target.value)}
                className="input">
                <option value="0">No Rounding</option>
                <option value="-1">Nearest 10</option>
                <option value="-2">Nearest 100</option>
                <option value="-3">Nearest 1,000</option>
              </select>
            </Field>
          </div>
        </Section>

        {/* Non-equipment multipliers */}
        <Section title="Non-Equipment Multipliers">
          <div className="space-y-3">
            {(
              [
                ['engineering_mult', 'engineering_label', 'Engineering'],
                ['pm_mult', 'pm_label', 'Project Management'],
                ['preinstall_mult', 'preinstall_label', 'Pre-Install / Rack Fab'],
                ['installation_mult', 'installation_label', 'Installation'],
                ['programming_mult', 'programming_label', 'Programming'],
                ['tax_mult', 'tax_label', 'Tax'],
                ['ga_mult', 'ga_label', 'G&A'],
              ] as [keyof FormValues, keyof FormValues, string][]
            ).map(([multKey, labelKey, placeholder]) => (
              <div key={String(multKey)} className="grid grid-cols-3 gap-3 items-center">
                <div className="col-span-2">
                  <input
                    type="text"
                    value={form[labelKey] as string}
                    onChange={(e) => setField(labelKey, e.target.value)}
                    placeholder={placeholder}
                    className="input w-full"
                  />
                </div>
                <div className="flex items-center gap-1">
                  <input
                    type="number"
                    value={form[multKey] as string}
                    step="0.1"
                    min="0"
                    max="100"
                    onChange={(e) => setField(multKey, e.target.value)}
                    className="input w-full"
                  />
                  <span className="text-gray-400 text-sm">%</span>
                </div>
              </div>
            ))}
          </div>
        </Section>

        {/* Feature flags */}
        <Section title="Features">
          <div className="space-y-3">
            <Toggle
              label="Combine Equipment & Non-Equipment on Estimate"
              checked={form.combine_equip_nonequip}
              onChange={(v) => setField('combine_equip_nonequip', v)}
            />
            <Toggle
              label="Separate License Line Items"
              checked={form.separate_license}
              onChange={(v) => setField('separate_license', v)}
            />
            <Toggle
              label="Include Support / Maintenance"
              checked={form.include_support}
              onChange={(v) => setField('include_support', v)}
            />
            <Toggle
              label="Ignore Hidden System Tabs on Export"
              checked={form.ignore_hidden_tabs}
              onChange={(v) => setField('ignore_hidden_tabs', v)}
            />
          </div>
        </Section>

        {saveError && <p className="text-red-400 text-sm">{saveError}</p>}
        {saveMsg && <p className="text-green-400 text-sm">{saveMsg}</p>}

        <div className="flex justify-end">
          <button
            type="submit"
            disabled={updateMut.isPending}
            className="px-6 py-2 bg-blue-600 hover:bg-blue-500 rounded font-medium disabled:opacity-50 transition-colors"
          >
            {updateMut.isPending ? 'Saving…' : 'Save Settings'}
          </button>
        </div>
      </form>

      {/* Team members */}
      <Section title="Team Members">
        <div className="space-y-2 mb-4">
          {members.map((m) => (
            <div key={m.user_id} className="flex items-center justify-between px-3 py-2 bg-gray-750 rounded border border-gray-700">
              <div>
                <span className="text-white text-sm">{m.display_name || m.email}</span>
                {m.display_name && (
                  <span className="text-gray-400 text-xs ml-2">{m.email}</span>
                )}
              </div>
              <span className={`text-xs px-2 py-0.5 rounded-full capitalize ${
                m.role === 'owner' ? 'bg-yellow-900/50 text-yellow-300' :
                m.role === 'editor' ? 'bg-blue-900/50 text-blue-300' :
                'bg-gray-700 text-gray-300'
              }`}>{m.role}</span>
            </div>
          ))}
        </div>
        <form onSubmit={handleAddMember} className="flex gap-2">
          <input
            type="email"
            placeholder="Add by email…"
            value={newMemberEmail}
            onChange={(e) => setNewMemberEmail(e.target.value)}
            className="flex-1 input"
          />
          <select
            value={newMemberRole}
            onChange={(e) => setNewMemberRole(e.target.value)}
            className="input w-32"
          >
            <option value="viewer">Viewer</option>
            <option value="editor">Editor</option>
            <option value="owner">Owner</option>
          </select>
          <button type="submit" disabled={addMemberMut.isPending}
            className="px-3 py-2 bg-blue-600 hover:bg-blue-500 rounded text-sm disabled:opacity-50">
            Add
          </button>
        </form>
        {memberError && <p className="text-red-400 text-sm mt-1">{memberError}</p>}
      </Section>

      {/* Danger zone */}
      <div className="border border-red-900/50 rounded-lg p-5">
        <h2 className="text-red-400 font-semibold mb-2">Danger Zone</h2>
        <p className="text-gray-400 text-sm mb-4">
          Permanently delete this project and all its systems, items, and issuances. This cannot be undone.
        </p>
        <button
          onClick={() => {
            if (confirm('Delete this project permanently? This cannot be undone.')) {
              deleteMut.mutate()
            }
          }}
          disabled={deleteMut.isPending}
          className="px-4 py-2 bg-red-700 hover:bg-red-600 rounded text-sm font-medium disabled:opacity-50 transition-colors"
        >
          {deleteMut.isPending ? 'Deleting…' : 'Delete Project'}
        </button>
      </div>
    </div>
  )
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-gray-800 rounded-lg p-5 space-y-4">
      <h2 className="text-sm font-semibold text-gray-300 uppercase tracking-wide border-b border-gray-700 pb-2">
        {title}
      </h2>
      {children}
    </div>
  )
}

function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-sm text-gray-300 mb-1">{label}</label>
      {hint && <p className="text-xs text-gray-500 mb-1">{hint}</p>}
      {children}
    </div>
  )
}

function Toggle({
  label,
  checked,
  onChange,
}: {
  label: string
  checked: boolean
  onChange: (v: boolean) => void
}) {
  return (
    <label className="flex items-center gap-3 cursor-pointer">
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        onClick={() => onChange(!checked)}
        className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${
          checked ? 'bg-blue-600' : 'bg-gray-600'
        }`}
      >
        <span
          className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${
            checked ? 'translate-x-6' : 'translate-x-1'
          }`}
        />
      </button>
      <span className="text-sm text-gray-300">{label}</span>
    </label>
  )
}
