import { useState, useRef } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import {
  listEquipment,
  createEquipment,
  updateEquipment,
  deleteEquipment,
  importEquipmentXlsx,
} from '../lib/api'
import { Equipment } from '../types'

const CATEGORIES = ['Display', 'Audio', 'Control', 'Video', 'Signal', 'Power', 'Furniture', 'Cable', 'Other']

interface EquipForm {
  item_id: string
  mfr: string
  model: string
  description: string
  notes: string
  msrp: string
  multiplier: string
  category: string
}

const emptyForm: EquipForm = {
  item_id: '',
  mfr: '',
  model: '',
  description: '',
  notes: '',
  msrp: '',
  multiplier: '1',
  category: '',
}

export default function EquipmentDbPage() {
  const qc = useQueryClient()
  const [search, setSearch] = useState('')
  const [categoryFilter, setCategoryFilter] = useState('')
  const [modalMode, setModalMode] = useState<'add' | 'edit' | null>(null)
  const [editTarget, setEditTarget] = useState<Equipment | null>(null)
  const [form, setForm] = useState<EquipForm>(emptyForm)
  const [formError, setFormError] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)

  const { data: equipment = [], isLoading } = useQuery<Equipment[]>({
    queryKey: ['equipment', search, categoryFilter],
    queryFn: () => listEquipment(search || undefined, categoryFilter || undefined, 0, 200),
    placeholderData: (prev) => prev,
  })

  const createMut = useMutation({
    mutationFn: (data: object) => createEquipment(data),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['equipment'] })
      closeModal()
    },
    onError: (e: any) => setFormError(e.response?.data?.detail || 'Failed to save'),
  })

  const updateMut = useMutation({
    mutationFn: ({ id, data }: { id: string; data: object }) => updateEquipment(id, data),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['equipment'] })
      closeModal()
    },
    onError: (e: any) => setFormError(e.response?.data?.detail || 'Failed to save'),
  })

  const deleteMut = useMutation({
    mutationFn: (id: string) => deleteEquipment(id),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['equipment'] }),
  })

  const importMut = useMutation({
    mutationFn: (file: File) => importEquipmentXlsx(file),
    onSuccess: (data) => {
      qc.invalidateQueries({ queryKey: ['equipment'] })
      alert(`Imported ${data.created} new, ${data.updated} updated.`)
    },
    onError: (e: any) => alert(e.response?.data?.detail || 'Import failed'),
  })

  function openAdd() {
    setForm(emptyForm)
    setFormError('')
    setEditTarget(null)
    setModalMode('add')
  }

  function openEdit(eq: Equipment) {
    setForm({
      item_id: eq.item_id,
      mfr: eq.mfr,
      model: eq.model,
      description: eq.description || '',
      notes: eq.notes || '',
      msrp: String(eq.msrp),
      multiplier: String(eq.multiplier),
      category: eq.category || '',
    })
    setFormError('')
    setEditTarget(eq)
    setModalMode('edit')
  }

  function closeModal() {
    setModalMode(null)
    setEditTarget(null)
    setForm(emptyForm)
    setFormError('')
  }

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    const payload = {
      item_id: form.item_id,
      mfr: form.mfr,
      model: form.model,
      description: form.description || null,
      notes: form.notes || null,
      msrp: parseFloat(form.msrp) || 0,
      multiplier: parseFloat(form.multiplier) || 1,
      category: form.category || null,
    }
    if (modalMode === 'edit' && editTarget) {
      updateMut.mutate({ id: editTarget.id, data: payload })
    } else {
      createMut.mutate(payload)
    }
  }

  function handleDelete(eq: Equipment) {
    if (confirm(`Delete ${eq.mfr} ${eq.model}? This cannot be undone.`)) {
      deleteMut.mutate(eq.id)
    }
  }

  function handleImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) importMut.mutate(file)
    e.target.value = ''
  }

  function field(key: keyof EquipForm, value: string) {
    setForm((f) => ({ ...f, [key]: value }))
  }

  return (
    <div className="p-6 max-w-7xl mx-auto space-y-4">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold">Equipment Database</h1>
          <p className="text-gray-400 text-sm">Shared equipment catalog for all projects</p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => fileRef.current?.click()}
            disabled={importMut.isPending}
            className="px-3 py-2 bg-gray-700 hover:bg-gray-600 rounded text-sm transition-colors disabled:opacity-50"
          >
            {importMut.isPending ? 'Importing…' : 'Import XLSX'}
          </button>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={handleImport} />
          <button
            onClick={openAdd}
            className="px-3 py-2 bg-blue-600 hover:bg-blue-500 rounded text-sm font-medium transition-colors"
          >
            + Add Equipment
          </button>
        </div>
      </div>

      {/* Filters */}
      <div className="flex gap-3">
        <input
          type="text"
          placeholder="Search manufacturer, model, description…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="flex-1 bg-gray-800 border border-gray-700 rounded px-3 py-2 text-sm focus:outline-none focus:border-blue-500"
        />
        <select
          value={categoryFilter}
          onChange={(e) => setCategoryFilter(e.target.value)}
          className="bg-gray-800 border border-gray-700 rounded px-3 py-2 text-sm focus:outline-none focus:border-blue-500"
        >
          <option value="">All Categories</option>
          {CATEGORIES.map((c) => (
            <option key={c} value={c}>{c}</option>
          ))}
        </select>
      </div>

      {/* Table */}
      <div className="bg-gray-800 rounded-lg overflow-hidden">
        {isLoading ? (
          <div className="p-8 text-center text-gray-400">Loading…</div>
        ) : equipment.length === 0 ? (
          <div className="p-8 text-center text-gray-400">No equipment found.</div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-gray-700 text-gray-300 text-xs uppercase tracking-wide">
                  <th className="text-left px-4 py-3">Item ID</th>
                  <th className="text-left px-4 py-3">Manufacturer</th>
                  <th className="text-left px-4 py-3">Model</th>
                  <th className="text-left px-4 py-3">Description</th>
                  <th className="text-left px-4 py-3">Category</th>
                  <th className="text-right px-4 py-3">MSRP</th>
                  <th className="text-right px-4 py-3">Mult</th>
                  <th className="text-right px-4 py-3">Unit Cost</th>
                  <th className="px-4 py-3" />
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-700">
                {equipment.map((eq) => (
                  <tr key={eq.id} className="hover:bg-gray-700/50 transition-colors group">
                    <td className="px-4 py-2.5 text-gray-400 font-mono text-xs">{eq.item_id}</td>
                    <td className="px-4 py-2.5 font-medium text-white">{eq.mfr}</td>
                    <td className="px-4 py-2.5 text-white">{eq.model}</td>
                    <td className="px-4 py-2.5 text-gray-300 max-w-xs truncate">{eq.description}</td>
                    <td className="px-4 py-2.5">
                      {eq.category && (
                        <span className="px-2 py-0.5 bg-gray-700 text-gray-300 text-xs rounded-full">
                          {eq.category}
                        </span>
                      )}
                    </td>
                    <td className="px-4 py-2.5 text-right text-gray-300">
                      ${eq.msrp.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </td>
                    <td className="px-4 py-2.5 text-right text-gray-300">{eq.multiplier.toFixed(2)}</td>
                    <td className="px-4 py-2.5 text-right text-white font-medium">
                      ${(eq.msrp * eq.multiplier).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </td>
                    <td className="px-4 py-2.5">
                      <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity justify-end">
                        <button
                          onClick={() => openEdit(eq)}
                          className="px-2 py-1 text-xs bg-blue-700 hover:bg-blue-600 rounded"
                        >
                          Edit
                        </button>
                        <button
                          onClick={() => handleDelete(eq)}
                          className="px-2 py-1 text-xs bg-red-700 hover:bg-red-600 rounded"
                        >
                          Delete
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <div className="px-4 py-2 bg-gray-750 border-t border-gray-700 text-xs text-gray-500">
          {equipment.length} item{equipment.length !== 1 ? 's' : ''}
        </div>
      </div>

      {/* Add/Edit Modal */}
      {modalMode && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4">
          <div className="bg-gray-800 rounded-xl shadow-2xl w-full max-w-lg border border-gray-700">
            <div className="flex items-center justify-between px-6 py-4 border-b border-gray-700">
              <h2 className="text-lg font-semibold">
                {modalMode === 'add' ? 'Add Equipment' : 'Edit Equipment'}
              </h2>
              <button onClick={closeModal} className="text-gray-400 hover:text-white text-xl leading-none">×</button>
            </div>
            <form onSubmit={handleSubmit} className="p-6 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <FormField label="Item ID *" value={form.item_id} onChange={(v) => field('item_id', v)} required />
                <FormField label="Category" value={form.category} onChange={(v) => field('category', v)} list="cats" />
                <datalist id="cats">
                  {CATEGORIES.map((c) => <option key={c} value={c} />)}
                </datalist>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <FormField label="Manufacturer *" value={form.mfr} onChange={(v) => field('mfr', v)} required />
                <FormField label="Model *" value={form.model} onChange={(v) => field('model', v)} required />
              </div>
              <FormField label="Description" value={form.description} onChange={(v) => field('description', v)} />
              <FormField label="Notes" value={form.notes} onChange={(v) => field('notes', v)} />
              <div className="grid grid-cols-2 gap-4">
                <FormField label="MSRP *" value={form.msrp} onChange={(v) => field('msrp', v)} type="number" required />
                <FormField label="Multiplier *" value={form.multiplier} onChange={(v) => field('multiplier', v)} type="number" required step="0.01" />
              </div>
              {formError && <p className="text-red-400 text-sm">{formError}</p>}
              <div className="flex justify-end gap-3 pt-2">
                <button type="button" onClick={closeModal} className="px-4 py-2 bg-gray-700 hover:bg-gray-600 rounded text-sm">
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={createMut.isPending || updateMut.isPending}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 rounded text-sm font-medium disabled:opacity-50"
                >
                  {createMut.isPending || updateMut.isPending ? 'Saving…' : 'Save'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  )
}

function FormField({
  label,
  value,
  onChange,
  type = 'text',
  required,
  step,
  list,
}: {
  label: string
  value: string
  onChange: (v: string) => void
  type?: string
  required?: boolean
  step?: string
  list?: string
}) {
  return (
    <div>
      <label className="block text-xs text-gray-400 mb-1">{label}</label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        required={required}
        step={step}
        list={list}
        className="w-full bg-gray-700 border border-gray-600 rounded px-3 py-2 text-sm focus:outline-none focus:border-blue-500"
      />
    </div>
  )
}
