import { useState, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import {
  ArrowLeft, Plus, Trash2, Search, ChevronDown,
  ChevronUp, AlertCircle, Save,
} from 'lucide-react'
import {
  getSystem, updateSystem, addItem, updateItem,
  deleteItem, setOfci, clearOfci, listEquipment,
} from '../lib/api'
import { System, SystemItem, Equipment } from '../types'
import { cn, formatCurrency } from '../lib/utils'

const OFCI_TYPES = ['OFE', 'OFCI', 'OFOI']

export default function SystemEditorPage() {
  const { projectId, systemId } = useParams<{ projectId: string; systemId: string }>()
  const navigate = useNavigate()
  const qc = useQueryClient()

  const { data: system, isLoading } = useQuery<System>({
    queryKey: ['system', projectId, systemId],
    queryFn: () => getSystem(projectId!, systemId!),
  })

  const updateSystemMutation = useMutation({
    mutationFn: (data: object) => updateSystem(projectId!, systemId!, data),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['system', projectId, systemId] }),
  })

  const addItemMutation = useMutation({
    mutationFn: (data: object) => addItem(projectId!, systemId!, data),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['system', projectId, systemId] }),
  })

  const updateItemMutation = useMutation({
    mutationFn: ({ itemId, data }: { itemId: string; data: object }) =>
      updateItem(projectId!, systemId!, itemId, data),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['system', projectId, systemId] }),
  })

  const deleteItemMutation = useMutation({
    mutationFn: (itemId: string) => deleteItem(projectId!, systemId!, itemId),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['system', projectId, systemId] }),
  })

  const [equipSearch, setEquipSearch] = useState('')
  const [showEquipPicker, setShowEquipPicker] = useState(false)
  const [insertAfterOrder, setInsertAfterOrder] = useState<number>(0)

  const { data: equipResults = [] } = useQuery<Equipment[]>({
    queryKey: ['equipment-search', equipSearch],
    queryFn: () => listEquipment(equipSearch, undefined, 0, 50),
    enabled: equipSearch.length > 1,
  })

  const regularItems = system?.items
    .filter(i => !i.is_section_header && !i.is_note_row)
    .sort((a, b) => a.display_order - b.display_order) ?? []

  const allItems = system?.items
    .slice()
    .sort((a, b) => a.display_order - b.display_order) ?? []

  const equipSubtotal = regularItems
    .filter(i => !i.is_ofci)
    .reduce((sum, i) => sum + (i.extended_cost ?? 0), 0)

  // Inline edit state
  const [editing, setEditing] = useState<{ id: string; field: string; value: string } | null>(null)

  const commitEdit = () => {
    if (!editing) return
    const numericFields = ['qty_per_room', 'msrp', 'multiplier']
    const value = numericFields.includes(editing.field)
      ? parseFloat(editing.value) || 0
      : editing.value
    updateItemMutation.mutate({ itemId: editing.id, data: { [editing.field]: value } })
    setEditing(null)
  }

  const handleOfci = (item: SystemItem, type: string) => {
    if (item.is_ofci && item.ofci_type === type) {
      clearOfci(projectId!, systemId!, item.id).then(() =>
        qc.invalidateQueries({ queryKey: ['system', projectId, systemId] })
      )
    } else {
      setOfci(projectId!, systemId!, item.id, type).then(() =>
        qc.invalidateQueries({ queryKey: ['system', projectId, systemId] })
      )
    }
  }

  if (isLoading) return <div className="p-6 text-muted-foreground">Loading…</div>
  if (!system) return <div className="p-6">System not found</div>

  const maxOrder = allItems.length > 0 ? Math.max(...allItems.map(i => i.display_order)) + 1 : 0

  return (
    <div className="flex h-full flex-col">
      {/* Toolbar */}
      <div className="flex items-center gap-3 border-b bg-card px-4 py-2">
        <button onClick={() => navigate(`/projects/${projectId}`)} className="rounded p-1 hover:bg-accent">
          <ArrowLeft size={16} />
        </button>
        <SystemNameEditor system={system} onSave={(name) => updateSystemMutation.mutate({ name })} />
        <div className="ml-4 flex items-center gap-1 text-sm text-muted-foreground">
          <span>{system.system_type === 'room_numbers' ? 'Rooms:' : 'Count:'}</span>
          <RoomInfoEditor
            value={system.room_info ?? ''}
            onSave={(v) => updateSystemMutation.mutate({ room_info: v })}
          />
        </div>
        <div className="ml-auto flex gap-2">
          <button
            onClick={() => {
              setInsertAfterOrder(maxOrder)
              addItemMutation.mutate({
                is_section_header: true,
                note_text: 'NEW SECTION',
                display_order: maxOrder,
              })
            }}
            className="flex items-center gap-1 rounded border px-2 py-1 text-xs hover:bg-accent"
          >
            <Plus size={12} /> Section
          </button>
          <button
            onClick={() => {
              addItemMutation.mutate({
                is_note_row: true,
                note_text: '',
                display_order: maxOrder,
              })
            }}
            className="flex items-center gap-1 rounded border px-2 py-1 text-xs hover:bg-accent"
          >
            <Plus size={12} /> Note Row
          </button>
          <button
            onClick={() => { setInsertAfterOrder(maxOrder); setShowEquipPicker(true) }}
            className="flex items-center gap-1 rounded bg-primary px-2 py-1 text-xs text-primary-foreground hover:bg-primary/90"
          >
            <Plus size={12} /> Add Equipment
          </button>
        </div>
      </div>

      {/* Equipment picker modal */}
      {showEquipPicker && (
        <EquipPicker
          search={equipSearch}
          onSearch={setEquipSearch}
          results={equipResults}
          onSelect={(equip) => {
            addItemMutation.mutate({
              equipment_id: equip.id,
              item_id: equip.item_id,
              description: equip.description,
              notes: equip.notes,
              msrp: equip.msrp,
              multiplier: equip.multiplier,
              qty_per_room: 1,
              display_order: insertAfterOrder,
            })
            setShowEquipPicker(false)
            setEquipSearch('')
          }}
          onClose={() => { setShowEquipPicker(false); setEquipSearch('') }}
        />
      )}

      {/* Grid */}
      <div className="flex-1 overflow-auto">
        <table className="w-full border-collapse text-sm">
          <thead className="sticky top-0 z-10 bg-[#1F4E79] text-white">
            <tr>
              <th className="px-2 py-2 text-left w-28">ITEM ID</th>
              <th className="px-2 py-2 text-left">DESCRIPTION</th>
              <th className="px-2 py-2 text-left w-36">MODEL</th>
              <th className="px-2 py-2 text-left w-48">NOTES</th>
              <th className="px-2 py-2 text-right w-16">QTY</th>
              <th className="px-2 py-2 text-right w-24">MSRP</th>
              <th className="px-2 py-2 text-right w-16">MULT</th>
              <th className="px-2 py-2 text-right w-24">UNIT COST</th>
              <th className="px-2 py-2 text-right w-28">EXTENDED</th>
              <th className="px-2 py-2 text-center w-24">OFCI</th>
              <th className="px-2 py-2 w-8"></th>
            </tr>
          </thead>
          <tbody>
            {allItems.map((item) => {
              if (item.is_section_header) {
                return (
                  <SectionHeaderRow
                    key={item.id}
                    item={item}
                    onRename={(text) => updateItemMutation.mutate({ itemId: item.id, data: { note_text: text } })}
                    onDelete={() => deleteItemMutation.mutate(item.id)}
                    onAddBelow={() => { setInsertAfterOrder(item.display_order + 1); setShowEquipPicker(true) }}
                  />
                )
              }
              if (item.is_note_row) {
                return (
                  <NoteRow
                    key={item.id}
                    item={item}
                    onUpdate={(text, bold) => updateItemMutation.mutate({ itemId: item.id, data: { note_text: text, is_bold_note: bold } })}
                    onDelete={() => deleteItemMutation.mutate(item.id)}
                  />
                )
              }
              return (
                <EquipmentRow
                  key={item.id}
                  item={item}
                  editing={editing}
                  onStartEdit={(id, field, value) => setEditing({ id, field, value })}
                  onEditChange={(v) => setEditing(e => e ? { ...e, value: v } : null)}
                  onCommitEdit={commitEdit}
                  onDelete={() => deleteItemMutation.mutate(item.id)}
                  onOfci={handleOfci}
                />
              )
            })}

            {/* Subtotal row */}
            <tr className="border-t-2 border-gray-300 bg-gray-50">
              <td colSpan={8} className="px-2 py-2 text-right font-semibold text-sm">
                TOTAL EQUIPMENT COST SUBTOTAL
              </td>
              <td className="px-2 py-2 text-right font-semibold">
                {formatCurrency(equipSubtotal)}
              </td>
              <td colSpan={2} />
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ── Sub-components ─────────────────────────────────────────────────────────────

function SystemNameEditor({ system, onSave }: { system: System; onSave: (name: string) => void }) {
  const [editing, setEditing] = useState(false)
  const [value, setValue] = useState(system.name)
  return editing ? (
    <input
      autoFocus
      value={value}
      onChange={e => setValue(e.target.value)}
      onBlur={() => { onSave(value); setEditing(false) }}
      onKeyDown={e => { if (e.key === 'Enter') { onSave(value); setEditing(false) } if (e.key === 'Escape') setEditing(false) }}
      className="border-b-2 border-primary bg-transparent text-lg font-bold outline-none"
    />
  ) : (
    <h1 className="cursor-pointer text-lg font-bold hover:text-primary" onClick={() => { setValue(system.name); setEditing(true) }}>
      {system.name}
    </h1>
  )
}

function RoomInfoEditor({ value, onSave }: { value: string; onSave: (v: string) => void }) {
  const [editing, setEditing] = useState(false)
  const [val, setVal] = useState(value)
  return editing ? (
    <input
      autoFocus
      value={val}
      onChange={e => setVal(e.target.value)}
      onBlur={() => { onSave(val); setEditing(false) }}
      onKeyDown={e => { if (e.key === 'Enter') { onSave(val); setEditing(false) } }}
      className="w-48 rounded border px-1 text-sm outline-none focus:ring-1 focus:ring-primary"
    />
  ) : (
    <span className="cursor-pointer underline decoration-dotted hover:text-primary" onClick={() => { setVal(value); setEditing(true) }}>
      {value || '(click to set)'}
    </span>
  )
}

function SectionHeaderRow({ item, onRename, onDelete, onAddBelow }: {
  item: SystemItem
  onRename: (text: string) => void
  onDelete: () => void
  onAddBelow: () => void
}) {
  const [editing, setEditing] = useState(false)
  const [val, setVal] = useState(item.note_text ?? '')
  return (
    <tr className="bg-[#D9E1F2]">
      <td colSpan={9} className="px-2 py-1.5">
        {editing ? (
          <input
            autoFocus
            value={val}
            onChange={e => setVal(e.target.value)}
            onBlur={() => { onRename(val); setEditing(false) }}
            onKeyDown={e => { if (e.key === 'Enter') { onRename(val); setEditing(false) } }}
            className="w-full bg-transparent font-bold outline-none"
          />
        ) : (
          <span className="cursor-pointer font-bold" onClick={() => { setVal(item.note_text ?? ''); setEditing(true) }}>
            {item.note_text || 'SECTION'}
          </span>
        )}
      </td>
      <td className="px-2 py-1.5">
        <button onClick={onAddBelow} className="mr-1 rounded p-0.5 text-xs hover:bg-white/50" title="Add equipment below">
          <Plus size={12} />
        </button>
      </td>
      <td className="px-1 py-1.5">
        <button onClick={onDelete} className="rounded p-0.5 text-destructive hover:bg-white/50">
          <Trash2 size={12} />
        </button>
      </td>
    </tr>
  )
}

function NoteRow({ item, onUpdate, onDelete }: {
  item: SystemItem
  onUpdate: (text: string, bold: boolean) => void
  onDelete: () => void
}) {
  const [editing, setEditing] = useState(false)
  const [val, setVal] = useState(item.note_text ?? '')
  return (
    <tr className={cn('border-b', item.is_bold_note ? 'font-bold' : '')}>
      <td className="px-2 py-1 text-xs text-muted-foreground italic">note</td>
      <td colSpan={8} className="px-2 py-1">
        {editing ? (
          <input
            autoFocus
            value={val}
            onChange={e => setVal(e.target.value)}
            onBlur={() => { onUpdate(val, item.is_bold_note); setEditing(false) }}
            onKeyDown={e => { if (e.key === 'Enter') { onUpdate(val, item.is_bold_note); setEditing(false) } }}
            className="w-full bg-transparent outline-none"
          />
        ) : (
          <span className="cursor-pointer" onClick={() => { setVal(item.note_text ?? ''); setEditing(true) }}>
            {item.note_text || <span className="italic text-muted-foreground">Empty note — click to edit</span>}
          </span>
        )}
      </td>
      <td colSpan={2} className="px-1 py-1">
        <button onClick={onDelete} className="rounded p-0.5 text-destructive hover:bg-accent">
          <Trash2 size={12} />
        </button>
      </td>
    </tr>
  )
}

function EquipmentRow({ item, editing, onStartEdit, onEditChange, onCommitEdit, onDelete, onOfci }: {
  item: SystemItem
  editing: { id: string; field: string; value: string } | null
  onStartEdit: (id: string, field: string, value: string) => void
  onEditChange: (v: string) => void
  onCommitEdit: () => void
  onDelete: () => void
  onOfci: (item: SystemItem, type: string) => void
}) {
  const isEditing = (field: string) => editing?.id === item.id && editing?.field === field
  const editVal = (field: string) => isEditing(field) ? editing!.value : undefined

  const rowBg = item.change_status === 'increased'
    ? 'bg-red-50'
    : item.change_status === 'decreased'
    ? 'bg-green-50'
    : item.is_ofci
    ? 'bg-yellow-50'
    : ''

  return (
    <tr className={cn('border-b hover:bg-accent/30 group', rowBg)}>
      {/* ITEM ID */}
      <td className="px-2 py-1 font-mono text-xs">
        <EditableCell
          value={item.item_id ?? ''}
          editValue={editVal('item_id')}
          onStart={() => onStartEdit(item.id, 'item_id', item.item_id ?? '')}
          onChange={onEditChange}
          onCommit={onCommitEdit}
        />
      </td>
      {/* DESCRIPTION */}
      <td className="px-2 py-1">
        <EditableCell
          value={item.description ?? ''}
          editValue={editVal('description')}
          onStart={() => onStartEdit(item.id, 'description', item.description ?? '')}
          onChange={onEditChange}
          onCommit={onCommitEdit}
        />
      </td>
      {/* MODEL (display only from equipment) */}
      <td className="px-2 py-1 text-xs text-muted-foreground">{item.item_id}</td>
      {/* NOTES */}
      <td className="px-2 py-1 text-xs">
        <EditableCell
          value={item.notes ?? ''}
          editValue={editVal('notes')}
          onStart={() => onStartEdit(item.id, 'notes', item.notes ?? '')}
          onChange={onEditChange}
          onCommit={onCommitEdit}
        />
      </td>
      {/* QTY */}
      <td className="px-2 py-1 text-right">
        {item.is_ofci ? (
          <span className="text-xs font-medium text-yellow-700">{item.ofci_type}</span>
        ) : (
          <EditableCell
            value={String(item.qty_per_room)}
            editValue={editVal('qty_per_room')}
            onStart={() => onStartEdit(item.id, 'qty_per_room', String(item.qty_per_room))}
            onChange={onEditChange}
            onCommit={onCommitEdit}
            numeric
          />
        )}
      </td>
      {/* MSRP */}
      <td className="px-2 py-1 text-right text-xs">
        {item.is_ofci ? '—' : (
          <EditableCell
            value={String(item.msrp)}
            editValue={editVal('msrp')}
            onStart={() => onStartEdit(item.id, 'msrp', String(item.msrp))}
            onChange={onEditChange}
            onCommit={onCommitEdit}
            numeric
            display={formatCurrency(item.msrp)}
          />
        )}
      </td>
      {/* MULT */}
      <td className="px-2 py-1 text-right text-xs">
        {item.is_ofci ? '—' : (
          <EditableCell
            value={String(item.multiplier)}
            editValue={editVal('multiplier')}
            onStart={() => onStartEdit(item.id, 'multiplier', String(item.multiplier))}
            onChange={onEditChange}
            onCommit={onCommitEdit}
            numeric
          />
        )}
      </td>
      {/* UNIT COST */}
      <td className="px-2 py-1 text-right text-xs">
        {item.is_ofci ? '—' : formatCurrency(item.unit_cost ?? 0)}
      </td>
      {/* EXTENDED */}
      <td className="px-2 py-1 text-right text-xs font-medium">
        {item.is_ofci ? '—' : formatCurrency(item.extended_cost ?? 0)}
      </td>
      {/* OFCI */}
      <td className="px-2 py-1 text-center">
        <div className="flex justify-center gap-0.5">
          {OFCI_TYPES.map((t) => (
            <button
              key={t}
              onClick={() => onOfci(item, t)}
              className={cn(
                'rounded px-1 py-0.5 text-[10px] font-medium transition-colors',
                item.is_ofci && item.ofci_type === t
                  ? 'bg-yellow-400 text-yellow-900'
                  : 'bg-gray-100 text-gray-500 hover:bg-yellow-100'
              )}
            >
              {t}
            </button>
          ))}
        </div>
      </td>
      {/* DELETE */}
      <td className="px-1 py-1">
        <button
          onClick={onDelete}
          className="rounded p-0.5 text-destructive opacity-0 group-hover:opacity-100 hover:bg-destructive/10 transition-opacity"
        >
          <Trash2 size={12} />
        </button>
      </td>
    </tr>
  )
}

function EditableCell({ value, editValue, onStart, onChange, onCommit, numeric = false, display }: {
  value: string
  editValue?: string
  onStart: () => void
  onChange: (v: string) => void
  onCommit: () => void
  numeric?: boolean
  display?: string
}) {
  const isEditing = editValue !== undefined
  return isEditing ? (
    <input
      autoFocus
      value={editValue}
      onChange={e => onChange(e.target.value)}
      onBlur={onCommit}
      onKeyDown={e => { if (e.key === 'Enter') onCommit() }}
      className={cn(
        'w-full rounded border border-primary px-1 outline-none',
        numeric && 'text-right'
      )}
    />
  ) : (
    <span
      onClick={onStart}
      className={cn('block cursor-text', numeric && 'text-right', !value && 'text-muted-foreground italic')}
    >
      {display ?? value || '—'}
    </span>
  )
}

function EquipPicker({ search, onSearch, results, onSelect, onClose }: {
  search: string
  onSearch: (q: string) => void
  results: Equipment[]
  onSelect: (e: Equipment) => void
  onClose: () => void
}) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
      <div className="w-full max-w-2xl rounded-xl border bg-card shadow-xl">
        <div className="flex items-center gap-2 border-b px-4 py-3">
          <Search size={16} className="text-muted-foreground" />
          <input
            autoFocus
            value={search}
            onChange={e => onSearch(e.target.value)}
            placeholder="Search equipment by MFR, model, description…"
            className="flex-1 bg-transparent outline-none"
          />
          <button onClick={onClose} className="text-muted-foreground hover:text-foreground">✕</button>
        </div>
        <div className="max-h-96 overflow-y-auto">
          {results.length === 0 && search.length > 1 && (
            <div className="p-6 text-center text-muted-foreground">No results found</div>
          )}
          {search.length <= 1 && (
            <div className="p-6 text-center text-muted-foreground">Type at least 2 characters to search</div>
          )}
          {results.map((equip) => (
            <button
              key={equip.id}
              onClick={() => onSelect(equip)}
              className="flex w-full items-center gap-4 border-b px-4 py-3 text-left hover:bg-accent"
            >
              <div className="w-28 shrink-0 font-mono text-xs text-muted-foreground">{equip.item_id}</div>
              <div className="flex-1">
                <div className="font-medium text-sm">{equip.mfr} — {equip.model}</div>
                <div className="text-xs text-muted-foreground line-clamp-1">{equip.description}</div>
              </div>
              <div className="shrink-0 text-right text-sm">
                <div className="font-medium">{formatCurrency(equip.msrp)}</div>
                <div className="text-xs text-muted-foreground">×{equip.multiplier}</div>
              </div>
            </button>
          ))}
        </div>
      </div>
    </div>
  )
}
