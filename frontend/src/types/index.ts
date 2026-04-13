export interface User {
  id: string
  email: string
  display_name: string
  role: 'admin' | 'user'
  created_at: string
}

export interface Equipment {
  id: string
  item_id: string
  mfr: string
  model: string
  description?: string
  notes?: string
  msrp: number
  multiplier: number
  category?: string
  created_at: string
  updated_at: string
}

export interface ProjectSettings {
  rounding_variable: number
  discount_pct: number
  contingency_pct: number
  engineering_mult: number
  pm_mult: number
  preinstall_mult: number
  installation_mult: number
  programming_mult: number
  tax_mult: number
  ga_mult: number
  engineering_label: string
  pm_label: string
  preinstall_label: string
  installation_label: string
  programming_label: string
  tax_label: string
  ga_label: string
  combine_equip_nonequip: boolean
  separate_license: boolean
  include_support: boolean
  ignore_hidden_tabs: boolean
}

export interface Project extends ProjectSettings {
  id: string
  name: string
  created_by_id: string
  created_at: string
  updated_at: string
}

export interface SystemItem {
  id: string
  system_id: string
  section_id?: string
  equipment_id?: string
  display_order: number
  item_id?: string
  description?: string
  notes?: string
  qty_per_room: number
  msrp: number
  multiplier: number
  is_section_header: boolean
  is_note_row: boolean
  note_text?: string
  is_bold_note: boolean
  is_ofci: boolean
  ofci_type?: string
  last_issuance_id?: string
  old_qty?: number
  change_status?: 'increased' | 'decreased' | 'unchanged' | 'new'
  unit_cost?: number
  extended_cost?: number
  created_at: string
  updated_at: string
}

export interface SystemSection {
  id: string
  system_id: string
  name: string
  display_order: number
  items: SystemItem[]
}

export interface System {
  id: string
  project_id: string
  name: string
  system_type: 'room_numbers' | 'system_count'
  room_info?: string
  display_order: number
  is_visible: boolean
  room_count: number
  sections: SystemSection[]
  items: SystemItem[]
  created_at: string
  updated_at: string
}

export interface SystemSummaryRow {
  system_id: string
  system_name: string
  room_info?: string
  room_count: number
  equipment_subtotal: number
  discount_amount: number
  discounted_equipment: number
  non_equipment_subtotal: number
  contingency_pct: number
  contingency_amount: number
  system_subtotal: number
  system_extended: number
}

export interface ProjectSummary {
  project_name: string
  systems: SystemSummaryRow[]
  totals: {
    total_equipment_subtotal: number
    total_discount: number
    total_discounted_equipment: number
    total_non_equipment: number
    total_contingency: number
    total_installed_cost: number
  }
  non_equipment_lines: { label: string; mult: number }[]
  settings: {
    discount_pct: number
    contingency_pct: number
    rounding_variable: number
  }
}

export interface Issuance {
  id: string
  project_id: string
  name: string
  issue_date?: string
  created_at: string
  system_names: string[]
}

export interface RevisionEntry {
  id: string
  issuance_id: string
  system_name?: string
  mfr?: string
  model?: string
  item_id?: string
  old_qty?: number
  new_qty?: number
  status?: string
  summary?: string
  created_at: string
}

export interface EquipmentCountRow {
  item_id: string
  mfr: string
  model: string
  total_qty: number
}

export interface PublishRequest {
  system_ids: string[]
  issuance_name: string
  issuance_date?: string
  include_notes: boolean
  include_cost: boolean
  include_labor_breakout: boolean
  also_export_pdf: boolean
}
