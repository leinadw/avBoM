import axios from 'axios'

const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:8000'

export const api = axios.create({
  baseURL: API_URL,
})

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('access_token')
  if (token) {
    config.headers.Authorization = `Bearer ${token}`
  }
  return config
})

api.interceptors.response.use(
  (res) => res,
  (err) => {
    if (err.response?.status === 401) {
      localStorage.removeItem('access_token')
      window.location.href = '/login'
    }
    return Promise.reject(err)
  }
)

// ── Auth ──────────────────────────────────────────────────────────────────────
export const authMicrosoft = (access_token: string) =>
  api.post('/auth/microsoft', { access_token }).then((r) => r.data)

// ── Equipment DB ──────────────────────────────────────────────────────────────
export const listEquipment = (q?: string, category?: string, skip = 0, limit = 100) =>
  api.get('/equipment/', { params: { q, category, skip, limit } }).then((r) => r.data)

export const createEquipment = (data: object) =>
  api.post('/equipment/', data).then((r) => r.data)

export const updateEquipment = (id: string, data: object) =>
  api.patch(`/equipment/${id}`, data).then((r) => r.data)

export const deleteEquipment = (id: string) =>
  api.delete(`/equipment/${id}`)

export const importEquipmentXlsx = (file: File) => {
  const fd = new FormData()
  fd.append('file', file)
  return api.post('/equipment/import-xlsx', fd).then((r) => r.data)
}

// ── Projects ──────────────────────────────────────────────────────────────────
export const listProjects = () =>
  api.get('/projects/').then((r) => r.data)

export const createProject = (data: object) =>
  api.post('/projects/', data).then((r) => r.data)

export const getProject = (id: string) =>
  api.get(`/projects/${id}`).then((r) => r.data)

export const updateProject = (id: string, data: object) =>
  api.patch(`/projects/${id}`, data).then((r) => r.data)

export const deleteProject = (id: string) =>
  api.delete(`/projects/${id}`)

export const importProjectXlsx = (file: File) => {
  const fd = new FormData()
  fd.append('file', file)
  return api.post(`/projects/import-xlsx`, fd).then((r) => r.data)
}

export const listMembers = (projectId: string) =>
  api.get(`/projects/${projectId}/members`).then((r) => r.data)

export const addMember = (projectId: string, email: string, role: string) =>
  api.post(`/projects/${projectId}/members/${encodeURIComponent(email)}`, null, { params: { role } }).then((r) => r.data)

// ── Systems ───────────────────────────────────────────────────────────────────
export const listSystems = (projectId: string) =>
  api.get(`/projects/${projectId}/systems/`).then((r) => r.data)

export const createSystem = (projectId: string, data: object) =>
  api.post(`/projects/${projectId}/systems/`, data).then((r) => r.data)

export const getSystem = (projectId: string, systemId: string) =>
  api.get(`/projects/${projectId}/systems/${systemId}`).then((r) => r.data)

export const updateSystem = (projectId: string, systemId: string, data: object) =>
  api.patch(`/projects/${projectId}/systems/${systemId}`, data).then((r) => r.data)

export const deleteSystem = (projectId: string, systemId: string) =>
  api.delete(`/projects/${projectId}/systems/${systemId}`)

// ── System Items ──────────────────────────────────────────────────────────────
export const addItem = (projectId: string, systemId: string, data: object) =>
  api.post(`/projects/${projectId}/systems/${systemId}/items`, data).then((r) => r.data)

export const updateItem = (projectId: string, systemId: string, itemId: string, data: object) =>
  api.patch(`/projects/${projectId}/systems/${systemId}/items/${itemId}`, data).then((r) => r.data)

export const deleteItem = (projectId: string, systemId: string, itemId: string) =>
  api.delete(`/projects/${projectId}/systems/${systemId}/items/${itemId}`)

export const setOfci = (projectId: string, systemId: string, itemId: string, ofci_type: string) =>
  api.post(`/projects/${projectId}/systems/${systemId}/items/${itemId}/ofci`, null, { params: { ofci_type } }).then((r) => r.data)

export const clearOfci = (projectId: string, systemId: string, itemId: string) =>
  api.delete(`/projects/${projectId}/systems/${systemId}/items/${itemId}/ofci`).then((r) => r.data)

// ── Summary ───────────────────────────────────────────────────────────────────
export const getProjectSummary = (projectId: string) =>
  api.get(`/projects/${projectId}/summary`).then((r) => r.data)

// ── Publish / Export ──────────────────────────────────────────────────────────
export const publishBom = (projectId: string, data: object) =>
  api.post(`/projects/${projectId}/publish/bom`, data, { responseType: 'blob' })

export const publishEstimate = (projectId: string, data: object) =>
  api.post(`/projects/${projectId}/publish/estimate`, data, { responseType: 'blob' })

export const getEquipmentCount = (projectId: string, system_ids: string[]) =>
  api.post(`/projects/${projectId}/equipment-count`, system_ids).then((r) => r.data)

// ── Issuances ─────────────────────────────────────────────────────────────────
export const listIssuances = (projectId: string) =>
  api.get(`/projects/${projectId}/issuances`).then((r) => r.data)

export const listRevisions = (projectId: string) =>
  api.get(`/projects/${projectId}/revisions`).then((r) => r.data)
