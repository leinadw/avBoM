import { Routes, Route, Navigate } from 'react-router-dom'
import { getStoredToken } from './lib/auth'
import LoginPage from './pages/LoginPage'
import ProjectsPage from './pages/ProjectsPage'
import ProjectPage from './pages/ProjectPage'
import SystemEditorPage from './pages/SystemEditorPage'
import SummaryPage from './pages/SummaryPage'
import EquipmentDbPage from './pages/EquipmentDbPage'
import IssuancesPage from './pages/IssuancesPage'
import EquipmentReportPage from './pages/EquipmentReportPage'
import ProjectSettingsPage from './pages/ProjectSettingsPage'
import Layout from './components/Layout'

function RequireAuth({ children }: { children: React.ReactNode }) {
  const token = getStoredToken()
  if (!token) return <Navigate to="/login" replace />
  return <>{children}</>
}

export default function App() {
  return (
    <Routes>
      <Route path="/login" element={<LoginPage />} />
      <Route
        path="/"
        element={
          <RequireAuth>
            <Layout />
          </RequireAuth>
        }
      >
        <Route index element={<Navigate to="/projects" replace />} />
        <Route path="projects" element={<ProjectsPage />} />
        <Route path="projects/:projectId" element={<ProjectPage />} />
        <Route path="projects/:projectId/systems/:systemId" element={<SystemEditorPage />} />
        <Route path="projects/:projectId/summary" element={<SummaryPage />} />
        <Route path="projects/:projectId/issuances" element={<IssuancesPage />} />
        <Route path="projects/:projectId/equipment-report" element={<EquipmentReportPage />} />
        <Route path="projects/:projectId/settings" element={<ProjectSettingsPage />} />
        <Route path="equipment-db" element={<EquipmentDbPage />} />
      </Route>
    </Routes>
  )
}
