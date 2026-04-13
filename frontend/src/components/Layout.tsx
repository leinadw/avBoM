import { Outlet, NavLink, useNavigate } from 'react-router-dom'
import { FolderOpen, Database, LogOut, Menu } from 'lucide-react'
import { clearStoredToken } from '../lib/auth'
import { useState } from 'react'
import { cn } from '../lib/utils'

export default function Layout() {
  const navigate = useNavigate()
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const user = JSON.parse(localStorage.getItem('current_user') || '{}')

  const handleLogout = () => {
    clearStoredToken()
    navigate('/login')
  }

  return (
    <div className="flex h-screen bg-background">
      {/* Sidebar */}
      <aside
        className={cn(
          'flex flex-col border-r bg-card transition-all duration-200',
          sidebarOpen ? 'w-56' : 'w-14'
        )}
      >
        {/* Logo */}
        <div className="flex h-14 items-center gap-2 border-b px-3">
          <button
            onClick={() => setSidebarOpen(!sidebarOpen)}
            className="rounded p-1 hover:bg-accent"
          >
            <Menu size={20} />
          </button>
          {sidebarOpen && (
            <span className="font-bold text-primary text-sm">AV BoM Tool</span>
          )}
        </div>

        {/* Nav */}
        <nav className="flex-1 space-y-1 p-2">
          <NavItem to="/projects" icon={<FolderOpen size={18} />} label="Projects" open={sidebarOpen} />
          <NavItem to="/equipment-db" icon={<Database size={18} />} label="Equipment DB" open={sidebarOpen} />
        </nav>

        {/* User */}
        <div className="border-t p-2">
          <div className={cn('flex items-center gap-2 rounded px-2 py-1', sidebarOpen && 'mb-1')}>
            <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full bg-primary text-primary-foreground text-xs font-bold">
              {(user.display_name || user.email || 'U')[0].toUpperCase()}
            </div>
            {sidebarOpen && (
              <span className="truncate text-xs text-muted-foreground">{user.display_name || user.email}</span>
            )}
          </div>
          <button
            onClick={handleLogout}
            className={cn(
              'flex w-full items-center gap-2 rounded px-2 py-1 text-sm text-muted-foreground hover:bg-accent hover:text-foreground',
              !sidebarOpen && 'justify-center'
            )}
          >
            <LogOut size={16} />
            {sidebarOpen && 'Sign out'}
          </button>
        </div>
      </aside>

      {/* Main content */}
      <main className="flex-1 overflow-auto">
        <Outlet />
      </main>
    </div>
  )
}

function NavItem({
  to, icon, label, open,
}: { to: string; icon: React.ReactNode; label: string; open: boolean }) {
  return (
    <NavLink
      to={to}
      className={({ isActive }) =>
        cn(
          'flex items-center gap-2 rounded px-2 py-2 text-sm transition-colors',
          isActive
            ? 'bg-primary text-primary-foreground'
            : 'text-muted-foreground hover:bg-accent hover:text-foreground',
          !open && 'justify-center'
        )
      }
    >
      {icon}
      {open && <span>{label}</span>}
    </NavLink>
  )
}
