import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { acquireMsToken, setStoredToken } from '../lib/auth'
import { authMicrosoft } from '../lib/api'

export default function LoginPage() {
  const navigate = useNavigate()
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  const handleLogin = async () => {
    setLoading(true)
    setError('')
    try {
      const msToken = await acquireMsToken()
      const { access_token, user } = await authMicrosoft(msToken)
      setStoredToken(access_token)
      localStorage.setItem('current_user', JSON.stringify(user))
      navigate('/projects')
    } catch (e: any) {
      setError(e.message || 'Login failed')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-gradient-to-br from-primary/10 to-background">
      <div className="w-full max-w-sm rounded-xl border bg-card p-8 shadow-lg">
        <div className="mb-6 text-center">
          <h1 className="text-2xl font-bold text-primary">AV BoM Tool</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Audiovisual Equipment List &amp; Bill of Materials
          </p>
        </div>

        {error && (
          <div className="mb-4 rounded bg-destructive/10 p-3 text-sm text-destructive">
            {error}
          </div>
        )}

        <button
          onClick={handleLogin}
          disabled={loading}
          className="flex w-full items-center justify-center gap-3 rounded-lg border bg-white px-4 py-3 text-sm font-medium shadow-sm hover:bg-gray-50 disabled:opacity-50"
        >
          {/* Microsoft logo SVG */}
          <svg width="20" height="20" viewBox="0 0 21 21">
            <rect x="1" y="1" width="9" height="9" fill="#f25022" />
            <rect x="11" y="1" width="9" height="9" fill="#00a4ef" />
            <rect x="1" y="11" width="9" height="9" fill="#7fba00" />
            <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
          </svg>
          {loading ? 'Signing in…' : 'Sign in with Microsoft'}
        </button>

        <p className="mt-4 text-center text-xs text-muted-foreground">
          Use your Microsoft 365 work account
        </p>
      </div>
    </div>
  )
}
