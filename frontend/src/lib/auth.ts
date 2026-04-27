export function getStoredToken(): string | null {
  return localStorage.getItem('access_token')
}

export function setStoredToken(token: string) {
  localStorage.setItem('access_token', token)
}

export function clearStoredToken() {
  localStorage.removeItem('access_token')
  localStorage.removeItem('current_user')
}
