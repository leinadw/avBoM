import { PublicClientApplication, Configuration, AccountInfo } from '@azure/msal-browser'

const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || '',
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || 'common'}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
}

export const msalInstance = new PublicClientApplication(msalConfig)

export const loginRequest = {
  scopes: ['User.Read', 'openid', 'profile', 'email'],
}

export async function acquireMsToken(): Promise<string> {
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length === 0) {
    await msalInstance.loginPopup(loginRequest)
  }
  const account = msalInstance.getAllAccounts()[0]
  const result = await msalInstance.acquireTokenSilent({ ...loginRequest, account })
  return result.accessToken
}

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
