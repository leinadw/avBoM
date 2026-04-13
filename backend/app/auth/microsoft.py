"""Microsoft OAuth2 / Azure AD authentication helpers."""
import httpx
from jose import jwt, JWTError
from app.config import settings

MICROSOFT_OPENID_CONFIG_URL = (
    f"https://login.microsoftonline.com/{settings.AZURE_TENANT_ID}/v2.0/.well-known/openid-configuration"
    if settings.AZURE_TENANT_ID
    else ""
)


async def get_microsoft_public_keys() -> dict:
    """Fetch JWKS from Microsoft."""
    async with httpx.AsyncClient() as client:
        oidc_resp = await client.get(MICROSOFT_OPENID_CONFIG_URL)
        oidc_config = oidc_resp.json()
        jwks_resp = await client.get(oidc_config["jwks_uri"])
        return jwks_resp.json()


async def verify_microsoft_token(token: str) -> dict:
    """Verify a Microsoft-issued JWT and return its claims."""
    keys = await get_microsoft_public_keys()
    try:
        claims = jwt.decode(
            token,
            keys,
            algorithms=["RS256"],
            audience=settings.AZURE_CLIENT_ID,
        )
        return claims
    except JWTError as e:
        raise ValueError(f"Invalid Microsoft token: {e}")
