"""Auth endpoints — Microsoft OAuth exchange."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from jose import jwt
from datetime import datetime, timedelta
import httpx

from app.db.database import get_db
from app.models.user import User, UserRole
from app.schemas.user import TokenResponse, UserOut, MicrosoftAuthRequest
from app.config import settings

router = APIRouter(prefix="/auth", tags=["auth"])


def _create_token(user_id: str) -> str:
    expire = datetime.utcnow() + timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
    return jwt.encode(
        {"sub": user_id, "exp": expire},
        settings.SECRET_KEY,
        algorithm=settings.ALGORITHM,
    )


@router.post("/microsoft", response_model=TokenResponse)
async def microsoft_login(body: MicrosoftAuthRequest, db: AsyncSession = Depends(get_db)):
    """Exchange a Microsoft access token for an app JWT."""
    # Fetch user profile from MS Graph
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {body.access_token}"},
        )
    if resp.status_code != 200:
        raise HTTPException(status_code=401, detail="Invalid Microsoft token")

    ms_user = resp.json()
    oid = ms_user.get("id", "")
    email = ms_user.get("mail") or ms_user.get("userPrincipalName", "")
    display_name = ms_user.get("displayName", email)

    # Upsert user
    result = await db.execute(select(User).where(User.azure_oid == oid))
    user = result.scalar_one_or_none()
    if not user:
        # Try by email
        result = await db.execute(select(User).where(User.email == email))
        user = result.scalar_one_or_none()

    if not user:
        user = User(email=email, display_name=display_name, azure_oid=oid)
        db.add(user)
        await db.commit()
        await db.refresh(user)
    else:
        user.azure_oid = oid
        user.display_name = display_name
        await db.commit()
        await db.refresh(user)

    token = _create_token(str(user.id))
    return TokenResponse(access_token=token, user=UserOut.model_validate(user))
