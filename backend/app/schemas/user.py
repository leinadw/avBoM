from pydantic import BaseModel, EmailStr
from datetime import datetime
import uuid


class UserOut(BaseModel):
    id: uuid.UUID
    email: str
    display_name: str
    role: str
    created_at: datetime

    model_config = {"from_attributes": True}


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: UserOut


class RegisterRequest(BaseModel):
    email: EmailStr
    password: str
    display_name: str


class LoginRequest(BaseModel):
    email: EmailStr
    password: str
