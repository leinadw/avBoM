from pydantic import BaseModel
from decimal import Decimal
from datetime import datetime
import uuid
from typing import Optional


class EquipmentBase(BaseModel):
    item_id: str
    mfr: str
    model: str
    description: Optional[str] = None
    notes: Optional[str] = None
    msrp: Decimal = Decimal("0")
    multiplier: Decimal = Decimal("1.0")
    category: Optional[str] = None


class EquipmentCreate(EquipmentBase):
    pass


class EquipmentUpdate(BaseModel):
    mfr: Optional[str] = None
    model: Optional[str] = None
    description: Optional[str] = None
    notes: Optional[str] = None
    msrp: Optional[Decimal] = None
    multiplier: Optional[Decimal] = None
    category: Optional[str] = None


class EquipmentOut(EquipmentBase):
    id: uuid.UUID
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}
