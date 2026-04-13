from pydantic import BaseModel
from datetime import datetime, date
from decimal import Decimal
from typing import Optional, List
import uuid


class IssuanceCreate(BaseModel):
    name: str
    issue_date: Optional[date] = None
    system_ids: List[uuid.UUID] = []


class IssuanceOut(BaseModel):
    id: uuid.UUID
    project_id: uuid.UUID
    name: str
    issue_date: Optional[date]
    created_at: datetime
    system_names: List[str] = []

    model_config = {"from_attributes": True}


class RevisionEntryOut(BaseModel):
    id: uuid.UUID
    issuance_id: uuid.UUID
    system_name: Optional[str]
    mfr: Optional[str]
    model: Optional[str]
    item_id: Optional[str]
    old_qty: Optional[Decimal]
    new_qty: Optional[Decimal]
    status: Optional[str]
    summary: Optional[str]
    created_at: datetime

    model_config = {"from_attributes": True}


class PublishRequest(BaseModel):
    system_ids: List[uuid.UUID]
    issuance_name: str
    issuance_date: Optional[date] = None
    include_notes: bool = True
    include_cost: bool = True
    include_labor_breakout: bool = True
    also_export_pdf: bool = False


class EquipmentCountRow(BaseModel):
    item_id: str
    mfr: str
    model: str
    total_qty: Decimal
