from pydantic import BaseModel
from decimal import Decimal
from datetime import datetime
from typing import Optional, List
import uuid


class SystemItemBase(BaseModel):
    equipment_id: Optional[uuid.UUID] = None
    section_id: Optional[uuid.UUID] = None
    item_id: Optional[str] = None
    description: Optional[str] = None
    notes: Optional[str] = None
    qty_per_room: Decimal = Decimal("0")
    msrp: Decimal = Decimal("0")
    multiplier: Decimal = Decimal("1.0")
    is_section_header: bool = False
    is_note_row: bool = False
    note_text: Optional[str] = None
    is_bold_note: bool = False
    is_ofci: bool = False
    ofci_type: Optional[str] = None
    display_order: int = 0


class SystemItemCreate(SystemItemBase):
    pass


class SystemItemUpdate(BaseModel):
    equipment_id: Optional[uuid.UUID] = None
    section_id: Optional[uuid.UUID] = None
    item_id: Optional[str] = None
    description: Optional[str] = None
    notes: Optional[str] = None
    qty_per_room: Optional[Decimal] = None
    msrp: Optional[Decimal] = None
    multiplier: Optional[Decimal] = None
    is_ofci: Optional[bool] = None
    ofci_type: Optional[str] = None
    display_order: Optional[int] = None
    is_section_header: Optional[bool] = None
    is_note_row: Optional[bool] = None
    note_text: Optional[str] = None
    is_bold_note: Optional[bool] = None


class SystemItemOut(SystemItemBase):
    id: uuid.UUID
    system_id: uuid.UUID
    last_issuance_id: Optional[uuid.UUID] = None
    old_qty: Optional[Decimal] = None
    change_status: Optional[str] = None
    # Computed
    unit_cost: Optional[Decimal] = None
    extended_cost: Optional[Decimal] = None
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


class SystemSectionBase(BaseModel):
    name: str = ""
    display_order: int = 0


class SystemSectionCreate(SystemSectionBase):
    pass


class SystemSectionOut(SystemSectionBase):
    id: uuid.UUID
    system_id: uuid.UUID
    items: List[SystemItemOut] = []

    model_config = {"from_attributes": True}


class SystemBase(BaseModel):
    name: str
    system_type: str = "room_numbers"
    room_info: Optional[str] = None
    display_order: int = 0
    is_visible: bool = True


class SystemCreate(SystemBase):
    copy_from_system_id: Optional[uuid.UUID] = None


class SystemUpdate(BaseModel):
    name: Optional[str] = None
    system_type: Optional[str] = None
    room_info: Optional[str] = None
    display_order: Optional[int] = None
    is_visible: Optional[bool] = None


class SystemOut(SystemBase):
    id: uuid.UUID
    project_id: uuid.UUID
    room_count: int
    sections: List[SystemSectionOut] = []
    items: List[SystemItemOut] = []
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


class SystemSummaryOut(BaseModel):
    """Computed summary for a system (for the Summary view)."""
    system_id: uuid.UUID
    system_name: str
    room_info: Optional[str]
    room_count: int
    equipment_subtotal: Decimal
    discount_amount: Decimal
    discounted_equipment: Decimal
    non_equipment_subtotal: Decimal
    contingency_pct: Decimal
    contingency_amount: Decimal
    system_subtotal: Decimal
    system_extended: Decimal  # subtotal × room_count
