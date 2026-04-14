from pydantic import BaseModel
from decimal import Decimal
from datetime import datetime
from typing import Optional
import uuid


class ProjectSettings(BaseModel):
    rounding_variable: int = 0
    discount_pct: Decimal = Decimal("0")
    contingency_pct: Decimal = Decimal("0")
    engineering_mult: Decimal = Decimal("0.02")
    pm_mult: Decimal = Decimal("0.02")
    preinstall_mult: Decimal = Decimal("0.02")
    installation_mult: Decimal = Decimal("0.08")
    programming_mult: Decimal = Decimal("0.08")
    tax_mult: Decimal = Decimal("0.08275")
    ga_mult: Decimal = Decimal("0.10")
    engineering_label: str = "ENGINEERING"
    pm_label: str = "PROJECT MANAGEMENT"
    preinstall_label: str = "PRE-INSTALLATION & RACK FABRICATION"
    installation_label: str = "INSTALLATION"
    programming_label: str = "PROGRAMMING"
    tax_label: str = "TAX"
    ga_label: str = "G&A"
    combine_equip_nonequip: bool = False
    separate_license: bool = False
    include_support: bool = False
    ignore_hidden_tabs: bool = False


class ProjectCreate(BaseModel):
    name: str
    settings: Optional[ProjectSettings] = None


class ProjectUpdate(BaseModel):
    name: Optional[str] = None
    settings: Optional[ProjectSettings] = None


class ProjectOut(BaseModel):
    id: uuid.UUID
    name: str
    created_by_id: uuid.UUID
    rounding_variable: int
    discount_pct: Decimal
    contingency_pct: Decimal
    engineering_mult: Decimal
    pm_mult: Decimal
    preinstall_mult: Decimal
    installation_mult: Decimal
    programming_mult: Decimal
    tax_mult: Decimal
    ga_mult: Decimal
    engineering_label: str
    pm_label: str
    preinstall_label: str
    installation_label: str
    programming_label: str
    tax_label: str
    ga_label: str
    combine_equip_nonequip: bool
    separate_license: bool
    include_support: bool
    ignore_hidden_tabs: bool
    created_at: datetime
    updated_at: datetime

    model_config = {"from_attributes": True}


class ProjectMemberOut(BaseModel):
    user_id: uuid.UUID
    email: str
    display_name: str
    role: str

    model_config = {"from_attributes": True}
