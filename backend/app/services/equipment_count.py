"""Equipment count service — mirrors VBA EquipCountSystem logic."""
from decimal import Decimal
from typing import List
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from sqlalchemy.orm import selectinload
import uuid

from app.models.system import System, SystemItem
from app.models.equipment import Equipment
from app.schemas.issuance import EquipmentCountRow
from app.services.summary import compute_system_summary
from app.models.project import Project


async def get_equipment_count(
    project_id: uuid.UUID,
    system_ids: List[uuid.UUID],
    db: AsyncSession,
) -> List[EquipmentCountRow]:
    """
    For each selected system, tally qty_per_room × room_count per item_id.
    Returns sorted list of EquipmentCountRow (by mfr, model).
    """
    result = await db.execute(select(Project).where(Project.id == project_id))
    project = result.scalar_one_or_none()
    if not project:
        return []

    sys_result = await db.execute(
        select(System)
        .where(System.id.in_(system_ids))
        .options(selectinload(System.items).selectinload(SystemItem.equipment))
    )
    systems = sys_result.scalars().all()

    # Build a lookup of room_count per system
    totals: dict[str, dict] = {}  # item_id -> {mfr, model, total_qty}

    for system in systems:
        room_count = system.room_count
        for item in system.items:
            if item.is_section_header or item.is_note_row or item.is_ofci:
                continue
            if not item.item_id or item.qty_per_room == 0:
                continue
            key = item.item_id
            qty = item.qty_per_room * room_count
            if key not in totals:
                mfr = ""
                model = ""
                if item.equipment:
                    mfr = item.equipment.mfr
                    model = item.equipment.model
                totals[key] = {"mfr": mfr, "model": model, "total_qty": Decimal("0")}
            totals[key]["total_qty"] += qty

    rows = [
        EquipmentCountRow(
            item_id=k,
            mfr=v["mfr"],
            model=v["model"],
            total_qty=v["total_qty"],
        )
        for k, v in totals.items()
    ]
    rows.sort(key=lambda r: (r.mfr.lower(), r.model.lower()))
    return rows
