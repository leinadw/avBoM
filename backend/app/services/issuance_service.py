"""Issuance / revision tracking service — mirrors VBA revUp / NameRevUp logic."""
from datetime import date, datetime
from decimal import Decimal
from typing import List, Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, update
from sqlalchemy.orm import selectinload
import uuid

from app.models.system import System, SystemItem
from app.models.issuance import Issuance, IssuanceSystem, RevisionEntry
from app.models.project import Project


async def create_issuance(
    project_id: uuid.UUID,
    name: str,
    issue_date: Optional[date],
    system_ids: List[uuid.UUID],
    db: AsyncSession,
) -> Issuance:
    """Create an issuance record, update revision tracking on all selected system items."""
    issuance = Issuance(
        project_id=project_id,
        name=name,
        issue_date=issue_date or date.today(),
    )
    db.add(issuance)
    await db.flush()  # get issuance.id

    # Link systems
    for sid in system_ids:
        db.add(IssuanceSystem(issuance_id=issuance.id, system_id=sid))

    # Update each selected system's items
    sys_result = await db.execute(
        select(System)
        .where(System.id.in_(system_ids))
        .options(selectinload(System.items).selectinload(SystemItem.equipment))
    )
    systems = sys_result.scalars().all()

    for system in systems:
        for item in system.items:
            if item.is_section_header or item.is_note_row:
                continue
            # Determine change status
            old_qty = item.old_qty if item.old_qty is not None else item.qty_per_room
            new_qty = item.qty_per_room

            if old_qty != new_qty:
                status = "increased" if new_qty > old_qty else "decreased"
                # Create revision entry
                mfr = item.equipment.mfr if item.equipment else ""
                model = item.equipment.model if item.equipment else ""
                rev = RevisionEntry(
                    project_id=project_id,
                    issuance_id=issuance.id,
                    system_id=system.id,
                    system_item_id=item.id,
                    system_name=system.name,
                    mfr=mfr,
                    model=model,
                    item_id=item.item_id,
                    old_qty=old_qty,
                    new_qty=new_qty,
                    status=status,
                )
                db.add(rev)

                # Update item tracking
                item.last_issuance_id = issuance.id
                item.old_qty = new_qty
                item.change_status = status
            else:
                item.change_status = "unchanged" if item.last_issuance_id else None

    await db.commit()
    await db.refresh(issuance)
    return issuance
