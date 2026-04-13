"""System tab CRUD — add/delete/rename systems, manage items."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from sqlalchemy.orm import selectinload
from typing import List
import uuid

from app.db.database import get_db
from app.models.user import User
from app.models.project import Project, ProjectMember
from app.models.system import System, SystemSection, SystemItem
from app.models.equipment import Equipment
from app.schemas.system import (
    SystemCreate, SystemUpdate, SystemOut,
    SystemItemCreate, SystemItemUpdate, SystemItemOut,
    SystemSectionCreate, SystemSectionOut,
    SystemSummaryOut,
)
from app.auth.deps import get_current_user
from app.services.summary import get_project_summary, compute_system_summary

router = APIRouter(prefix="/projects/{project_id}/systems", tags=["systems"])


async def _get_project_member(project_id: uuid.UUID, user: User, db: AsyncSession, min_role: str = "viewer"):
    result = await db.execute(
        select(ProjectMember).where(
            ProjectMember.project_id == project_id,
            ProjectMember.user_id == user.id,
        )
    )
    member = result.scalar_one_or_none()
    if not member:
        raise HTTPException(status_code=403, detail="Access denied")
    roles_order = {"viewer": 0, "editor": 1, "owner": 2}
    if roles_order.get(member.role.value, 0) < roles_order.get(min_role, 0):
        raise HTTPException(status_code=403, detail="Insufficient permissions")
    return member


async def _load_system(system_id: uuid.UUID, db: AsyncSession) -> System:
    result = await db.execute(
        select(System)
        .where(System.id == system_id)
        .options(
            selectinload(System.sections),
            selectinload(System.items).selectinload(SystemItem.equipment),
        )
    )
    system = result.scalar_one_or_none()
    if not system:
        raise HTTPException(status_code=404, detail="System not found")
    return system


def _enrich_item(item: SystemItem) -> dict:
    d = {c.name: getattr(item, c.name) for c in item.__table__.columns}
    if not item.is_section_header and not item.is_note_row and not item.is_ofci:
        unit_cost = item.msrp * item.multiplier
        d["unit_cost"] = unit_cost
        d["extended_cost"] = unit_cost * item.qty_per_room
    else:
        d["unit_cost"] = None
        d["extended_cost"] = None
    return d


# ── Systems ────────────────────────────────────────────────────────────────────

@router.get("/", response_model=List[SystemOut])
async def list_systems(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db)
    result = await db.execute(
        select(System)
        .where(System.project_id == project_id)
        .options(
            selectinload(System.sections),
            selectinload(System.items).selectinload(SystemItem.equipment),
        )
        .order_by(System.display_order)
    )
    systems = result.scalars().all()
    out = []
    for s in systems:
        d = SystemOut.model_validate(s)
        d.room_count = s.room_count
        out.append(d)
    return out


@router.post("/", response_model=SystemOut, status_code=201)
async def create_system(
    project_id: uuid.UUID,
    body: SystemCreate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")

    # Check name uniqueness
    existing = await db.execute(
        select(System).where(System.project_id == project_id, System.name == body.name)
    )
    if existing.scalar_one_or_none():
        raise HTTPException(status_code=409, detail="System name already in use")

    system = System(
        project_id=project_id,
        name=body.name,
        system_type=body.system_type,
        room_info=body.room_info,
        display_order=body.display_order,
    )

    # Copy from existing system if requested
    if body.copy_from_system_id:
        source = await _load_system(body.copy_from_system_id, db)
        db.add(system)
        await db.flush()
        for src_item in source.items:
            new_item = SystemItem(
                system_id=system.id,
                equipment_id=src_item.equipment_id,
                display_order=src_item.display_order,
                item_id=src_item.item_id,
                description=src_item.description,
                notes=src_item.notes,
                qty_per_room=src_item.qty_per_room,
                msrp=src_item.msrp,
                multiplier=src_item.multiplier,
                is_section_header=src_item.is_section_header,
                is_note_row=src_item.is_note_row,
                note_text=src_item.note_text,
                is_bold_note=src_item.is_bold_note,
            )
            db.add(new_item)
    else:
        db.add(system)
        await db.flush()

    await db.commit()
    return await _load_system(system.id, db)


@router.get("/{system_id}", response_model=SystemOut)
async def get_system(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db)
    return await _load_system(system_id, db)


@router.patch("/{system_id}", response_model=SystemOut)
async def update_system(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    body: SystemUpdate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")
    system = await _load_system(system_id, db)
    for field, val in body.model_dump(exclude_unset=True).items():
        setattr(system, field, val)
    await db.commit()
    return await _load_system(system_id, db)


@router.delete("/{system_id}", status_code=204)
async def delete_system(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")
    system = await _load_system(system_id, db)
    await db.delete(system)
    await db.commit()


# ── System Items ───────────────────────────────────────────────────────────────

@router.post("/{system_id}/items", response_model=SystemItemOut, status_code=201)
async def add_item(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    body: SystemItemCreate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")

    item = SystemItem(system_id=system_id, **body.model_dump())

    # If equipment_id given, pull defaults from DB
    if body.equipment_id and not body.msrp:
        result = await db.execute(select(Equipment).where(Equipment.id == body.equipment_id))
        equip = result.scalar_one_or_none()
        if equip:
            item.item_id = item.item_id or equip.item_id
            item.description = item.description or equip.description
            item.notes = item.notes or equip.notes
            item.msrp = equip.msrp
            item.multiplier = equip.multiplier

    db.add(item)
    await db.commit()
    await db.refresh(item)
    d = _enrich_item(item)
    return SystemItemOut(**d)


@router.patch("/{system_id}/items/{item_id}", response_model=SystemItemOut)
async def update_item(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    item_id: uuid.UUID,
    body: SystemItemUpdate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")
    result = await db.execute(select(SystemItem).where(SystemItem.id == item_id))
    item = result.scalar_one_or_none()
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    for field, val in body.model_dump(exclude_unset=True).items():
        setattr(item, field, val)
    await db.commit()
    await db.refresh(item)
    d = _enrich_item(item)
    return SystemItemOut(**d)


@router.delete("/{system_id}/items/{item_id}", status_code=204)
async def delete_item(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    item_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")
    result = await db.execute(select(SystemItem).where(SystemItem.id == item_id))
    item = result.scalar_one_or_none()
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    await db.delete(item)
    await db.commit()


@router.post("/{system_id}/items/{item_id}/ofci")
async def set_ofci(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    item_id: uuid.UUID,
    ofci_type: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")
    result = await db.execute(select(SystemItem).where(SystemItem.id == item_id))
    item = result.scalar_one_or_none()
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    item.is_ofci = True
    item.ofci_type = ofci_type
    await db.commit()
    return {"ok": True}


@router.delete("/{system_id}/items/{item_id}/ofci")
async def clear_ofci(
    project_id: uuid.UUID,
    system_id: uuid.UUID,
    item_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_project_member(project_id, current_user, db, min_role="editor")
    result = await db.execute(select(SystemItem).where(SystemItem.id == item_id))
    item = result.scalar_one_or_none()
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    item.is_ofci = False
    item.ofci_type = None
    await db.commit()
    return {"ok": True}


# ── Summary ────────────────────────────────────────────────────────────────────

