"""Publish BoM / Estimate + Equipment Count endpoints."""
from fastapi import APIRouter, Depends, HTTPException, Body
from fastapi.responses import StreamingResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from sqlalchemy.orm import selectinload
from typing import List
import uuid
import io

from app.db.database import get_db
from app.models.user import User
from app.models.project import Project, ProjectMember
from app.models.system import System, SystemItem
from app.schemas.issuance import PublishRequest, EquipmentCountRow, IssuanceCreate, IssuanceOut
from app.auth.deps import get_current_user
from app.services.export_xlsx import generate_bom, generate_estimate
from app.services.equipment_count import get_equipment_count
from app.services.issuance_service import create_issuance
from app.services.summary import get_project_summary

router = APIRouter(prefix="/projects/{project_id}", tags=["publish"])


async def _get_member(project_id: uuid.UUID, user: User, db: AsyncSession):
    result = await db.execute(
        select(ProjectMember).where(
            ProjectMember.project_id == project_id,
            ProjectMember.user_id == user.id,
        )
    )
    m = result.scalar_one_or_none()
    if not m:
        raise HTTPException(status_code=403, detail="Access denied")
    return m


@router.get("/summary")
async def summary(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_member(project_id, current_user, db)
    return await get_project_summary(project_id, db)


@router.post("/publish/bom")
async def publish_bom(
    project_id: uuid.UUID,
    body: PublishRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_member(project_id, current_user, db)

    result = await db.execute(select(Project).where(Project.id == project_id))
    project = result.scalar_one_or_none()
    if not project:
        raise HTTPException(status_code=404, detail="Project not found")

    # Create issuance
    issuance = await create_issuance(
        project_id=project_id,
        name=body.issuance_name,
        issue_date=body.issuance_date,
        system_ids=body.system_ids,
        db=db,
    )

    # Load systems with items
    sys_result = await db.execute(
        select(System)
        .where(System.id.in_(body.system_ids))
        .options(selectinload(System.items).selectinload(SystemItem.equipment))
        .order_by(System.display_order)
    )
    systems = sys_result.scalars().all()

    xlsx_bytes = await generate_bom(
        project=project,
        systems=systems,
        issuance_name=body.issuance_name,
        include_notes=body.include_notes,
    )

    filename = f"27_41_16_Appendix_A_{body.issuance_name}.xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/publish/estimate")
async def publish_estimate(
    project_id: uuid.UUID,
    body: PublishRequest,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_member(project_id, current_user, db)

    result = await db.execute(select(Project).where(Project.id == project_id))
    project = result.scalar_one_or_none()
    if not project:
        raise HTTPException(status_code=404, detail="Project not found")

    issuance = await create_issuance(
        project_id=project_id,
        name=body.issuance_name,
        issue_date=body.issuance_date,
        system_ids=body.system_ids,
        db=db,
    )

    sys_result = await db.execute(
        select(System)
        .where(System.id.in_(body.system_ids))
        .options(selectinload(System.items).selectinload(SystemItem.equipment))
        .order_by(System.display_order)
    )
    systems = sys_result.scalars().all()

    xlsx_bytes = await generate_estimate(
        project=project,
        systems=systems,
        issuance_name=body.issuance_name,
        include_notes=body.include_notes,
        include_cost=body.include_cost,
        include_labor_breakout=body.include_labor_breakout,
    )

    filename = f"AV_EoPC_{body.issuance_name}.xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/equipment-count", response_model=List[EquipmentCountRow])
async def equipment_count(
    project_id: uuid.UUID,
    system_ids: List[uuid.UUID] = Body(...),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_member(project_id, current_user, db)
    return await get_equipment_count(project_id, system_ids, db)


@router.post("/issuances", response_model=IssuanceOut, status_code=201)
async def create_issuance_endpoint(
    project_id: uuid.UUID,
    body: IssuanceCreate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    await _get_member(project_id, current_user, db)
    issuance = await create_issuance(
        project_id=project_id,
        name=body.name,
        issue_date=body.issue_date,
        system_ids=body.system_ids,
        db=db,
    )
    # Reload with relationships to avoid lazy-load in async context
    from app.models.issuance import Issuance, IssuanceSystem
    iso_result = await db.execute(
        select(Issuance)
        .where(Issuance.id == issuance.id)
        .options(selectinload(Issuance.issuance_systems))
    )
    issuance = iso_result.scalar_one()
    sys_result = await db.execute(
        select(System).where(System.project_id == project_id)
    )
    system_names = {str(s.id): s.name for s in sys_result.scalars().all()}
    return IssuanceOut(
        id=issuance.id,
        project_id=issuance.project_id,
        name=issuance.name,
        issue_date=issuance.issue_date,
        created_at=issuance.created_at,
        system_names=[system_names.get(str(s.system_id), "") for s in issuance.issuance_systems],
    )


@router.get("/issuances")
async def list_issuances(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    from app.models.issuance import Issuance, IssuanceSystem
    await _get_member(project_id, current_user, db)
    result = await db.execute(
        select(Issuance)
        .where(Issuance.project_id == project_id)
        .options(selectinload(Issuance.issuance_systems))
        .order_by(Issuance.created_at.desc())
    )
    issuances = result.scalars().all()
    sys_result = await db.execute(select(System).where(System.project_id == project_id))
    system_names = {str(s.id): s.name for s in sys_result.scalars().all()}
    return [
        IssuanceOut(
            id=i.id,
            project_id=i.project_id,
            name=i.name,
            issue_date=i.issue_date,
            created_at=i.created_at,
            system_names=[system_names.get(str(s.system_id), "") for s in i.issuance_systems],
        )
        for i in issuances
    ]


@router.get("/revisions")
async def list_revisions(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    from app.models.issuance import RevisionEntry
    await _get_member(project_id, current_user, db)
    result = await db.execute(
        select(RevisionEntry)
        .where(RevisionEntry.project_id == project_id)
        .order_by(RevisionEntry.created_at.desc())
    )
    return result.scalars().all()
