"""Project CRUD + settings + member management."""
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from sqlalchemy.orm import selectinload
from typing import List
import uuid

from app.db.database import get_db
from app.models.user import User
from app.models.project import Project, ProjectMember
from app.schemas.project import ProjectCreate, ProjectUpdate, ProjectOut, ProjectMemberOut
from app.auth.deps import get_current_user

router = APIRouter(prefix="/projects", tags=["projects"])


async def _get_project_or_404(project_id: uuid.UUID, db: AsyncSession) -> Project:
    result = await db.execute(
        select(Project)
        .where(Project.id == project_id)
        .options(selectinload(Project.members).selectinload(ProjectMember.user))
    )
    project = result.scalar_one_or_none()
    if not project:
        raise HTTPException(status_code=404, detail="Project not found")
    return project


async def _assert_access(project: Project, user: User, min_role: str = "viewer"):
    roles_order = {"viewer": 0, "editor": 1, "owner": 2}
    for m in project.members:
        if m.user_id == user.id:
            if roles_order.get(m.role.value, 0) >= roles_order.get(min_role, 0):
                return
    raise HTTPException(status_code=403, detail="Access denied")


@router.get("/", response_model=List[ProjectOut])
async def list_projects(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Project)
        .join(ProjectMember)
        .where(ProjectMember.user_id == current_user.id)
        .order_by(Project.updated_at.desc())
    )
    return result.scalars().all()


@router.post("/", response_model=ProjectOut, status_code=201)
async def create_project(
    body: ProjectCreate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    s = body.settings or {}
    settings_dict = s.model_dump() if hasattr(s, "model_dump") else {}
    project = Project(name=body.name, created_by_id=current_user.id, **settings_dict)
    db.add(project)
    await db.flush()
    db.add(ProjectMember(project_id=project.id, user_id=current_user.id, role="owner"))
    await db.commit()
    await db.refresh(project)
    return project


@router.get("/{project_id}", response_model=ProjectOut)
async def get_project(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    project = await _get_project_or_404(project_id, db)
    await _assert_access(project, current_user)
    return project


@router.patch("/{project_id}", response_model=ProjectOut)
async def update_project(
    project_id: uuid.UUID,
    body: ProjectUpdate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    project = await _get_project_or_404(project_id, db)
    await _assert_access(project, current_user, min_role="editor")
    if body.name is not None:
        project.name = body.name
    if body.settings:
        for field, val in body.settings.model_dump(exclude_unset=True).items():
            setattr(project, field, val)
    await db.commit()
    await db.refresh(project)
    return project


@router.delete("/{project_id}", status_code=204)
async def delete_project(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    project = await _get_project_or_404(project_id, db)
    await _assert_access(project, current_user, min_role="owner")
    await db.delete(project)
    await db.commit()


@router.get("/{project_id}/members", response_model=List[ProjectMemberOut])
async def list_members(
    project_id: uuid.UUID,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    project = await _get_project_or_404(project_id, db)
    await _assert_access(project, current_user)
    return [
        ProjectMemberOut(
            user_id=m.user_id,
            email=m.user.email,
            display_name=m.user.display_name,
            role=m.role.value,
        )
        for m in project.members
    ]


@router.post("/{project_id}/members/{user_email}")
async def add_member(
    project_id: uuid.UUID,
    user_email: str,
    role: str = "editor",
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    project = await _get_project_or_404(project_id, db)
    await _assert_access(project, current_user, min_role="owner")
    result = await db.execute(select(User).where(User.email == user_email))
    target = result.scalar_one_or_none()
    if not target:
        raise HTTPException(status_code=404, detail="User not found")
    db.add(ProjectMember(project_id=project_id, user_id=target.id, role=role))
    await db.commit()
    return {"ok": True}


@router.post("/import-xlsx")
async def import_project_xlsx(
    file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Import an existing .xlsx and create a new project from it."""
    from app.services.import_xlsx import import_xlsx
    content = await file.read()
    project = await import_xlsx(content, current_user, db)
    return {"project_id": str(project.id), "name": project.name}
