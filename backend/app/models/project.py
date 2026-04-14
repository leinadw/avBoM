import uuid
from datetime import datetime
from decimal import Decimal
from sqlalchemy import String, Numeric, DateTime, Boolean, ForeignKey, Enum as SAEnum, Integer
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db.database import Base
import enum


class ProjectMemberRole(str, enum.Enum):
    owner = "owner"
    editor = "editor"
    viewer = "viewer"


class Project(Base):
    __tablename__ = "projects"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    name: Mapped[str] = mapped_column(String(255))
    created_by_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("users.id"))

    # Cost settings
    rounding_variable: Mapped[int] = mapped_column(Integer, default=0)
    discount_pct: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=0)
    contingency_pct: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=0)

    # Non-equipment multipliers
    engineering_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.02"))
    pm_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.02"))
    preinstall_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.02"))
    installation_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.08"))
    programming_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.08"))
    tax_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.08275"))
    ga_mult: Mapped[Decimal] = mapped_column(Numeric(8, 6), default=Decimal("0.10"))

    # Non-equipment label overrides (from PROJECT_SETTINGS)
    engineering_label: Mapped[str] = mapped_column(String(100), default="ENGINEERING")
    pm_label: Mapped[str] = mapped_column(String(100), default="PROJECT MANAGEMENT")
    preinstall_label: Mapped[str] = mapped_column(String(100), default="PRE-INSTALLATION & RACK FABRICATION")
    installation_label: Mapped[str] = mapped_column(String(100), default="INSTALLATION")
    programming_label: Mapped[str] = mapped_column(String(100), default="PROGRAMMING")
    tax_label: Mapped[str] = mapped_column(String(100), default="TAX")
    ga_label: Mapped[str] = mapped_column(String(100), default="G&A")

    # Optional features
    combine_equip_nonequip: Mapped[bool] = mapped_column(Boolean, default=False)
    separate_license: Mapped[bool] = mapped_column(Boolean, default=False)
    include_support: Mapped[bool] = mapped_column(Boolean, default=False)
    ignore_hidden_tabs: Mapped[bool] = mapped_column(Boolean, default=False)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    members = relationship("ProjectMember", back_populates="project", cascade="all, delete-orphan")
    systems = relationship("System", back_populates="project", cascade="all, delete-orphan", order_by="System.display_order")
    issuances = relationship("Issuance", back_populates="project", cascade="all, delete-orphan")
    archives = relationship("ProjectArchive", back_populates="project", cascade="all, delete-orphan")


class ProjectMember(Base):
    __tablename__ = "project_members"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    project_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"))
    user_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("users.id"))
    role: Mapped[ProjectMemberRole] = mapped_column(SAEnum(ProjectMemberRole), default=ProjectMemberRole.editor)
    joined_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project = relationship("Project", back_populates="members")
    user = relationship("User", back_populates="project_memberships")
