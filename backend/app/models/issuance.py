import uuid
from datetime import datetime, date
from decimal import Decimal
from sqlalchemy import String, Text, Numeric, DateTime, Date, ForeignKey
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db.database import Base


class Issuance(Base):
    __tablename__ = "issuances"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    project_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"))
    name: Mapped[str] = mapped_column(String(255))
    issue_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    project = relationship("Project", back_populates="issuances")
    issuance_systems = relationship("IssuanceSystem", back_populates="issuance", cascade="all, delete-orphan")
    revision_entries = relationship("RevisionEntry", back_populates="issuance", cascade="all, delete-orphan")


class IssuanceSystem(Base):
    __tablename__ = "issuance_systems"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    issuance_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("issuances.id", ondelete="CASCADE"))
    system_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("systems.id", ondelete="CASCADE"))
    issued_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    issuance = relationship("Issuance", back_populates="issuance_systems")
    system = relationship("System", back_populates="issuance_systems")


class RevisionEntry(Base):
    __tablename__ = "revision_entries"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    project_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"))
    issuance_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("issuances.id", ondelete="CASCADE"))
    system_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("systems.id"), nullable=True)
    system_item_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("system_items.id"), nullable=True)
    system_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    mfr: Mapped[str | None] = mapped_column(String(255), nullable=True)
    model: Mapped[str | None] = mapped_column(String(255), nullable=True)
    item_id: Mapped[str | None] = mapped_column(String(100), nullable=True)
    old_qty: Mapped[Decimal | None] = mapped_column(Numeric(10, 2), nullable=True)
    new_qty: Mapped[Decimal | None] = mapped_column(Numeric(10, 2), nullable=True)
    status: Mapped[str | None] = mapped_column(String(20), nullable=True)
    summary: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    issuance = relationship("Issuance", back_populates="revision_entries")
    system = relationship("System")
    system_item = relationship("SystemItem")
