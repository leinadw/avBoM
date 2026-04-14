import uuid
from datetime import datetime
from decimal import Decimal
from sqlalchemy import String, Text, Numeric, DateTime, Boolean, ForeignKey, Enum as SAEnum, Integer
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db.database import Base
import enum


class SystemType(str, enum.Enum):
    room_numbers = "room_numbers"
    system_count = "system_count"


class OfciType(str, enum.Enum):
    OFE = "OFE"
    OFCI = "OFCI"
    OFOI = "OFOI"
    custom = "custom"


class ChangeStatus(str, enum.Enum):
    increased = "increased"
    decreased = "decreased"
    unchanged = "unchanged"
    new = "new"
    removed = "removed"


class System(Base):
    __tablename__ = "systems"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    project_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"))
    name: Mapped[str] = mapped_column(String(255))
    system_type: Mapped[SystemType] = mapped_column(SAEnum(SystemType), default=SystemType.room_numbers)
    room_info: Mapped[str | None] = mapped_column(String(500), nullable=True)
    display_order: Mapped[int] = mapped_column(Integer, default=0)
    is_visible: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    project = relationship("Project", back_populates="systems")
    sections = relationship("SystemSection", back_populates="system", cascade="all, delete-orphan", order_by="SystemSection.display_order")
    items = relationship("SystemItem", back_populates="system", cascade="all, delete-orphan", order_by="SystemItem.display_order")
    issuance_systems = relationship("IssuanceSystem", back_populates="system")

    @property
    def room_count(self) -> int:
        """Calculate room count from room_info."""
        if not self.room_info:
            return 1
        if self.system_type == SystemType.system_count:
            try:
                return int(self.room_info)
            except (ValueError, TypeError):
                return 1
        # room_numbers: count commas + 1
        return self.room_info.count(",") + 1


class SystemSection(Base):
    __tablename__ = "system_sections"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    system_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("systems.id", ondelete="CASCADE"))
    name: Mapped[str] = mapped_column(String(255), default="")
    display_order: Mapped[int] = mapped_column(Integer, default=0)

    system = relationship("System", back_populates="sections")
    items = relationship("SystemItem", back_populates="section")


class SystemItem(Base):
    __tablename__ = "system_items"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    system_id: Mapped[uuid.UUID] = mapped_column(ForeignKey("systems.id", ondelete="CASCADE"))
    section_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("system_sections.id"), nullable=True)
    equipment_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("equipment.id"), nullable=True)
    display_order: Mapped[int] = mapped_column(Integer, default=0)

    # Equipment fields (override DB values if set directly)
    item_id: Mapped[str | None] = mapped_column(String(100), nullable=True)
    description: Mapped[str | None] = mapped_column(Text, nullable=True)
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    qty_per_room: Mapped[Decimal] = mapped_column(Numeric(10, 2), default=0)
    msrp: Mapped[Decimal] = mapped_column(Numeric(12, 2), default=0)
    multiplier: Mapped[Decimal] = mapped_column(Numeric(6, 4), default=1.0)

    # Special row types
    is_section_header: Mapped[bool] = mapped_column(Boolean, default=False)
    is_note_row: Mapped[bool] = mapped_column(Boolean, default=False)
    note_text: Mapped[str | None] = mapped_column(Text, nullable=True)
    is_bold_note: Mapped[bool] = mapped_column(Boolean, default=False)

    # OFCI
    is_ofci: Mapped[bool] = mapped_column(Boolean, default=False)
    ofci_type: Mapped[str | None] = mapped_column(String(50), nullable=True)

    # Change tracking
    last_issuance_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("issuances.id"), nullable=True)
    old_qty: Mapped[Decimal | None] = mapped_column(Numeric(10, 2), nullable=True)
    change_status: Mapped[str | None] = mapped_column(String(20), nullable=True)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    system = relationship("System", back_populates="items")
    section = relationship("SystemSection", back_populates="items")
    equipment = relationship("Equipment", back_populates="system_items")
    last_issuance = relationship("Issuance", foreign_keys=[last_issuance_id])
