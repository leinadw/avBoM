import uuid
from datetime import datetime
from decimal import Decimal
from sqlalchemy import String, Text, Numeric, DateTime, ForeignKey
from sqlalchemy.orm import Mapped, mapped_column, relationship
from app.db.database import Base


class Equipment(Base):
    __tablename__ = "equipment"

    id: Mapped[uuid.UUID] = mapped_column(primary_key=True, default=uuid.uuid4)
    item_id: Mapped[str] = mapped_column(String(100), unique=True, index=True)
    mfr: Mapped[str] = mapped_column(String(255), index=True)
    model: Mapped[str] = mapped_column(String(255), index=True)
    description: Mapped[str | None] = mapped_column(Text, nullable=True)
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    msrp: Mapped[Decimal] = mapped_column(Numeric(12, 2), default=0)
    multiplier: Mapped[Decimal] = mapped_column(Numeric(6, 4), default=1.0)
    category: Mapped[str | None] = mapped_column(String(100), nullable=True, index=True)
    created_by_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey("users.id"), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    created_by_user = relationship("User", back_populates="equipment_created")
    system_items = relationship("SystemItem", back_populates="equipment")
