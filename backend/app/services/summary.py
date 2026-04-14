"""Summary calculation service — mirrors the VBA pullNumEST / pullNum logic."""
from decimal import Decimal, ROUND_HALF_UP
from typing import List
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from sqlalchemy.orm import selectinload

from app.models.project import Project
from app.models.system import System, SystemItem
from app.schemas.system import SystemSummaryOut


def _round_value(value: Decimal, rounding_variable: int) -> Decimal:
    """Apply Excel-style ROUND() based on rounding_variable.
    0 = no rounding, -1 = tens, -2 = hundreds, -3 = thousands, etc.
    """
    if rounding_variable == 0:
        return value
    factor = Decimal(10) ** (-rounding_variable)
    rounded = (value / factor).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * factor
    return rounded


def compute_system_summary(system: System, project: Project) -> SystemSummaryOut:
    """Compute cost summary for one system."""
    # Sum extended cost of all regular equipment items
    equipment_subtotal = Decimal("0")
    for item in system.items:
        if item.is_section_header or item.is_note_row:
            continue
        # unit_cost = MSRP × multiplier (if OFCI, cost = 0 for estimate purposes)
        if item.is_ofci:
            continue
        unit_cost = item.msrp * item.multiplier
        extended = unit_cost * item.qty_per_room
        equipment_subtotal += extended

    # Apply rounding
    equipment_subtotal = _round_value(equipment_subtotal, project.rounding_variable)

    # Discount
    discount_pct = project.discount_pct
    discount_amount = _round_value(equipment_subtotal * discount_pct, project.rounding_variable)
    discounted_equipment = _round_value(equipment_subtotal - discount_amount, project.rounding_variable)

    # Non-equipment multipliers applied to discounted equipment
    total_ne_mult = (
        project.engineering_mult
        + project.pm_mult
        + project.preinstall_mult
        + project.installation_mult
        + project.programming_mult
        + project.tax_mult
        + project.ga_mult
    )
    non_equipment_subtotal = _round_value(discounted_equipment * total_ne_mult, project.rounding_variable)

    # Contingency on (discounted equipment + non-equipment)
    contingency_pct = project.contingency_pct
    base_for_contingency = discounted_equipment + non_equipment_subtotal
    contingency_amount = _round_value(base_for_contingency * contingency_pct, project.rounding_variable)

    system_subtotal = discounted_equipment + non_equipment_subtotal + contingency_amount
    system_subtotal = _round_value(system_subtotal, project.rounding_variable)

    room_count = system.room_count
    system_extended = _round_value(system_subtotal * room_count, project.rounding_variable)

    return SystemSummaryOut(
        system_id=system.id,
        system_name=system.name,
        room_info=system.room_info,
        room_count=room_count,
        equipment_subtotal=equipment_subtotal,
        discount_amount=discount_amount,
        discounted_equipment=discounted_equipment,
        non_equipment_subtotal=non_equipment_subtotal,
        contingency_pct=contingency_pct,
        contingency_amount=contingency_amount,
        system_subtotal=system_subtotal,
        system_extended=system_extended,
    )


async def get_project_summary(project_id, db: AsyncSession) -> dict:
    """Return full project summary: per-system rows + project totals."""
    result = await db.execute(
        select(Project).where(Project.id == project_id)
    )
    project = result.scalar_one_or_none()
    if not project:
        return {}

    sys_result = await db.execute(
        select(System)
        .where(System.project_id == project_id, System.is_visible == True)
        .options(selectinload(System.items))
        .order_by(System.display_order)
    )
    systems = sys_result.scalars().all()

    rows = [compute_system_summary(s, project) for s in systems]

    # Project totals
    total_equipment = sum(r.equipment_subtotal for r in rows)
    total_discount = sum(r.discount_amount for r in rows)
    total_discounted = sum(r.discounted_equipment for r in rows)
    total_non_equip = sum(r.non_equipment_subtotal for r in rows)
    total_contingency = sum(r.contingency_amount for r in rows)
    total_installed = sum(r.system_extended for r in rows)

    # Non-equipment line breakout
    ne_lines = [
        {"label": project.engineering_label, "mult": float(project.engineering_mult)},
        {"label": project.pm_label, "mult": float(project.pm_mult)},
        {"label": project.preinstall_label, "mult": float(project.preinstall_mult)},
        {"label": project.installation_label, "mult": float(project.installation_mult)},
        {"label": project.programming_label, "mult": float(project.programming_mult)},
        {"label": project.tax_label, "mult": float(project.tax_mult)},
        {"label": project.ga_label, "mult": float(project.ga_mult)},
    ]

    return {
        "project_name": project.name,
        "systems": [r.model_dump() for r in rows],
        "totals": {
            "total_equipment_subtotal": float(total_equipment),
            "total_discount": float(total_discount),
            "total_discounted_equipment": float(total_discounted),
            "total_non_equipment": float(total_non_equip),
            "total_contingency": float(total_contingency),
            "total_installed_cost": float(total_installed),
        },
        "non_equipment_lines": ne_lines,
        "settings": {
            "discount_pct": float(project.discount_pct),
            "contingency_pct": float(project.contingency_pct),
            "rounding_variable": project.rounding_variable,
        },
    }
