"""Initial schema

Revision ID: 0001
Revises:
Create Date: 2026-04-13
"""
from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql

revision = "0001"
down_revision = None
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "users",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("email", sa.String(255), nullable=False, unique=True),
        sa.Column("display_name", sa.String(255), nullable=False),
        sa.Column("azure_oid", sa.String(255), nullable=True, unique=True),
        sa.Column("role", sa.Enum("admin", "user", name="userrole"), nullable=False, server_default="user"),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )
    op.create_index("ix_users_email", "users", ["email"])
    op.create_index("ix_users_azure_oid", "users", ["azure_oid"])

    op.create_table(
        "equipment",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("item_id", sa.String(100), nullable=False, unique=True),
        sa.Column("mfr", sa.String(255), nullable=False),
        sa.Column("model", sa.String(255), nullable=False),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("notes", sa.Text(), nullable=True),
        sa.Column("msrp", sa.Numeric(12, 2), nullable=False, server_default="0"),
        sa.Column("multiplier", sa.Numeric(6, 4), nullable=False, server_default="1.0"),
        sa.Column("category", sa.String(100), nullable=True),
        sa.Column("created_by_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("users.id"), nullable=True),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )
    op.create_index("ix_equipment_item_id", "equipment", ["item_id"])
    op.create_index("ix_equipment_mfr", "equipment", ["mfr"])

    op.create_table(
        "projects",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("name", sa.String(255), nullable=False),
        sa.Column("created_by_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("users.id"), nullable=False),
        sa.Column("rounding_variable", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("discount_pct", sa.Numeric(8, 6), nullable=False, server_default="0"),
        sa.Column("contingency_pct", sa.Numeric(8, 6), nullable=False, server_default="0"),
        sa.Column("engineering_mult", sa.Numeric(8, 6), nullable=False, server_default="0.02"),
        sa.Column("pm_mult", sa.Numeric(8, 6), nullable=False, server_default="0.02"),
        sa.Column("preinstall_mult", sa.Numeric(8, 6), nullable=False, server_default="0.02"),
        sa.Column("installation_mult", sa.Numeric(8, 6), nullable=False, server_default="0.08"),
        sa.Column("programming_mult", sa.Numeric(8, 6), nullable=False, server_default="0.08"),
        sa.Column("tax_mult", sa.Numeric(8, 6), nullable=False, server_default="0.08275"),
        sa.Column("ga_mult", sa.Numeric(8, 6), nullable=False, server_default="0.10"),
        sa.Column("engineering_label", sa.String(100), nullable=False, server_default="ENGINEERING"),
        sa.Column("pm_label", sa.String(100), nullable=False, server_default="PROJECT MANAGEMENT"),
        sa.Column("preinstall_label", sa.String(100), nullable=False, server_default="PRE-INSTALLATION & RACK FABRICATION"),
        sa.Column("installation_label", sa.String(100), nullable=False, server_default="INSTALLATION"),
        sa.Column("programming_label", sa.String(100), nullable=False, server_default="PROGRAMMING"),
        sa.Column("tax_label", sa.String(100), nullable=False, server_default="TAX"),
        sa.Column("ga_label", sa.String(100), nullable=False, server_default="G&A"),
        sa.Column("combine_equip_nonequip", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("separate_license", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("include_support", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("ignore_hidden_tabs", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "project_members",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("project_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("projects.id", ondelete="CASCADE"), nullable=False),
        sa.Column("user_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("users.id"), nullable=False),
        sa.Column("role", sa.Enum("owner", "editor", "viewer", name="projectmemberrole"), nullable=False, server_default="editor"),
        sa.Column("joined_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "systems",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("project_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("projects.id", ondelete="CASCADE"), nullable=False),
        sa.Column("name", sa.String(255), nullable=False),
        sa.Column("system_type", sa.Enum("room_numbers", "system_count", name="systemtype"), nullable=False, server_default="room_numbers"),
        sa.Column("room_info", sa.String(500), nullable=True),
        sa.Column("display_order", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("is_visible", sa.Boolean(), nullable=False, server_default="true"),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "system_sections",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("system_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("systems.id", ondelete="CASCADE"), nullable=False),
        sa.Column("name", sa.String(255), nullable=False, server_default=""),
        sa.Column("display_order", sa.Integer(), nullable=False, server_default="0"),
    )

    op.create_table(
        "issuances",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("project_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("projects.id", ondelete="CASCADE"), nullable=False),
        sa.Column("name", sa.String(255), nullable=False),
        sa.Column("issue_date", sa.Date(), nullable=True),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "system_items",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("system_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("systems.id", ondelete="CASCADE"), nullable=False),
        sa.Column("section_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("system_sections.id"), nullable=True),
        sa.Column("equipment_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("equipment.id"), nullable=True),
        sa.Column("display_order", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("item_id", sa.String(100), nullable=True),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("notes", sa.Text(), nullable=True),
        sa.Column("qty_per_room", sa.Numeric(10, 2), nullable=False, server_default="0"),
        sa.Column("msrp", sa.Numeric(12, 2), nullable=False, server_default="0"),
        sa.Column("multiplier", sa.Numeric(6, 4), nullable=False, server_default="1.0"),
        sa.Column("is_section_header", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("is_note_row", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("note_text", sa.Text(), nullable=True),
        sa.Column("is_bold_note", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("is_ofci", sa.Boolean(), nullable=False, server_default="false"),
        sa.Column("ofci_type", sa.String(50), nullable=True),
        sa.Column("last_issuance_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("issuances.id"), nullable=True),
        sa.Column("old_qty", sa.Numeric(10, 2), nullable=True),
        sa.Column("change_status", sa.String(20), nullable=True),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
        sa.Column("updated_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "issuance_systems",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("issuance_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("issuances.id", ondelete="CASCADE"), nullable=False),
        sa.Column("system_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("systems.id", ondelete="CASCADE"), nullable=False),
        sa.Column("issued_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "revision_entries",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("project_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("projects.id", ondelete="CASCADE"), nullable=False),
        sa.Column("issuance_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("issuances.id", ondelete="CASCADE"), nullable=False),
        sa.Column("system_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("systems.id"), nullable=True),
        sa.Column("system_item_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("system_items.id"), nullable=True),
        sa.Column("system_name", sa.String(255), nullable=True),
        sa.Column("mfr", sa.String(255), nullable=True),
        sa.Column("model", sa.String(255), nullable=True),
        sa.Column("item_id", sa.String(100), nullable=True),
        sa.Column("old_qty", sa.Numeric(10, 2), nullable=True),
        sa.Column("new_qty", sa.Numeric(10, 2), nullable=True),
        sa.Column("status", sa.String(20), nullable=True),
        sa.Column("summary", sa.Text(), nullable=True),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )

    op.create_table(
        "project_archives",
        sa.Column("id", postgresql.UUID(as_uuid=True), primary_key=True),
        sa.Column("project_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("projects.id", ondelete="CASCADE"), nullable=False),
        sa.Column("name", sa.String(255), nullable=False),
        sa.Column("archive_type", sa.String(50), nullable=False),
        sa.Column("created_by_id", postgresql.UUID(as_uuid=True), sa.ForeignKey("users.id"), nullable=True),
        sa.Column("created_at", sa.DateTime(), nullable=False, server_default=sa.func.now()),
    )


def downgrade() -> None:
    op.drop_table("project_archives")
    op.drop_table("revision_entries")
    op.drop_table("issuance_systems")
    op.drop_table("system_items")
    op.drop_table("issuances")
    op.drop_table("system_sections")
    op.drop_table("systems")
    op.drop_table("project_members")
    op.drop_table("projects")
    op.drop_table("equipment")
    op.drop_table("users")
    op.execute("DROP TYPE IF EXISTS userrole")
    op.execute("DROP TYPE IF EXISTS projectmemberrole")
    op.execute("DROP TYPE IF EXISTS systemtype")
