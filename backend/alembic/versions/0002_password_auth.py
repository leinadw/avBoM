"""Switch from Azure AD to password auth

Revision ID: 0002
Revises: 0001
Create Date: 2026-04-27
"""
from alembic import op
import sqlalchemy as sa

revision = "0002"
down_revision = "0001"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column("users", sa.Column("hashed_password", sa.String(255), nullable=True))
    # Drop azure_oid index and column — no longer needed
    op.drop_index("ix_users_azure_oid", table_name="users")
    op.drop_column("users", "azure_oid")


def downgrade() -> None:
    op.add_column("users", sa.Column("azure_oid", sa.String(255), nullable=True))
    op.create_index("ix_users_azure_oid", "users", ["azure_oid"])
    op.drop_column("users", "hashed_password")
