from __future__ import annotations

from datetime import datetime
from sqlalchemy import Column, DateTime, Float, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import declarative_base
from .db import engine

Base = declarative_base()


class LearningMapping(Base):
    __tablename__ = "learning_mappings"

    id = Column(Integer, primary_key=True, autoincrement=True)
    marketplace = Column(String(10), nullable=False)
    input_signature = Column(String(1500), nullable=False)

    product_type = Column(String(500), nullable=True)
    source_node_id = Column(String(100), nullable=True)
    source_category_name = Column(String(800), nullable=True)
    category_name_eng = Column(String(800), nullable=True)
    pim_category_name = Column(String(800), nullable=True)
    ean = Column(String(150), nullable=True)

    target_value = Column(String(500), nullable=False)  # najčešće node id, ali može biti i ručni naziv ako node nije poznat
    target_category_name = Column(String(800), nullable=True)
    target_path = Column(Text, nullable=True)
    confidence = Column(Float, nullable=True)
    status = Column(String(80), nullable=False)
    note = Column(Text, nullable=True)
    source = Column(String(80), nullable=False, default="USER")

    created_by = Column(String(100), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow)
    usage_count = Column(Integer, default=0)

    __table_args__ = (
        UniqueConstraint("marketplace", "input_signature", name="uq_learning_marketplace_signature"),
    )


def init_db() -> None:
    Base.metadata.create_all(bind=engine)
