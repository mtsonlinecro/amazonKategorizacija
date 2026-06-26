from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, Optional
from rapidfuzz import fuzz
from sqlalchemy import select
from .db import SessionLocal
from .models import LearningMapping

SOURCE_COLUMNS = [
    "product type",
    "NODE ID",
    "Amazon category name",
    "category name eng",
    "PIM category name",
    "EAN",
]

SAVE_STATUSES = {"CONFIRMED", "CORRECTED", "LEARNED_MATCH", "LEARNED_SIMILAR"}


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    return text


def clean_lower(value: Any) -> str:
    return clean(value).lower()


def build_signature(row: Dict[str, Any]) -> str:
    parts = [clean_lower(row.get(col, "")) for col in SOURCE_COLUMNS]
    # EAN nekad zna biti prazan; signature ne smije biti samo EAN, nego kombinacija teksta i nodea.
    return " | ".join(parts)


def row_snapshot(row: Dict[str, Any]) -> Dict[str, str]:
    return {col: clean(row.get(col, "")) for col in SOURCE_COLUMNS}


def find_exact_mapping(row: Dict[str, Any], marketplace: str) -> Optional[LearningMapping]:
    signature = build_signature(row)
    with SessionLocal() as db:
        mapping = db.execute(
            select(LearningMapping).where(
                LearningMapping.marketplace == marketplace,
                LearningMapping.input_signature == signature,
            )
        ).scalar_one_or_none()
        if mapping:
            mapping.usage_count = (mapping.usage_count or 0) + 1
            mapping.updated_at = datetime.utcnow()
            db.commit()
        return mapping


def find_similar_mapping(row: Dict[str, Any], marketplace: str, min_score: int = 94):
    signature = build_signature(row)
    if not signature.strip(" |"):
        return None, 0
    with SessionLocal() as db:
        mappings = db.execute(
            select(LearningMapping).where(LearningMapping.marketplace == marketplace)
        ).scalars().all()
        best = None
        best_score = 0
        for mapping in mappings:
            score = fuzz.token_sort_ratio(signature, mapping.input_signature)
            if score > best_score:
                best = mapping
                best_score = score
        if best and best_score >= min_score:
            best.usage_count = (best.usage_count or 0) + 1
            best.updated_at = datetime.utcnow()
            db.commit()
            return best, best_score
        return None, best_score


def save_mapping(
    row: Dict[str, Any],
    marketplace: str,
    target_value: str,
    target_category_name: str = "",
    target_path: str = "",
    confidence: float = 1.0,
    status: str = "CORRECTED",
    note: str = "",
    source: str = "USER",
    created_by: str = "winforms_user",
) -> None:
    snapshot = row_snapshot(row)
    signature = build_signature(row)
    with SessionLocal() as db:
        existing = db.execute(
            select(LearningMapping).where(
                LearningMapping.marketplace == marketplace,
                LearningMapping.input_signature == signature,
            )
        ).scalar_one_or_none()
        if existing:
            existing.target_value = clean(target_value)
            existing.target_category_name = clean(target_category_name)
            existing.target_path = clean(target_path)
            existing.confidence = float(confidence or 0)
            existing.status = clean(status)
            existing.note = clean(note)
            existing.source = clean(source)
            existing.updated_at = datetime.utcnow()
            existing.created_by = clean(created_by)
        else:
            db.add(
                LearningMapping(
                    marketplace=marketplace,
                    input_signature=signature,
                    product_type=snapshot.get("product type"),
                    source_node_id=snapshot.get("NODE ID"),
                    source_category_name=snapshot.get("Amazon category name"),
                    category_name_eng=snapshot.get("category name eng"),
                    pim_category_name=snapshot.get("PIM category name"),
                    ean=snapshot.get("EAN"),
                    target_value=clean(target_value),
                    target_category_name=clean(target_category_name),
                    target_path=clean(target_path),
                    confidence=float(confidence or 0),
                    status=clean(status),
                    note=clean(note),
                    source=clean(source),
                    created_by=clean(created_by),
                )
            )
        db.commit()


def list_mappings(limit: int = 500):
    with SessionLocal() as db:
        return db.execute(
            select(LearningMapping).order_by(LearningMapping.id.desc()).limit(limit)
        ).scalars().all()
