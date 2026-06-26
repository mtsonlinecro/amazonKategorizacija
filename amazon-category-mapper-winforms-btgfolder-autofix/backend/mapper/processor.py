from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Any
import pandas as pd

from .io_csv import MARKETPLACES, read_input_file, write_csv
from .mapping_table import build_mapping_index, load_mappings, lookup_direct_mapping
from .learning_store import find_exact_mapping, find_similar_mapping, save_mapping
from .ai_matcher import guess_amazon_category_with_openai

DIRECT_MARKETPLACES = {"FR", "IT", "ES"}
AI_FALLBACK_MARKETPLACES = {"NL", "PL", "IE", "SE"}


def ensure_output_columns(df: pd.DataFrame, marketplaces: List[str]) -> pd.DataFrame:
    for mp in marketplaces:
        if mp not in df.columns:
            df[mp] = ""
        for suffix in ["status", "source", "note", "confidence", "category_name", "source_url"]:
            col = f"{mp}_{suffix}"
            if col not in df.columns:
                df[col] = ""
    return df


def _clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def mapping_to_row(mapping, mp: str) -> Dict[str, Any]:
    return {
        mp: mapping.target_value,
        f"{mp}_status": "LEARNED_MATCH",
        f"{mp}_source": "LEARNING_DB",
        f"{mp}_note": mapping.note or "Pronađeno u learning bazi.",
        f"{mp}_confidence": mapping.confidence if mapping.confidence is not None else 1.0,
        f"{mp}_category_name": mapping.target_category_name or "",
        f"{mp}_source_url": "",
    }


def process_file(
    input_file_path: str | Path,
    mapping_file_path: str | Path,
    output_file_path: str | Path,
    selected_marketplaces: List[str] | None = None,
    use_ai: bool = False,
    overwrite_existing: bool = True,
    max_ai_rows: int = 50,
) -> pd.DataFrame:
    selected_marketplaces = [m.upper().strip() for m in (selected_marketplaces or ["FR", "IT", "ES"])]
    selected_marketplaces = [m for m in selected_marketplaces if m in MARKETPLACES]

    input_df = read_input_file(input_file_path)
    input_df = ensure_output_columns(input_df, selected_marketplaces)

    mapping_df = load_mappings(mapping_file_path)
    mapping_index = build_mapping_index(mapping_df)

    ai_calls = 0
    for idx, row in input_df.iterrows():
        row_dict = row.to_dict()
        source_node_id = _clean(row_dict.get("NODE ID"))

        for mp in selected_marketplaces:
            current_value = _clean(input_df.at[idx, mp]) if mp in input_df.columns else ""
            if current_value and not overwrite_existing:
                input_df.at[idx, f"{mp}_status"] = "ALREADY_FILLED"
                input_df.at[idx, f"{mp}_source"] = "INPUT"
                continue

            # 1. learning baza prije svega
            learned = find_exact_mapping(row_dict, mp)
            if learned:
                for key, value in mapping_to_row(learned, mp).items():
                    input_df.at[idx, key] = value
                continue

            similar, score = find_similar_mapping(row_dict, mp)
            if similar:
                data = mapping_to_row(similar, mp)
                data[f"{mp}_status"] = "LEARNED_SIMILAR"
                data[f"{mp}_confidence"] = round(float(score) / 100, 3)
                data[f"{mp}_note"] = f"Slično naučeno mapiranje, score={score}. Provjeriti ako je važno."
                for key, value in data.items():
                    input_df.at[idx, key] = value
                continue

            # 2. direktni Amazon EU mapping za FR/IT/ES
            if mp in DIRECT_MARKETPLACES:
                direct = lookup_direct_mapping(mapping_index, source_node_id, mp)
                input_df.at[idx, mp] = direct.target_node_id
                input_df.at[idx, f"{mp}_status"] = direct.status
                input_df.at[idx, f"{mp}_source"] = "EUROPEAN_BROWSE_NODE_MAPPING" if direct.status == "DIRECT_MAPPING" else "MAPPING_TABLE"
                input_df.at[idx, f"{mp}_note"] = direct.note
                input_df.at[idx, f"{mp}_confidence"] = 1.0 if direct.status == "DIRECT_MAPPING" else 0.0
                input_df.at[idx, f"{mp}_category_name"] = direct.node_path
                input_df.at[idx, f"{mp}_source_url"] = ""

                # popuni i originalne stupce Node Id in FR/IT/ES ako postoje
                node_col = f"Node Id in {mp}"
                if node_col in input_df.columns and (overwrite_existing or not _clean(input_df.at[idx, node_col])):
                    input_df.at[idx, node_col] = direct.target_node_id
                continue

            # 3. AI/web fallback samo za ostale države i samo ako korisnik uključi
            if mp in AI_FALLBACK_MARKETPLACES and use_ai and ai_calls < max_ai_rows:
                ai_calls += 1
                guess = guess_amazon_category_with_openai(row_dict, mp)
                input_df.at[idx, mp] = guess.get("value", "")
                input_df.at[idx, f"{mp}_status"] = guess.get("status", "NEED_REVIEW")
                input_df.at[idx, f"{mp}_source"] = "OPENAI_WEB_SEARCH"
                input_df.at[idx, f"{mp}_note"] = guess.get("reason", "")
                input_df.at[idx, f"{mp}_confidence"] = guess.get("confidence", 0.0)
                input_df.at[idx, f"{mp}_category_name"] = guess.get("category_name", "")
                input_df.at[idx, f"{mp}_source_url"] = guess.get("source_url", "")
            else:
                input_df.at[idx, mp] = ""
                input_df.at[idx, f"{mp}_status"] = "NEED_REVIEW" if mp in AI_FALLBACK_MARKETPLACES else "UNSUPPORTED"
                input_df.at[idx, f"{mp}_source"] = "NONE"
                input_df.at[idx, f"{mp}_note"] = "Nema direktnog mapping stupca. Uključi AI fallback ili ručno ispravi." if mp in AI_FALLBACK_MARKETPLACES else "Nepodržana država."
                input_df.at[idx, f"{mp}_confidence"] = 0.0
                input_df.at[idx, f"{mp}_category_name"] = ""
                input_df.at[idx, f"{mp}_source_url"] = ""

    write_csv(input_df, output_file_path)
    return input_df


def save_corrections_from_rows(rows: List[Dict[str, Any]], selected_marketplaces: List[str] | None = None) -> int:
    selected_marketplaces = [m.upper().strip() for m in (selected_marketplaces or MARKETPLACES)]
    count = 0
    for row in rows:
        for mp in selected_marketplaces:
            value = _clean(row.get(mp))
            status = _clean(row.get(f"{mp}_status"))
            category_name = _clean(row.get(f"{mp}_category_name"))
            note = _clean(row.get(f"{mp}_note"))
            confidence_text = _clean(row.get(f"{mp}_confidence"))
            try:
                confidence = float(confidence_text.replace(",", ".")) if confidence_text else 1.0
            except Exception:
                confidence = 1.0

            # spremamo samo kad je korisnik potvrdio/ispravio; ne učimo iz NEED_REVIEW i NO_MATCH
            if status not in {"CONFIRMED", "CORRECTED"}:
                continue
            if not value and not category_name:
                continue
            save_mapping(
                row=row,
                marketplace=mp,
                target_value=value or category_name,
                target_category_name=category_name,
                target_path=category_name,
                confidence=confidence,
                status=status,
                note=note,
                source="USER_CORRECTION",
                created_by="winforms_user",
            )
            count += 1
    return count
