from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Any
import pandas as pd

SUPPORTED_DIRECT_MARKETS = {
    "UK": "Node Id in UK",
    "FR": "Node Id in FR",
    "IT": "Node Id in IT",
    "ES": "Node Id in ES",
}

REQUIRED_COLUMNS = ["Node root", "Node ID", "Node Path", "Node Id in UK", "Node Id in FR", "Node Id in IT", "Node Id in ES"]


@dataclass
class DirectMappingResult:
    marketplace: str
    source_node_id: str
    target_node_id: str
    node_root: str
    node_path: str
    status: str
    note: str


def _clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    # Amazonovi node ID-evi znaju doći kao 123.0 ako Excel pogodi broj.
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def load_mappings(mapping_file_path: str | Path) -> pd.DataFrame:
    path = Path(mapping_file_path)
    if not path.exists():
        raise FileNotFoundError(f"Ne postoji mapping datoteka: {path}")

    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        try:
            df = pd.read_excel(path, sheet_name="MAPPINGS", dtype=str)
        except ValueError as exc:
            raise ValueError("Amazon Excel mora imati sheet/tab koji se zove 'MAPPINGS'.") from exc
    else:
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("U MAPPINGS tabu fale stupci: " + ", ".join(missing))

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].map(_clean)

    df = df[df["Node ID"].astype(str).str.strip() != ""].copy()
    return df


def build_mapping_index(mapping_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    index: Dict[str, Dict[str, str]] = {}
    for _, row in mapping_df.iterrows():
        node_id = _clean(row.get("Node ID"))
        if not node_id:
            continue
        index[node_id] = {col: _clean(row.get(col)) for col in REQUIRED_COLUMNS}
    return index


def lookup_direct_mapping(mapping_index: Dict[str, Dict[str, str]], source_node_id: str, marketplace: str) -> DirectMappingResult:
    marketplace = marketplace.upper().strip()
    node_id = _clean(source_node_id)

    if marketplace not in SUPPORTED_DIRECT_MARKETS:
        return DirectMappingResult(marketplace, node_id, "", "", "", "UNSUPPORTED_DIRECT_MAPPING", f"Nema direktnog stupca za {marketplace} u European Browse Node Mapping tablici.")

    if not node_id:
        return DirectMappingResult(marketplace, node_id, "", "", "", "MISSING_SOURCE_NODE", "U ulaznom CSV-u fali NODE ID.")

    record = mapping_index.get(node_id)
    if not record:
        return DirectMappingResult(marketplace, node_id, "", "", "", "NO_DIRECT_MAPPING", "NODE ID nije pronađen u MAPPINGS tabu.")

    target_col = SUPPORTED_DIRECT_MARKETS[marketplace]
    target_node_id = _clean(record.get(target_col))
    if not target_node_id:
        return DirectMappingResult(marketplace, node_id, "", record.get("Node root", ""), record.get("Node Path", ""), "MISSING_TARGET_NODE", f"Pronađen source node, ali nema vrijednosti u stupcu '{target_col}'.")

    return DirectMappingResult(
        marketplace=marketplace,
        source_node_id=node_id,
        target_node_id=target_node_id,
        node_root=record.get("Node root", ""),
        node_path=record.get("Node Path", ""),
        status="DIRECT_MAPPING",
        note="Popunjeno iz službene European Browse Node Mapping tablice.",
    )
