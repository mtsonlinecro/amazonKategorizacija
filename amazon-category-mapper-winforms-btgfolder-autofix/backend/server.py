from __future__ import annotations

import csv
import json
import os
import re
import sqlite3
import tempfile
import urllib.request
from datetime import datetime
from difflib import SequenceMatcher
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any, Dict, List, Tuple

from openpyxl import load_workbook

try:
    import xlrd  # podrška za stare Amazon .xls datoteke
except Exception:
    xlrd = None

HOST = "127.0.0.1"
PORT = 8008
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "learning.db"
TEMP_DIR = Path(tempfile.gettempdir()) / "amazon_category_mapper_backend"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

# Korisnikov ulazni CSV/XLSX format.
# Važno: Node Id in FR/IT/ES = node ID, a FR/IT/ES/NL/PL/IE/SE = naziv/path kategorije.
INPUT_COLUMNS = [
    "product type", "NODE ID", "Amazon category name", "Node Id in UK", "Node Id in FR",
    "Node Id in IT", "Node Id in ES", "category name eng", "PIM category name", "EAN",
    "FR", "IT", "NL", "ES", "PL", "IE", "SE",
]
MARKETPLACES = ["FR", "IT", "NL", "ES", "PL", "IE", "SE"]
DIRECT_MARKETPLACES = {"FR", "IT", "ES"}
AI_FALLBACK_MARKETPLACES = {"NL", "PL", "IE", "SE"}
DIRECT_COLUMNS = {"UK": "Node Id in UK", "FR": "Node Id in FR", "IT": "Node Id in IT", "ES": "Node Id in ES"}
REQUIRED_MAPPING_COLUMNS = ["Node root", "Node ID", "Node Path", "Node Id in UK", "Node Id in FR", "Node Id in IT", "Node Id in ES"]
INPUT_REQUIRED_ANY = ["product type", "NODE ID", "Amazon category name", "category name eng", "PIM category name", "EAN"]


def load_env() -> Dict[str, str]:
    values: Dict[str, str] = {}
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip()
    values.update({k: v for k, v in os.environ.items() if k.startswith("OPENAI_")})
    return values


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_header(value: Any) -> str:
    return clean(value).replace("\ufeff", "").strip()


def normalize_column_key(value: Any) -> str:
    text = normalize_header(value).lower()
    return re.sub(r"[^a-z0-9]+", "", text)


COLUMN_ALIASES = {
    # mapping file
    "noderoot": "Node root",
    "nodeid": "Node ID",
    "nodepath": "Node Path",
    "nodeidinuk": "Node Id in UK",
    "nodeiduk": "Node Id in UK",
    "nodeidinfr": "Node Id in FR",
    "nodeidfr": "Node Id in FR",
    "nodeidinit": "Node Id in IT",
    "nodeidit": "Node Id in IT",
    "nodeidines": "Node Id in ES",
    "nodeides": "Node Id in ES",
    # input file
    "nodeidinput": "NODE ID",
    "node": "NODE ID",
    "producttype": "product type",
    "amazoncategoryname": "Amazon category name",
    "categorynameeng": "category name eng",
    "pimcategoryname": "PIM category name",
    "ean": "EAN",
    # category catalog
    "marketplace": "marketplace",
    "country": "marketplace",
    "store": "marketplace",
    "nodeidcategory": "node_id",
    "browseNodeId": "node_id",
    "browsenodeid": "node_id",
    "browse_node_id": "node_id",
    "node_id": "node_id",
    "categoryname": "category_name",
    "displayname": "category_name",
    "name": "category_name",
    "path": "node_path",
    "nodepathcategory": "node_path",
    "categorypath": "node_path",
    "node_path": "node_path",
}

# Dodaj marketplace stupce kao alias same na sebe
for _mp in MARKETPLACES + ["UK", "DE"]:
    COLUMN_ALIASES[normalize_column_key(_mp)] = _mp


def canonical_header(value: Any) -> str:
    raw = normalize_header(value)
    key = normalize_column_key(raw)
    return COLUMN_ALIASES.get(key, raw)


def find_header_row(values: List[Tuple[Any, ...]], required: List[str], min_score: int = 1) -> int:
    if not values:
        return 0
    required_keys = {normalize_column_key(c) for c in required}
    best_idx = 0
    best_score = -1
    for idx, row in enumerate(values[:150]):
        raw_keys = {normalize_column_key(v) for v in row if normalize_header(v)}
        canonical_keys = {normalize_column_key(COLUMN_ALIASES.get(k, k)) for k in raw_keys}
        score = len(required_keys.intersection(canonical_keys))
        if score > best_score:
            best_score = score
            best_idx = idx
        if required_keys and required_keys.issubset(canonical_keys):
            return idx
    return best_idx if best_score >= min_score else 0


def rows_from_matrix(values: List[Tuple[Any, ...]], header_idx: int = 0) -> List[Dict[str, str]]:
    if not values:
        return []
    headers = [canonical_header(v) for v in values[header_idx]]
    rows: List[Dict[str, str]] = []
    for row_values in values[header_idx + 1:]:
        row: Dict[str, str] = {}
        for i, header in enumerate(headers):
            if not header:
                continue
            value = clean(row_values[i] if i < len(row_values) else "")
            row[header] = value
        if any(v for v in row.values()):
            rows.append(row)
    return rows


def read_csv_rows(path: Path, find_header: bool = False, required: List[str] | None = None) -> List[Dict[str, str]]:
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    sample = raw[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except Exception:
        # Excel u HR često koristi ; kao separator
        dialect = csv.excel
        if sample.count(";") > sample.count(","):
            dialect.delimiter = ";"
    reader = csv.reader(raw.splitlines(), dialect=dialect)
    values = [tuple(r) for r in reader]
    header_idx = find_header_row(values, required or [], min_score=1) if find_header and required else 0
    return rows_from_matrix(values, header_idx)


def read_xlsx_rows(path: Path, sheet_name: str | None = None, find_header: bool = False, required: List[str] | None = None) -> List[Dict[str, str]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    if sheet_name:
        sheet_lookup = {s.lower(): s for s in wb.sheetnames}
        wanted = sheet_lookup.get(sheet_name.lower())
        if not wanted:
            raise ValueError(f"Excel mora imati sheet/tab '{sheet_name}'. Dostupni tabovi: {', '.join(wb.sheetnames)}")
        ws = wb[wanted]
    else:
        ws = wb[wb.sheetnames[0]]
    values = list(ws.iter_rows(values_only=True))
    header_idx = find_header_row(values, required or [], min_score=1) if find_header and required else 0
    return rows_from_matrix(values, header_idx)


def read_xls_rows(path: Path, sheet_name: str | None = None, find_header: bool = False, required: List[str] | None = None) -> List[Dict[str, str]]:
    if xlrd is None:
        raise ValueError("Za .xls treba paket xlrd. Pokreni: .venv\\Scripts\\python -m pip install xlrd==2.0.1")
    book = xlrd.open_workbook(str(path))
    if sheet_name:
        names = book.sheet_names()
        lookup = {n.lower(): n for n in names}
        wanted = lookup.get(sheet_name.lower())
        if not wanted:
            raise ValueError(f"Excel mora imati sheet/tab '{sheet_name}'. Dostupni tabovi: {', '.join(names)}")
        sh = book.sheet_by_name(wanted)
    else:
        sh = book.sheet_by_index(0)
    values: List[Tuple[Any, ...]] = []
    for r in range(sh.nrows):
        values.append(tuple(sh.cell_value(r, c) for c in range(sh.ncols)))
    header_idx = find_header_row(values, required or [], min_score=1) if find_header and required else 0
    return rows_from_matrix(values, header_idx)


def read_generic_rows(path: Path, sheet_name: str | None = None, find_header: bool = False, required: List[str] | None = None) -> List[Dict[str, str]]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return read_xlsx_rows(path, sheet_name=sheet_name, find_header=find_header, required=required)
    if suffix == ".xls":
        return read_xls_rows(path, sheet_name=sheet_name, find_header=find_header, required=required)
    return read_csv_rows(path, find_header=find_header, required=required)


def normalize_input_row(row: Dict[str, str]) -> Dict[str, str]:
    if not clean(row.get("NODE ID")) and clean(row.get("Node ID")):
        row["NODE ID"] = clean(row.get("Node ID"))
    row.pop("Node ID", None)
    return row


def read_input_rows(path: Path) -> List[Dict[str, str]]:
    # Pronađi header red, jer neki exporti imaju uvodne retke.
    rows = read_generic_rows(path, find_header=True, required=INPUT_REQUIRED_ANY)
    for row in rows:
        normalize_input_row(row)
        for col in INPUT_COLUMNS:
            row.setdefault(col, "")
    # Makni prazne redove i redove koji očito nisu proizvod
    rows = [r for r in rows if any(clean(r.get(c)) for c in ["NODE ID", "product type", "Amazon category name", "EAN", "PIM category name"])]
    if not rows:
        raise ValueError("Ulazna datoteka je pročitana, ali nije pronađen nijedan podatkovni red. Provjeri da header sadrži npr. 'NODE ID', 'product type' ili 'Amazon category name'.")
    return rows


def write_csv_rows(rows: List[Dict[str, Any]], path: Path) -> None:
    headers: List[str] = []
    for preferred in INPUT_COLUMNS:
        if preferred not in headers:
            headers.append(preferred)
    for mp in MARKETPLACES:
        for suffix in ["status", "source", "note", "confidence", "category_name", "node_id", "source_url"]:
            col = f"{mp}_{suffix}"
            if col not in headers:
                headers.append(col)
    for row in rows:
        for key in row.keys():
            if key not in headers:
                headers.append(key)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: clean(row.get(k, "")) for k in headers})


def read_mapping_rows(path: Path) -> List[Dict[str, str]]:
    rows = read_generic_rows(path, sheet_name="MAPPINGS", find_header=True, required=REQUIRED_MAPPING_COLUMNS)
    if not rows:
        raise ValueError("MAPPINGS tab je prazan ili nije pronađen nijedan podatkovni red.")
    first_keys = set(rows[0].keys())
    missing = [c for c in REQUIRED_MAPPING_COLUMNS if c not in first_keys]
    if missing:
        found = ", ".join(rows[0].keys())
        raise ValueError("U MAPPINGS tabu fale stupci: " + ", ".join(missing) + ". Pronađeni stupci su: " + found)
    return [r for r in rows if clean(r.get("Node ID"))]


def build_mapping_index(mapping_rows: List[Dict[str, str]]) -> Dict[str, Dict[str, str]]:
    index: Dict[str, Dict[str, str]] = {}
    for row in mapping_rows:
        node_id = clean(row.get("Node ID"))
        if node_id:
            index[node_id] = {col: clean(row.get(col)) for col in REQUIRED_MAPPING_COLUMNS}
    return index


def build_category_index(rows: List[Dict[str, str]]) -> Dict[Tuple[str, str], Dict[str, str]]:
    idx: Dict[Tuple[str, str], Dict[str, str]] = {}
    for r in rows:
        mp = clean(r.get("marketplace")).upper()
        node_id = clean(r.get("node_id"))
        if mp and node_id:
            idx[(mp, node_id)] = r
    return idx


def infer_marketplace_from_text(text: str, fallback: str = "") -> str:
    t = clean(text).lower()
    explicit = clean(fallback).upper()
    if explicit and explicit != "AUTO":
        return explicit
    # Amazon BTG sheet names and filenames often contain prefixes like fr-grocery or de-food.
    for code in ["FR", "DE", "IT", "ES", "UK", "NL", "PL", "IE", "SE"]:
        c = code.lower()
        if t.startswith(c + "-") or t.startswith(c + "_") or f"/{c}-" in t or f"_{c}_" in t or f"-{c}-" in t:
            return code
        if c in {"fr", "de", "it", "es", "nl", "pl", "ie", "se", "uk"} and re.search(rf"(^|[^a-z]){c}([^a-z]|$)", t):
            return code
    return ""


def normalize_btg_path(path: str) -> str:
    text = clean(path)
    # BTG files usually separate path levels with '/', while user's output uses ' > '.
    if "/" in text and ">" not in text:
        parts = [p.strip() for p in text.split("/") if p.strip()]
        return " > ".join(parts)
    return text


def category_name_from_path(path: str) -> str:
    text = clean(path)
    if ">" in text:
        return text.split(">")[-1].strip()
    if "/" in text:
        return text.split("/")[-1].strip()
    return text


def workbook_sheet_names(path: Path) -> List[str]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        wb = load_workbook(path, read_only=True, data_only=True)
        return list(wb.sheetnames)
    if suffix == ".xls":
        if xlrd is None:
            raise ValueError("Za .xls treba paket xlrd. Pokreni: .venv\Scripts\python -m pip install xlrd==2.0.1")
        return xlrd.open_workbook(str(path)).sheet_names()
    return [""]


def read_btg_catalog(path: Path, marketplace_hint: str = "") -> List[Dict[str, str]]:
    """
    Čita Amazon Browse Tree Guide datoteke.
    Očekuje sheet s kolonama: Node ID, Node Path, Refinement Link.
    Marketplace se može zadati ručno ili se pokušava zaključiti iz imena sheeta/datoteke
    (npr. fr-grocery, de-food).
    """
    result: List[Dict[str, str]] = []
    suffix = path.suffix.lower()
    if suffix in {".csv", ".txt"}:
        rows = read_csv_rows(path, find_header=True, required=["Node ID", "Node Path"])
        mp = infer_marketplace_from_text(path.name, marketplace_hint)
        for r in rows:
            node_id = clean(r.get("Node ID") or r.get("node_id"))
            raw_path = clean(r.get("Node Path") or r.get("node_path"))
            if node_id and raw_path and mp:
                norm_path = normalize_btg_path(raw_path)
                result.append({"marketplace": mp, "node_id": node_id, "node_path": norm_path, "category_name": category_name_from_path(norm_path), "source": path.name})
        return result

    for sheet in workbook_sheet_names(path):
        low = sheet.lower()
        if low in {"instructions", "refinements", "values", "anleitung"}:
            continue
        try:
            rows = read_generic_rows(path, sheet_name=sheet, find_header=True, required=["Node ID", "Node Path"])
        except Exception:
            continue
        mp = infer_marketplace_from_text(sheet, marketplace_hint) or infer_marketplace_from_text(path.name, marketplace_hint)
        if not mp:
            continue
        for r in rows:
            node_id = clean(r.get("Node ID") or r.get("node_id"))
            raw_path = clean(r.get("Node Path") or r.get("node_path"))
            if not node_id or not raw_path:
                continue
            # Preskoči slučajno krivo pročitane refinement/value tabove
            if node_id.lower() == "node id" or raw_path.lower() == "node path":
                continue
            norm_path = normalize_btg_path(raw_path)
            result.append({
                "marketplace": mp,
                "node_id": node_id,
                "node_path": norm_path,
                "category_name": category_name_from_path(norm_path),
                "source": f"{path.name}::{sheet}",
            })
    return result


def read_category_catalog(path: Path | None, marketplace_hint: str = "") -> List[Dict[str, str]]:
    if path is None or not path.exists() or path.stat().st_size == 0:
        return []

    # 1) Prvo probaj normalizirani catalog format:
    # marketplace,node_id,node_path,category_name
    try:
        rows = read_generic_rows(path, find_header=True, required=["marketplace", "node_id"])
        result: List[Dict[str, str]] = []
        for r in rows:
            mp = clean(r.get("marketplace")).upper() or infer_marketplace_from_text(path.name, marketplace_hint)
            node_id = clean(r.get("node_id") or r.get("Node ID") or r.get("Browse Node ID"))
            category_name = clean(r.get("category_name"))
            node_path = clean(r.get("node_path") or r.get("Node Path"))
            if not category_name and node_path:
                category_name = category_name_from_path(node_path)
            if mp and node_id and (category_name or node_path):
                result.append({"marketplace": mp, "node_id": node_id, "category_name": category_name, "node_path": normalize_btg_path(node_path or category_name), "source": path.name})
        if result:
            return result
    except Exception:
        pass

    # 2) Ako nije normalizirani catalog, tretiraj kao Amazon BTG.
    return read_btg_catalog(path, marketplace_hint=marketplace_hint)


def get_category_index_from_db() -> Dict[Tuple[str, str], Dict[str, str]]:
    idx: Dict[Tuple[str, str], Dict[str, str]] = {}
    if not DB_PATH.exists():
        return idx
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        try:
            rows = con.execute("SELECT marketplace, node_id, node_path, category_name, source_file FROM category_catalog").fetchall()
        except Exception:
            return idx
        for item in rows:
            mp = clean(item["marketplace"]).upper()
            node_id = clean(item["node_id"])
            if mp and node_id:
                idx[(mp, node_id)] = {
                    "marketplace": mp,
                    "node_id": node_id,
                    "node_path": clean(item["node_path"]),
                    "category_name": clean(item["category_name"]),
                    "source": clean(item["source_file"]),
                }
    return idx


def import_category_catalog_file(path: Path, marketplace_hint: str = "") -> Dict[str, Any]:
    rows = read_category_catalog(path, marketplace_hint=marketplace_hint)
    now = datetime.utcnow().isoformat(timespec="seconds")
    inserted = 0
    updated = 0
    skipped = 0
    with sqlite3.connect(DB_PATH) as con:
        for r in rows:
            mp = clean(r.get("marketplace")).upper()
            node_id = clean(r.get("node_id"))
            node_path = clean(r.get("node_path"))
            category_name = clean(r.get("category_name")) or category_name_from_path(node_path)
            source_file = clean(r.get("source")) or path.name
            if not mp or not node_id or not node_path:
                skipped += 1
                continue
            existed = con.execute("SELECT 1 FROM category_catalog WHERE marketplace=? AND node_id=?", (mp, node_id)).fetchone() is not None
            con.execute("""
            INSERT INTO category_catalog (marketplace, node_id, node_path, category_name, source_file, imported_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(marketplace, node_id) DO UPDATE SET
                node_path=excluded.node_path,
                category_name=excluded.category_name,
                source_file=excluded.source_file,
                updated_at=excluded.updated_at
            """, (mp, node_id, node_path, category_name, source_file, now, now))
            if existed:
                updated += 1
            else:
                inserted += 1
        con.commit()
    return {"total_read": len(rows), "inserted": inserted, "updated": updated, "skipped": skipped}


def catalog_stats() -> List[Dict[str, Any]]:
    init_db()
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        try:
            items = con.execute("SELECT marketplace, COUNT(*) AS cnt FROM category_catalog GROUP BY marketplace ORDER BY marketplace").fetchall()
        except Exception:
            return []
        return [dict(x) for x in items]


def is_supported_catalog_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in {".xls", ".xlsx", ".xlsm", ".csv", ".txt"}


def marketplace_from_folder(path: Path) -> str:
    # Ako korisnik složi folder npr. BTG/FR/fr_garden.xls, parent folder FR je dobar hint.
    for part in reversed(path.parts[:-1]):
        mp = clean(part).upper()
        if mp in {"FR", "DE", "IT", "ES", "UK", "NL", "PL", "IE", "SE"}:
            return mp
    return ""


def import_category_catalog_folder(folder_path: Path, marketplace_hint: str = "AUTO") -> Dict[str, Any]:
    if not folder_path.exists() or not folder_path.is_dir():
        raise ValueError(f"Folder ne postoji ili nije folder: {folder_path}")

    files = [p for p in folder_path.rglob("*") if is_supported_catalog_file(p)]
    if not files:
        raise ValueError("U odabranom folderu nisam pronašao .xls/.xlsx/.xlsm/.csv BTG/category datoteke.")

    imported_files: List[Dict[str, Any]] = []
    total_read = inserted = updated = skipped = failed = 0

    for file_path in sorted(files):
        hint = clean(marketplace_hint).upper()
        if not hint or hint == "AUTO":
            hint = marketplace_from_folder(file_path) or "AUTO"
        try:
            res = import_category_catalog_file(file_path, marketplace_hint=hint)
            total_read += int(res.get("total_read", 0))
            inserted += int(res.get("inserted", 0))
            updated += int(res.get("updated", 0))
            skipped += int(res.get("skipped", 0))
            imported_files.append({"file": str(file_path), "ok": True, **res})
        except Exception as exc:
            failed += 1
            imported_files.append({"file": str(file_path), "ok": False, "error": str(exc)})

    return {
        "folder": str(folder_path),
        "files_found": len(files),
        "files_failed": failed,
        "total_read": total_read,
        "inserted": inserted,
        "updated": updated,
        "skipped": skipped,
        "stats": catalog_stats(),
        "files": imported_files[:200],
    }


def build_category_index_from_input(rows: List[Dict[str, str]]) -> Dict[Tuple[str, str], Dict[str, str]]:
    """
    Ako ulazni CSV/XLSX već ima neke provjerene nazive kategorija, iskoristi ih kao lokalni catalog.

    Primjer korisnikovog formata:
      Node Id in FR = 3575160031
      FR = Cuisine et Maison > Produits et accessoires de nettoyage > Seaux

    Ovo NIJE izmišljanje: naziv se koristi samo ako već postoji u ulaznoj datoteci ili kasnije
    u korisnikovoj potvrđenoj/ispravljenoj datoteci.
    """
    idx: Dict[Tuple[str, str], Dict[str, str]] = {}
    for row in rows:
        for mp in ["FR", "IT", "ES"]:
            node_id = clean(row.get(f"Node Id in {mp}")) or clean(row.get(f"{mp}_node_id"))
            value = clean(row.get(mp))
            if not node_id or not value:
                continue
            category_name = value.split(">")[-1].strip() if ">" in value else value
            idx[(mp, node_id)] = {
                "marketplace": mp,
                "node_id": node_id,
                "category_name": category_name,
                "node_path": value,
            }
    return idx


def lookup_category_name(category_index: Dict[Tuple[str, str], Dict[str, str]], marketplace: str, node_id: str) -> Tuple[str, str]:
    rec = category_index.get((marketplace.upper(), clean(node_id)))
    if not rec:
        return "", ""
    path = clean(rec.get("node_path"))
    name = clean(rec.get("category_name"))
    # U glavnom stupcu FR/IT/ES bolje je imati puni path ako postoji.
    return path or name, name


def ensure_output_columns(row: Dict[str, Any], marketplaces: List[str]) -> None:
    for mp in marketplaces:
        row.setdefault(mp, "")
        row.setdefault(f"{mp}_node_id", "")
        for suffix in ["status", "source", "note", "confidence", "category_name", "source_url"]:
            row.setdefault(f"{mp}_{suffix}", "")


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS learning_mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            marketplace TEXT NOT NULL,
            input_signature TEXT NOT NULL,
            source_node_id TEXT,
            source_category_name TEXT,
            product_type TEXT,
            pim_category_name TEXT,
            ean TEXT,
            target_value TEXT NOT NULL,
            target_category_name TEXT,
            target_node_id TEXT,
            status TEXT NOT NULL,
            confidence REAL,
            note TEXT,
            source TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            usage_count INTEGER NOT NULL DEFAULT 0,
            UNIQUE(marketplace, input_signature)
        )
        """)
        con.execute("""
        CREATE TABLE IF NOT EXISTS category_catalog (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            marketplace TEXT NOT NULL,
            node_id TEXT NOT NULL,
            node_path TEXT NOT NULL,
            category_name TEXT,
            source_file TEXT,
            imported_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(marketplace, node_id)
        )
        """)
        # Dodaj target_node_id ako postoji stara baza
        try:
            con.execute("ALTER TABLE learning_mappings ADD COLUMN target_node_id TEXT")
        except Exception:
            pass
        con.commit()


def build_signature(row: Dict[str, Any]) -> str:
    parts = [
        clean(row.get("product type")).lower(),
        clean(row.get("NODE ID")).lower(),
        clean(row.get("Amazon category name")).lower(),
        clean(row.get("category name eng")).lower(),
        clean(row.get("PIM category name")).lower(),
    ]
    return " | ".join(parts)


def find_exact_mapping(row: Dict[str, Any], marketplace: str) -> Dict[str, Any] | None:
    signature = build_signature(row)
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        cur = con.execute("SELECT * FROM learning_mappings WHERE marketplace=? AND input_signature=?", (marketplace, signature))
        item = cur.fetchone()
        if item:
            con.execute("UPDATE learning_mappings SET usage_count=usage_count+1 WHERE id=?", (item["id"],))
            con.commit()
            return dict(item)
    return None


def token_sort_ratio(a: str, b: str) -> float:
    at = " ".join(sorted(a.lower().split()))
    bt = " ".join(sorted(b.lower().split()))
    return SequenceMatcher(None, at, bt).ratio() * 100


def find_similar_mapping(row: Dict[str, Any], marketplace: str, min_score: float = 94) -> Tuple[Dict[str, Any] | None, float]:
    signature = build_signature(row)
    best: Dict[str, Any] | None = None
    best_score = 0.0
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        for item in con.execute("SELECT * FROM learning_mappings WHERE marketplace=?", (marketplace,)):
            score = token_sort_ratio(signature, item["input_signature"])
            if score > best_score:
                best = dict(item)
                best_score = score
        if best and best_score >= min_score:
            con.execute("UPDATE learning_mappings SET usage_count=usage_count+1 WHERE id=?", (best["id"],))
            con.commit()
            return best, best_score
    return None, best_score


def save_mapping(row: Dict[str, Any], marketplace: str, target_value: str, target_category_name: str, target_node_id: str, confidence: float, status: str, note: str, source: str = "USER_CORRECTION") -> None:
    now = datetime.utcnow().isoformat(timespec="seconds")
    signature = build_signature(row)
    data = {
        "marketplace": marketplace,
        "input_signature": signature,
        "source_node_id": clean(row.get("NODE ID")),
        "source_category_name": clean(row.get("Amazon category name")),
        "product_type": clean(row.get("product type")),
        "pim_category_name": clean(row.get("PIM category name")),
        "ean": clean(row.get("EAN")),
        "target_value": clean(target_value),
        "target_category_name": clean(target_category_name),
        "target_node_id": clean(target_node_id),
        "status": status,
        "confidence": confidence,
        "note": note,
        "source": source,
        "created_by": "winforms_user",
        "now": now,
    }
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
        INSERT INTO learning_mappings (
            marketplace, input_signature, source_node_id, source_category_name, product_type, pim_category_name, ean,
            target_value, target_category_name, target_node_id, status, confidence, note, source, created_by, created_at, updated_at, usage_count
        ) VALUES (
            :marketplace, :input_signature, :source_node_id, :source_category_name, :product_type, :pim_category_name, :ean,
            :target_value, :target_category_name, :target_node_id, :status, :confidence, :note, :source, :created_by, :now, :now, 0
        )
        ON CONFLICT(marketplace, input_signature) DO UPDATE SET
            target_value=excluded.target_value,
            target_category_name=excluded.target_category_name,
            target_node_id=excluded.target_node_id,
            status=excluded.status,
            confidence=excluded.confidence,
            note=excluded.note,
            source=excluded.source,
            updated_at=excluded.updated_at
        """, data)
        con.commit()


def direct_mapping(mapping_index: Dict[str, Dict[str, str]], category_index: Dict[Tuple[str, str], Dict[str, str]], source_node_id: str, marketplace: str) -> Dict[str, str]:
    node_id = clean(source_node_id)
    if not node_id:
        return {"target_node_id": "", "category_value": "", "status": "MISSING_SOURCE_NODE", "note": "U ulaznom CSV-u fali NODE ID.", "category_name": "", "confidence": "0"}
    record = mapping_index.get(node_id)
    if not record:
        return {"target_node_id": "", "category_value": "", "status": "NO_DIRECT_MAPPING", "note": "NODE ID nije pronađen u MAPPINGS tabu.", "category_name": "", "confidence": "0"}
    target_col = DIRECT_COLUMNS.get(marketplace)
    target_id = clean(record.get(target_col)) if target_col else ""
    if not target_id:
        return {"target_node_id": "", "category_value": "", "status": "MISSING_TARGET_NODE", "note": f"Pronađen source node, ali nema vrijednosti u stupcu '{target_col}'.", "category_name": "", "confidence": "0"}

    category_value, category_name = lookup_category_name(category_index, marketplace, target_id)
    if category_value:
        return {
            "target_node_id": target_id,
            "category_value": category_value,
            "status": "DIRECT_MAPPING_WITH_NAME",
            "note": "Node ID je iz European Browse Node Mapping tablice, a naziv/path iz učitanog category catalog/BTG izvora.",
            "category_name": category_name,
            "confidence": "1",
        }
    return {
        "target_node_id": target_id,
        "category_value": "",
        "status": "DIRECT_NODE_MAPPING_NAME_MISSING",
        "note": f"Node ID je pronađen ({target_id}), ali naziv kategorije nije popunjen jer u lokalnoj BTG/category_catalog bazi nema zapisa za {marketplace} node {target_id}. Uvezi odgovarajući {marketplace} BTG ili provjeri da je BTG folder odabran.",
        "category_name": "",
        "confidence": "0.7",
    }


def openai_guess(row: Dict[str, Any], marketplace: str) -> Dict[str, Any]:
    env = load_env()
    api_key = env.get("OPENAI_API_KEY", "")
    model = env.get("OPENAI_MODEL", "gpt-4.1-mini")
    if not api_key:
        return {"target_node_id": "", "category_value": "", "status": "AI_DISABLED", "reason": "OPENAI_API_KEY nije postavljen.", "confidence": 0, "category_name": "", "source_url": ""}

    prompt = f"""
Trebam pomoć za Amazon category/browse node mapping za marketplace {marketplace}.
Vrati strogi JSON s poljima: target_node_id, category_value, category_name, confidence, status, reason, source_url.

PRAVILA:
- Ne izmišljaj ni node ID ni naziv kategorije.
- Ako ne možeš potvrditi da kategorija stvarno postoji na Amazonu ili službenom Amazon/Seller Central izvoru, status mora biti NEED_REVIEW i target_node_id može biti prazan.
- source_url mora biti URL izvora gdje je kategorija pronađena.
- Ako je rezultat samo procjena, status mora biti NEED_REVIEW, ne CONFIRMED.

Ulazni podaci:
product type: {clean(row.get('product type'))}
DE/source NODE ID: {clean(row.get('NODE ID'))}
Amazon category name DE: {clean(row.get('Amazon category name'))}
category name eng: {clean(row.get('category name eng'))}
PIM category name: {clean(row.get('PIM category name'))}
EAN: {clean(row.get('EAN'))}
""".strip()
    payload = {
        "model": model,
        "input": prompt,
        "tools": [{"type": "web_search_preview"}],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "amazon_category_guess",
                "strict": True,
                "schema": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "target_node_id": {"type": "string"},
                        "category_value": {"type": "string"},
                        "category_name": {"type": "string"},
                        "confidence": {"type": "number"},
                        "status": {"type": "string", "enum": ["AI_WEB_MATCH", "MATCH_MEDIUM", "NEED_REVIEW", "NO_MATCH"]},
                        "reason": {"type": "string"},
                        "source_url": {"type": "string"},
                    },
                    "required": ["target_node_id", "category_value", "category_name", "confidence", "status", "reason", "source_url"],
                },
            }
        },
    }
    try:
        req = urllib.request.Request(
            "https://api.openai.com/v1/responses",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        text = ""
        for item in data.get("output", []):
            for content in item.get("content", []):
                if content.get("type") in {"output_text", "text"}:
                    text += content.get("text", "")
        parsed = json.loads(text) if text else {}
        parsed.setdefault("target_node_id", "")
        parsed.setdefault("category_value", "")
        parsed.setdefault("category_name", "")
        parsed.setdefault("confidence", 0)
        parsed.setdefault("status", "NEED_REVIEW")
        parsed.setdefault("reason", "")
        parsed.setdefault("source_url", "")
        # Oprez: AI rezultat uvijek ostavljamo za pregled osim ako korisnik kasnije potvrdi.
        if parsed.get("status") == "AI_WEB_MATCH":
            parsed["status"] = "AI_NEEDS_REVIEW"
            parsed["reason"] = "AI je pronašao mogući izvor, ali rezultat nije službeni import iz BTG/API-ja; ručno potvrditi. " + clean(parsed.get("reason"))
        return parsed
    except Exception as exc:
        return {"target_node_id": "", "category_value": "", "status": "NEED_REVIEW", "reason": f"AI fallback nije uspio: {exc}", "confidence": 0, "category_name": "", "source_url": ""}


def set_marketplace_result(row: Dict[str, Any], mp: str, target_node_id: str, category_value: str, category_name: str, status: str, source: str, note: str, confidence: str | float, source_url: str = "") -> None:
    # Glavni stupac FR/IT/ES = naziv/path kategorije
    row[mp] = category_value
    row[f"{mp}_node_id"] = target_node_id
    row[f"{mp}_category_name"] = category_name or category_value
    row[f"{mp}_status"] = status
    row[f"{mp}_source"] = source
    row[f"{mp}_note"] = note
    row[f"{mp}_confidence"] = clean(confidence)
    row[f"{mp}_source_url"] = source_url
    node_col = f"Node Id in {mp}"
    if node_col in row:
        row[node_col] = target_node_id


def process_file(input_path: Path, mapping_path: Path, output_path: Path, marketplaces: List[str], use_ai: bool, overwrite_existing: bool, max_ai_rows: int, category_path: Path | None = None) -> None:
    selected = [m.upper().strip() for m in marketplaces if m.upper().strip() in MARKETPLACES]
    rows = read_input_rows(input_path)
    mapping_index = build_mapping_index(read_mapping_rows(mapping_path))

    # Nazivi kategorija NE dolaze iz European mapping tablice; ona daje node ID-eve.
    # Zato ih punimo samo iz provjerenih izvora:
    # 1) već postojeći nazivi u ulaznom CSV/XLSX-u, ako ih ima
    # 2) opcionalni Category catalog/BTG file koji korisnik učita
    category_index = get_category_index_from_db()
    category_index.update(build_category_index_from_input(rows))
    category_index.update(build_category_index(read_category_catalog(category_path)))
    ai_calls = 0

    for row in rows:
        ensure_output_columns(row, selected)
        for mp in selected:
            current_value = clean(row.get(mp))
            if current_value and not overwrite_existing:
                row[f"{mp}_status"] = "ALREADY_FILLED"
                row[f"{mp}_source"] = "INPUT"
                continue

            learned = find_exact_mapping(row, mp)
            if learned:
                set_marketplace_result(
                    row, mp,
                    target_node_id=clean(learned.get("target_node_id")),
                    category_value=clean(learned.get("target_value")),
                    category_name=clean(learned.get("target_category_name")),
                    status="LEARNED_MATCH",
                    source="LEARNING_DB",
                    note=clean(learned.get("note")) or "Pronađeno u learning bazi.",
                    confidence=learned.get("confidence", 1),
                )
                continue

            similar, score = find_similar_mapping(row, mp)
            if similar:
                set_marketplace_result(
                    row, mp,
                    target_node_id=clean(similar.get("target_node_id")),
                    category_value=clean(similar.get("target_value")),
                    category_name=clean(similar.get("target_category_name")),
                    status="LEARNED_SIMILAR",
                    source="LEARNING_DB",
                    note=f"Slično naučeno mapiranje, score={round(score, 1)}. Provjeriti ako je važno.",
                    confidence=round(score / 100, 3),
                )
                continue

            if mp in DIRECT_MARKETPLACES:
                result = direct_mapping(mapping_index, category_index, row.get("NODE ID", ""), mp)
                set_marketplace_result(
                    row, mp,
                    target_node_id=result["target_node_id"],
                    category_value=result["category_value"],
                    category_name=result["category_name"],
                    status=result["status"],
                    source="EUROPEAN_BROWSE_NODE_MAPPING" if result["status"].startswith("DIRECT") else "MAPPING_TABLE",
                    note=result["note"],
                    confidence=result["confidence"],
                )
                continue

            if mp in AI_FALLBACK_MARKETPLACES and use_ai and ai_calls < max_ai_rows:
                ai_calls += 1
                guess = openai_guess(row, mp)
                set_marketplace_result(
                    row, mp,
                    target_node_id=clean(guess.get("target_node_id")),
                    category_value=clean(guess.get("category_value")) or clean(guess.get("category_name")),
                    category_name=clean(guess.get("category_name")),
                    status=clean(guess.get("status", "NEED_REVIEW")),
                    source="OPENAI_WEB_SEARCH",
                    note=clean(guess.get("reason")),
                    confidence=clean(guess.get("confidence", 0)),
                    source_url=clean(guess.get("source_url")),
                )
            else:
                set_marketplace_result(
                    row, mp,
                    target_node_id="",
                    category_value="",
                    category_name="",
                    status="NEED_REVIEW" if mp in AI_FALLBACK_MARKETPLACES else "UNSUPPORTED",
                    source="NONE",
                    note="Nema direktnog mapping stupca. Uključi AI fallback ili ručno ispravi." if mp in AI_FALLBACK_MARKETPLACES else "Nepodržana država.",
                    confidence="0",
                )

    write_csv_rows(rows, output_path)


def save_corrections(rows: List[Dict[str, Any]], marketplaces: List[str]) -> int:
    count = 0
    for row in rows:
        for mp in marketplaces:
            value = clean(row.get(mp))
            node_id = clean(row.get(f"{mp}_node_id")) or clean(row.get(f"Node Id in {mp}"))
            status = clean(row.get(f"{mp}_status"))
            category_name = clean(row.get(f"{mp}_category_name")) or value
            note = clean(row.get(f"{mp}_note"))
            if status not in {"CONFIRMED", "CORRECTED"}:
                continue
            if not value and not category_name:
                continue
            try:
                confidence = float(clean(row.get(f"{mp}_confidence", "1")).replace(",", "."))
            except Exception:
                confidence = 1.0
            save_mapping(row, mp, value or category_name, category_name, node_id, confidence, status, note)
            count += 1
    return count


def _parse_content_disposition(headers: str) -> Tuple[str | None, str | None]:
    cd_line = ""
    for line in headers.splitlines():
        if line.lower().startswith("content-disposition:"):
            cd_line = line
            break
    if not cd_line:
        return None, None
    name = None
    filename = None
    parts = [p.strip() for p in cd_line.split(";")]
    for part in parts:
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        key = key.strip().lower()
        value = value.strip().strip('"')
        if key == "name":
            name = value
        elif key == "filename":
            filename = value
    return name, filename


def parse_multipart(body: bytes, content_type: str) -> Tuple[Dict[str, str], Dict[str, Tuple[str, bytes]]]:
    m = re.search(r"boundary=(?P<b>[^;]+)", content_type)
    if not m:
        raise ValueError(f"Nedostaje multipart boundary. Content-Type: {content_type}")
    boundary = m.group("b").strip().strip('"').encode("utf-8")
    delimiter = b"--" + boundary
    fields: Dict[str, str] = {}
    files: Dict[str, Tuple[str, bytes]] = {}
    for part in body.split(delimiter):
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        if b"\r\n\r\n" in part:
            raw_headers, content = part.split(b"\r\n\r\n", 1)
        elif b"\n\n" in part:
            raw_headers, content = part.split(b"\n\n", 1)
        else:
            continue
        headers = raw_headers.decode("utf-8", errors="ignore")
        name, filename = _parse_content_disposition(headers)
        if not name:
            continue
        if content.endswith(b"\r\n"):
            content = content[:-2]
        elif content.endswith(b"\n"):
            content = content[:-1]
        if filename is not None:
            files[name] = (filename, content)
        else:
            fields[name] = content.decode("utf-8", errors="replace")
    return fields, files


class Handler(BaseHTTPRequestHandler):
    server_version = "AmazonCategoryMapperLite/1.3"

    def _json(self, status: int, data: Any) -> None:
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:
        if self.path.startswith("/health"):
            self._json(200, {"status": "ok", "backend": "lite-no-pandas-btgcatalog"})
        elif self.path.startswith("/api/catalog-stats"):
            self._json(200, {"items": catalog_stats()})
        else:
            self._json(404, {"detail": "Not found"})

    def do_POST(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            if self.path.startswith("/api/map-csv"):
                self.handle_map_csv(body)
            elif self.path.startswith("/api/import-category-catalog"):
                self.handle_import_category_catalog(body)
            elif self.path.startswith("/api/import-category-folder"):
                self.handle_import_category_folder(body)
            elif self.path.startswith("/api/save-corrections"):
                data = json.loads(body.decode("utf-8"))
                rows = data.get("rows", [])
                marketplaces = [clean(m).upper() for m in data.get("marketplaces", MARKETPLACES)]
                saved = save_corrections(rows, marketplaces)
                self._json(200, {"saved": saved})
            else:
                self._json(404, {"detail": "Not found"})
        except Exception as exc:
            self._json(500, {"detail": str(exc)})

    def handle_import_category_catalog(self, body: bytes) -> None:
        fields, files = parse_multipart(body, self.headers.get("Content-Type", ""))
        category_file = None
        for name in ["category_file", "categoryCatalog", "category_catalog", "btg_file", "catalog_file", "file"]:
            if name in files:
                category_file = files[name]
                break
        if category_file is None:
            available_files = ", ".join(files.keys()) if files else "nema poslanih file polja"
            raise ValueError("Nedostaje BTG/category datoteka. Očekujem file polje 'category_file'. Dobiveno: " + available_files)
        marketplace_hint = clean(fields.get("marketplace", "AUTO")).upper()
        category_name, category_content = category_file
        category_suffix = Path(category_name or "category_catalog.xls").suffix or ".xls"
        category_path = TEMP_DIR / f"import_category_{os.getpid()}_{datetime.utcnow().timestamp()}{category_suffix}"
        category_path.write_bytes(category_content)
        result = import_category_catalog_file(category_path, marketplace_hint=marketplace_hint)
        result["stats"] = catalog_stats()
        self._json(200, result)

    def handle_import_category_folder(self, body: bytes) -> None:
        try:
            data = json.loads(body.decode("utf-8")) if body else {}
        except Exception:
            data = {}
        folder_text = clean(data.get("folder_path")) or clean(data.get("folder"))
        marketplace_hint = clean(data.get("marketplace", "AUTO")).upper()
        if not folder_text:
            raise ValueError("Nedostaje folder_path. Pošalji lokalnu putanju do foldera s BTG datotekama.")
        result = import_category_catalog_folder(Path(folder_text), marketplace_hint=marketplace_hint)
        self._json(200, result)

    def handle_map_csv(self, body: bytes) -> None:
        fields, files = parse_multipart(body, self.headers.get("Content-Type", ""))
        def pick_file(*names: str):
            for name in names:
                if name in files:
                    return files[name]
            return None
        input_file = pick_file("input_file", "input", "inputFile", "csv_file", "file")
        mapping_file = pick_file("mapping_file", "mapping", "mappingFile", "mapping_excel", "amazon_mapping")
        category_file = pick_file("category_file", "categoryCatalog", "category_catalog", "btg_file", "catalog_file")
        if input_file is None or mapping_file is None:
            available_files = ", ".join(files.keys()) if files else "nema poslanih file polja"
            available_fields = ", ".join(fields.keys()) if fields else "nema poslanih običnih polja"
            raise ValueError("Nedostaju uploadane datoteke. Backend očekuje file polja 'input_file' i 'mapping_file'. " + f"Dobivena file polja: {available_files}. Dobivena ostala polja: {available_fields}.")
        selected = [m.strip().upper() for m in fields.get("marketplaces", "FR,IT,ES").split(",") if m.strip()]
        use_ai = fields.get("use_ai", "false").lower() == "true"
        overwrite_existing = fields.get("overwrite_existing", "true").lower() == "true"
        btg_folder_text = clean(fields.get("btg_folder_path") or fields.get("btg_folder") or fields.get("category_folder"))
        btg_marketplace_hint = clean(fields.get("btg_marketplace") or fields.get("catalog_marketplace") or "AUTO").upper()
        try:
            max_ai_rows = int(fields.get("max_ai_rows", "50"))
        except Exception:
            max_ai_rows = 50
        input_name, input_content = input_file
        mapping_name, mapping_content = mapping_file
        input_suffix = Path(input_name or "input.csv").suffix or ".csv"
        mapping_suffix = Path(mapping_name or "mapping.xlsx").suffix or ".xlsx"
        input_path = TEMP_DIR / f"input_{os.getpid()}_{datetime.utcnow().timestamp()}{input_suffix}"
        mapping_path = TEMP_DIR / f"mapping_{os.getpid()}_{datetime.utcnow().timestamp()}{mapping_suffix}"
        category_path = None
        output_path = TEMP_DIR / f"amazon_categories_mapped_{os.getpid()}.csv"
        input_path.write_bytes(input_content)
        mapping_path.write_bytes(mapping_content)
        if category_file is not None:
            category_name, category_content = category_file
            category_suffix = Path(category_name or "category_catalog.csv").suffix or ".csv"
            category_path = TEMP_DIR / f"category_{os.getpid()}_{datetime.utcnow().timestamp()}{category_suffix}"
            category_path.write_bytes(category_content)

        # Ako je korisnik odabrao BTG folder u WinForms aplikaciji, automatski ga uvezi/azuriraj
        # prije mapiranja. Ovo sprjecava situaciju gdje postoje BTG datoteke, ali nisu spremljene
        # u lokalnu SQLite category_catalog bazu.
        if btg_folder_text:
            btg_path = Path(btg_folder_text)
            if btg_path.exists() and btg_path.is_dir():
                import_category_catalog_folder(btg_path, marketplace_hint=btg_marketplace_hint or "AUTO")

        process_file(input_path, mapping_path, output_path, selected, use_ai, overwrite_existing, max_ai_rows, category_path=category_path)
        payload = output_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", 'attachment; filename="amazon_categories_mapped.csv"')
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def main() -> None:
    init_db()
    print(f"Backend bez pandas/numpy pokrenut na http://{HOST}:{PORT}")
    print("Verzija: BTG folder import + category catalog import")
    print("Zaustavljanje: CTRL+C")
    HTTPServer((HOST, PORT), Handler).serve_forever()


if __name__ == "__main__":
    main()
