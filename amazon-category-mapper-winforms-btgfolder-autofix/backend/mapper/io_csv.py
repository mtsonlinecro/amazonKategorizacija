from __future__ import annotations

from pathlib import Path
import pandas as pd

INPUT_COLUMNS = [
    "product type",
    "NODE ID",
    "Amazon category name",
    "Node Id in UK",
    "Node Id in FR",
    "Node Id in IT",
    "Node Id in ES",
    "category name eng",
    "PIM category name",
    "EAN",
    "FR",
    "IT",
    "NL",
    "ES",
    "PL",
    "IE",
    "SE",
]

MARKETPLACES = ["FR", "IT", "NL", "ES", "PL", "IE", "SE"]


def read_input_file(path: str | Path) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        df = pd.read_excel(p, dtype=str)
    else:
        # Najčešće radi za CSV iz Excela. Ako je separator ;, pandas će ga uz sep=None pogoditi.
        df = pd.read_csv(p, dtype=str, encoding="utf-8-sig", sep=None, engine="python")
    df.columns = [str(c).strip() for c in df.columns]
    for col in INPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df.fillna("")


def write_csv(df: pd.DataFrame, path: str | Path) -> None:
    df.to_csv(path, index=False, encoding="utf-8-sig")
