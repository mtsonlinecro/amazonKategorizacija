from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import List, Dict, Any

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from mapper.models import init_db
from mapper.processor import process_file, save_corrections_from_rows
from mapper.learning_store import list_mappings

app = FastAPI(title="Amazon Category Mapper Backend", version="1.0.0")
init_db()

TEMP_DIR = Path(tempfile.gettempdir()) / "amazon_category_mapper_backend"
TEMP_DIR.mkdir(parents=True, exist_ok=True)


class CorrectionsRequest(BaseModel):
    rows: List[Dict[str, Any]]
    marketplaces: List[str] = ["FR", "IT", "NL", "ES", "PL", "IE", "SE"]


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/api/map-csv")
async def map_csv(
    input_file: UploadFile = File(...),
    mapping_file: UploadFile = File(...),
    marketplaces: str = Form("FR,IT,ES"),
    use_ai: bool = Form(False),
    overwrite_existing: bool = Form(True),
    max_ai_rows: int = Form(50),
):
    try:
        selected = [m.strip().upper() for m in marketplaces.split(",") if m.strip()]
        input_suffix = Path(input_file.filename or "input.csv").suffix or ".csv"
        mapping_suffix = Path(mapping_file.filename or "mapping.xlsx").suffix or ".xlsx"

        with tempfile.NamedTemporaryFile(delete=False, suffix=input_suffix, dir=TEMP_DIR) as f:
            input_path = Path(f.name)
            f.write(await input_file.read())

        with tempfile.NamedTemporaryFile(delete=False, suffix=mapping_suffix, dir=TEMP_DIR) as f:
            mapping_path = Path(f.name)
            f.write(await mapping_file.read())

        output_path = TEMP_DIR / f"mapped_{os.getpid()}_{input_path.stem}.csv"
        process_file(
            input_file_path=input_path,
            mapping_file_path=mapping_path,
            output_file_path=output_path,
            selected_marketplaces=selected,
            use_ai=use_ai,
            overwrite_existing=overwrite_existing,
            max_ai_rows=max_ai_rows,
        )
        return FileResponse(
            output_path,
            media_type="text/csv",
            filename="amazon_categories_mapped.csv",
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/save-corrections")
def save_corrections(request: CorrectionsRequest):
    try:
        count = save_corrections_from_rows(request.rows, request.marketplaces)
        return {"saved": count}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/api/learning")
def get_learning(limit: int = 200):
    data = []
    for m in list_mappings(limit=limit):
        data.append({
            "id": m.id,
            "marketplace": m.marketplace,
            "source_node_id": m.source_node_id,
            "source_category_name": m.source_category_name,
            "target_value": m.target_value,
            "target_category_name": m.target_category_name,
            "status": m.status,
            "confidence": m.confidence,
            "usage_count": m.usage_count,
            "created_at": m.created_at.isoformat() if m.created_at else None,
        })
    return data
