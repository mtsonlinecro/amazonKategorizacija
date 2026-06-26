from __future__ import annotations

import json
import os
from typing import Any, Dict
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()


def _clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    return text


def _json_from_response(response: Any) -> Dict[str, Any]:
    # OpenAI Responses SDK obično ima output_text. Ako ne, pokušaj parsirati razne strukture.
    text = getattr(response, "output_text", None)
    if not text:
        try:
            chunks = []
            for item in getattr(response, "output", []) or []:
                for content in getattr(item, "content", []) or []:
                    t = getattr(content, "text", None)
                    if t:
                        chunks.append(t)
            text = "\n".join(chunks)
        except Exception:
            text = ""
    if not text:
        return {}
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        # ponekad model vrati ```json blok ako structured outputs nije prošao
        text = text.replace("```json", "").replace("```", "").strip()
        try:
            return json.loads(text)
        except Exception:
            return {"value": "", "category_name": "", "status": "NEED_REVIEW", "confidence": 0.0, "reason": text[:1000], "source_url": ""}


def guess_amazon_category_with_openai(row: Dict[str, Any], marketplace: str, model: str | None = None) -> Dict[str, Any]:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        return {
            "value": "",
            "category_name": "",
            "status": "AI_DISABLED",
            "confidence": 0.0,
            "reason": "OPENAI_API_KEY nije postavljen u backend/.env.",
            "source_url": "",
        }

    client = OpenAI(api_key=api_key)
    model = model or os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

    prompt = f"""
Ti si pomoćni alat za Amazon category mapping.
Zadatak: za Amazon marketplace {marketplace} pokušaj pronaći najvjerojatniji browse node ID ili barem službeni naziv kategorije.

VAŽNA PRAVILA:
- Ne smiješ izmišljati node ID.
- Ako ne možeš pronaći stvarni Amazon node ID ili dovoljno pouzdanu službenu kategoriju, vrati status NEED_REVIEW.
- Ako pronađeš samo naziv kategorije bez node ID-a, value ostavi prazno, category_name popuni i status stavi NEED_REVIEW ili MATCH_MEDIUM.
- Ako si siguran da imaš node ID iz pouzdanog izvora, status može biti AI_WEB_MATCH.
- Uvijek vrati samo JSON.

Podaci o proizvodu / izvornoj kategoriji:
product type: {_clean(row.get('product type'))}
NODE ID: {_clean(row.get('NODE ID'))}
Amazon category name DE: {_clean(row.get('Amazon category name'))}
category name eng: {_clean(row.get('category name eng'))}
PIM category name: {_clean(row.get('PIM category name'))}
EAN: {_clean(row.get('EAN'))}
Node Id in UK: {_clean(row.get('Node Id in UK'))}
Node Id in FR: {_clean(row.get('Node Id in FR'))}
Node Id in IT: {_clean(row.get('Node Id in IT'))}
Node Id in ES: {_clean(row.get('Node Id in ES'))}

Vrati JSON u ovom obliku:
{{
  "value": "node id ako je sigurno poznat, inače prazno",
  "category_name": "naziv kategorije ako je poznat",
  "status": "AI_WEB_MATCH ili MATCH_MEDIUM ili NEED_REVIEW ili NO_MATCH",
  "confidence": 0.0,
  "reason": "kratko objašnjenje",
  "source_url": "URL izvora ako postoji"
}}
""".strip()

    schema = {
        "type": "object",
        "properties": {
            "value": {"type": "string"},
            "category_name": {"type": "string"},
            "status": {"type": "string", "enum": ["AI_WEB_MATCH", "MATCH_MEDIUM", "NEED_REVIEW", "NO_MATCH"]},
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "reason": {"type": "string"},
            "source_url": {"type": "string"},
        },
        "required": ["value", "category_name", "status", "confidence", "reason", "source_url"],
        "additionalProperties": False,
    }

    try:
        response = client.responses.create(
            model=model,
            tools=[{"type": "web_search"}],
            input=prompt,
            text={
                "format": {
                    "type": "json_schema",
                    "name": "amazon_category_guess",
                    "schema": schema,
                    "strict": True,
                }
            },
        )
        data = _json_from_response(response)
    except Exception as exc:
        return {
            "value": "",
            "category_name": "",
            "status": "NEED_REVIEW",
            "confidence": 0.0,
            "reason": f"OpenAI poziv nije uspio ili web_search nije omogućen za model/account: {exc}",
            "source_url": "",
        }

    return {
        "value": _clean(data.get("value")),
        "category_name": _clean(data.get("category_name")),
        "status": _clean(data.get("status")) or "NEED_REVIEW",
        "confidence": float(data.get("confidence") or 0),
        "reason": _clean(data.get("reason")),
        "source_url": _clean(data.get("source_url")),
    }
