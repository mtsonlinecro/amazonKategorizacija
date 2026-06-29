from __future__ import annotations

import json
import os
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict

BASE_DIR = Path(__file__).resolve().parent
CACHE_DIR = BASE_DIR / "cache"
CACHE_DIR.mkdir(exist_ok=True)
CACHE_FILE = CACHE_DIR / "translation_cache.json"

# Marketplace -> jezik za prijevod.
# IE je Amazon Ireland, ali BTG je najčešće engleski pa prevodimo u EN.
MP_TO_LANG = {
    "DE": "DE",
    "FR": "FR",
    "IT": "IT",
    "ES": "ES",
    "NL": "NL",
    "PL": "PL",
    "IE": "EN",
    "SE": "SV",
}


def _clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _load_api_settings() -> Dict[str, str]:
    values: Dict[str, str] = {}

    try:
        import api_keys  # type: ignore
        for key in [
            "TRANSLATION_ENABLED",
            "TRANSLATION_PROVIDER",
            "DEEPL_API_KEY",
            "DEEPL_USE_FREE_API",
            "GOOGLE_TRANSLATE_API_KEY",
            "AZURE_TRANSLATOR_KEY",
            "AZURE_TRANSLATOR_REGION",
            "AZURE_TRANSLATOR_ENDPOINT",
            "OPENAI_API_KEY",
            "OPENAI_TRANSLATION_MODEL",
        ]:
            value = getattr(api_keys, key, "")
            if isinstance(value, bool):
                values[key] = "true" if value else "false"
            elif value is not None:
                values[key] = str(value)
    except Exception:
        pass

    # Stara ai_settings.py datoteka ostaje podržana za OpenAI key.
    try:
        import ai_settings  # type: ignore
        if not values.get("OPENAI_API_KEY"):
            values["OPENAI_API_KEY"] = _clean(getattr(ai_settings, "OPENAI_API_KEY", ""))
        if not values.get("OPENAI_TRANSLATION_MODEL"):
            values["OPENAI_TRANSLATION_MODEL"] = _clean(getattr(ai_settings, "OPENAI_MODEL", ""))
    except Exception:
        pass

    # Environment varijable imaju zadnju riječ.
    for key in list(values.keys()) + [
        "TRANSLATION_ENABLED",
        "TRANSLATION_PROVIDER",
        "DEEPL_API_KEY",
        "DEEPL_USE_FREE_API",
        "GOOGLE_TRANSLATE_API_KEY",
        "AZURE_TRANSLATOR_KEY",
        "AZURE_TRANSLATOR_REGION",
        "AZURE_TRANSLATOR_ENDPOINT",
        "OPENAI_API_KEY",
        "OPENAI_TRANSLATION_MODEL",
    ]:
        if os.environ.get(key):
            values[key] = os.environ[key]

    values.setdefault("TRANSLATION_ENABLED", "true")
    values.setdefault("TRANSLATION_PROVIDER", "auto")
    values.setdefault("DEEPL_USE_FREE_API", "true")
    values.setdefault("AZURE_TRANSLATOR_ENDPOINT", "https://api.cognitive.microsofttranslator.com")
    values.setdefault("OPENAI_TRANSLATION_MODEL", "gpt-4.1-mini")
    return values


def _is_enabled(settings: Dict[str, str]) -> bool:
    return _clean(settings.get("TRANSLATION_ENABLED", "true")).lower() in {"1", "true", "yes", "da"}


def _select_provider(settings: Dict[str, str]) -> str:
    if not _is_enabled(settings):
        return "disabled"

    requested = _clean(settings.get("TRANSLATION_PROVIDER", "auto")).lower()
    providers = ["deepl", "google", "azure", "openai"] if requested == "auto" else [requested]
    for provider in providers:
        if provider == "deepl" and _clean(settings.get("DEEPL_API_KEY")):
            return "deepl"
        if provider == "google" and _clean(settings.get("GOOGLE_TRANSLATE_API_KEY")):
            return "google"
        if provider == "azure" and _clean(settings.get("AZURE_TRANSLATOR_KEY")):
            return "azure"
        if provider == "openai" and _clean(settings.get("OPENAI_API_KEY")):
            return "openai"
    return "none"


def translation_info() -> Dict[str, str]:
    settings = _load_api_settings()
    provider = _select_provider(settings)
    return {
        "translation_enabled": "true" if _is_enabled(settings) else "false",
        "translation_provider": provider,
        "translation_cache": str(CACHE_FILE),
    }


def _load_cache() -> Dict[str, str]:
    try:
        if CACHE_FILE.exists():
            data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items()}
    except Exception:
        pass
    return {}


def _save_cache(cache: Dict[str, str]) -> None:
    try:
        CACHE_DIR.mkdir(exist_ok=True)
        CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def clear_translation_cache() -> Dict[str, Any]:
    deleted = False
    try:
        if CACHE_FILE.exists():
            CACHE_FILE.unlink()
            deleted = True
    except Exception:
        pass
    return {"deleted_translation_cache": deleted, "translation_cache": str(CACHE_FILE)}


def _target_lang_for_marketplace(marketplace: str) -> str:
    return MP_TO_LANG.get(_clean(marketplace).upper(), _clean(marketplace).upper())


def _deepl_target(lang: str) -> str:
    lang = lang.upper()
    if lang == "EN":
        return "EN-GB"
    return lang


def _google_azure_lang(lang: str) -> str:
    lang = lang.upper()
    if lang == "SV":
        return "sv"
    if lang == "EN":
        return "en"
    return lang.lower()


def _http_json(url: str, payload: Any, headers: Dict[str, str], timeout: int = 25) -> Any:
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", **headers},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _translate_deepl(text: str, source_lang: str, target_lang: str, settings: Dict[str, str]) -> str:
    api_key = _clean(settings.get("DEEPL_API_KEY"))
    use_free = _clean(settings.get("DEEPL_USE_FREE_API", "true")).lower() in {"1", "true", "yes", "da"}
    url = "https://api-free.deepl.com/v2/translate" if use_free else "https://api.deepl.com/v2/translate"
    data = {
        "auth_key": api_key,
        "text": text,
        "target_lang": _deepl_target(target_lang),
    }
    if source_lang and source_lang.upper() != "AUTO":
        data["source_lang"] = source_lang.upper()
    encoded = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=encoded, method="POST")
    with urllib.request.urlopen(req, timeout=25) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    translations = payload.get("translations") or []
    return _clean(translations[0].get("text")) if translations else ""


def _translate_google(text: str, source_lang: str, target_lang: str, settings: Dict[str, str]) -> str:
    api_key = _clean(settings.get("GOOGLE_TRANSLATE_API_KEY"))
    url = "https://translation.googleapis.com/language/translate/v2?key=" + urllib.parse.quote(api_key)
    data = {
        "q": text,
        "target": _google_azure_lang(target_lang),
        "format": "text",
    }
    if source_lang and source_lang.upper() != "AUTO":
        data["source"] = _google_azure_lang(source_lang)
    encoded = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=encoded, method="POST")
    with urllib.request.urlopen(req, timeout=25) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    translations = (((payload.get("data") or {}).get("translations")) or [])
    return _clean(translations[0].get("translatedText")) if translations else ""


def _translate_azure(text: str, source_lang: str, target_lang: str, settings: Dict[str, str]) -> str:
    key = _clean(settings.get("AZURE_TRANSLATOR_KEY"))
    region = _clean(settings.get("AZURE_TRANSLATOR_REGION"))
    endpoint = _clean(settings.get("AZURE_TRANSLATOR_ENDPOINT", "https://api.cognitive.microsofttranslator.com")).rstrip("/")
    params = {"api-version": "3.0", "to": _google_azure_lang(target_lang)}
    if source_lang and source_lang.upper() != "AUTO":
        params["from"] = _google_azure_lang(source_lang)
    url = endpoint + "/translate?" + urllib.parse.urlencode(params)
    headers = {"Ocp-Apim-Subscription-Key": key}
    if region:
        headers["Ocp-Apim-Subscription-Region"] = region
    payload = _http_json(url, [{"Text": text}], headers=headers, timeout=25)
    try:
        return _clean(payload[0]["translations"][0]["text"])
    except Exception:
        return ""


def _translate_openai(text: str, source_lang: str, target_lang: str, settings: Dict[str, str]) -> str:
    api_key = _clean(settings.get("OPENAI_API_KEY"))
    model = _clean(settings.get("OPENAI_TRANSLATION_MODEL", "gpt-4.1-mini")) or "gpt-4.1-mini"
    prompt = (
        "Prevedi samo ovaj Amazon kategorijski pojam. "
        "Vrati samo prijevod bez objašnjenja, bez navodnika i bez dodatnog teksta.\n"
        f"Izvorni jezik: {source_lang or 'AUTO'}\n"
        f"Ciljni jezik: {target_lang}\n"
        f"Pojam: {text}"
    )
    payload = {"model": model, "input": prompt}
    req = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=35) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    pieces = []
    for item in data.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"}:
                pieces.append(_clean(content.get("text")))
    return " ".join(x for x in pieces if x).strip()


def translate_category_term(text: str, source_lang: str, target_marketplace: str) -> Dict[str, str]:
    """
    Prevodi kratki kategorijski pojam, npr. zadnji dio DE patha:
    Rasenkanten -> PL/NL/SV/EN.

    Ne prevodi cijeli BTG folder. Cache čuva samo kratke prijevode pojmova.
    """
    text = _clean(text)
    if not text:
        return {"text": "", "provider": "none", "status": "empty", "note": ""}

    target_lang = _target_lang_for_marketplace(target_marketplace)
    source_lang = (_clean(source_lang) or "AUTO").upper()
    if target_lang.upper() == source_lang.upper():
        return {"text": text, "provider": "none", "status": "same_language", "note": ""}

    settings = _load_api_settings()
    provider = _select_provider(settings)
    if provider in {"disabled", "none"}:
        return {
            "text": "",
            "provider": provider,
            "status": "disabled" if provider == "disabled" else "missing_api_key",
            "note": "Prijevod nije aktivan ili nije postavljen API ključ.",
        }

    cache_key = "|".join([provider, source_lang, target_lang.upper(), text.lower()])
    cache = _load_cache()
    if cache_key in cache:
        return {"text": cache[cache_key], "provider": provider, "status": "cached", "note": ""}

    try:
        if provider == "deepl":
            translated = _translate_deepl(text, source_lang, target_lang, settings)
        elif provider == "google":
            translated = _translate_google(text, source_lang, target_lang, settings)
        elif provider == "azure":
            translated = _translate_azure(text, source_lang, target_lang, settings)
        elif provider == "openai":
            translated = _translate_openai(text, source_lang, target_lang, settings)
        else:
            translated = ""
    except Exception as exc:
        return {"text": "", "provider": provider, "status": "error", "note": str(exc)}

    translated = _clean(translated)
    if translated:
        cache[cache_key] = translated
        _save_cache(cache)
        return {"text": translated, "provider": provider, "status": "translated", "note": ""}

    return {"text": "", "provider": provider, "status": "empty_result", "note": "Provider nije vratio prijevod."}
