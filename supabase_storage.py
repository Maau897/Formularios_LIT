from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timezone
from typing import Any

try:
    from supabase import Client, create_client
    SUPABASE_SDK_AVAILABLE = True
except ImportError:  # pragma: no cover
    Client = object  # type: ignore[assignment]
    create_client = None
    SUPABASE_SDK_AVAILABLE = False


@dataclass
class SupabaseStorageConfig:
    url: str
    key: str
    enabled: bool = False
    table_name: str = "formularios_periodos"
    signatures_bucket: str = "firmas-digitales"
    signatures_prefix: str = ""
    templates_bucket: str = ""
    templates_prefix: str = ""
    traceability_table_name: str = "formularios_trazabilidad"


_CONFIG = SupabaseStorageConfig(url="", key="", enabled=False)
_CLIENT: Client | None = None


def configure_supabase_storage(
    *,
    url: str | None,
    key: str | None,
    enabled: bool = False,
    table_name: str = "formularios_periodos",
    signatures_bucket: str = "firmas-digitales",
    signatures_prefix: str = "",
    templates_bucket: str = "",
    templates_prefix: str = "",
    traceability_table_name: str = "formularios_trazabilidad",
) -> None:
    global _CONFIG, _CLIENT
    _CONFIG = SupabaseStorageConfig(
        url=(url or "").strip(),
        key=(key or "").strip(),
        enabled=bool(enabled and url and key),
        table_name=table_name,
        signatures_bucket=(signatures_bucket or "").strip(),
        signatures_prefix=(signatures_prefix or "").strip().strip("/"),
        templates_bucket=(templates_bucket or "").strip(),
        templates_prefix=(templates_prefix or "").strip().strip("/"),
        traceability_table_name=(traceability_table_name or "").strip() or "formularios_trazabilidad",
    )
    _CLIENT = None


def supabase_storage_enabled() -> bool:
    return bool(SUPABASE_SDK_AVAILABLE and _CONFIG.enabled and _CONFIG.url and _CONFIG.key)


def get_storage_backend_label() -> str:
    return "Supabase" if supabase_storage_enabled() else "Local JSON"


def signatures_storage_enabled() -> bool:
    return bool(supabase_storage_enabled() and _CONFIG.signatures_bucket)


def templates_storage_enabled() -> bool:
    return bool(supabase_storage_enabled() and _CONFIG.templates_bucket)


def get_signatures_storage_cache_key() -> tuple[bool, str, str]:
    return (
        signatures_storage_enabled(),
        _CONFIG.signatures_bucket,
        _CONFIG.signatures_prefix,
    )


def get_templates_storage_cache_key() -> tuple[bool, str, str]:
    return (
        templates_storage_enabled(),
        _CONFIG.templates_bucket,
        _CONFIG.templates_prefix,
    )


def _client() -> Client:
    global _CLIENT
    if not supabase_storage_enabled():
        raise RuntimeError("Supabase de almacenamiento no esta configurado.")
    if _CLIENT is None:
        _CLIENT = create_client(_CONFIG.url, _CONFIG.key)
    return _CLIENT


def _table():
    return _client().table(_CONFIG.table_name)


def _storage_bucket():
    if not signatures_storage_enabled():
        raise RuntimeError("Supabase Storage de firmas no esta configurado.")
    return _client().storage.from_(_CONFIG.signatures_bucket)


def _templates_storage_bucket():
    if not templates_storage_enabled():
        raise RuntimeError("Supabase Storage de plantillas no esta configurado.")
    return _client().storage.from_(_CONFIG.templates_bucket)


def _traceability_table():
    return _client().table(_CONFIG.traceability_table_name)


def _build_storage_asset_path(prefix: str, name: str) -> str:
    normalized_name = name.strip().lstrip("/")
    normalized_prefix = prefix.strip().strip("/")
    if not normalized_name:
        raise ValueError("El nombre del archivo es obligatorio.")
    return f"{normalized_prefix}/{normalized_name}" if normalized_prefix else normalized_name


def save_period_payload(payload: dict[str, Any], updated_by: str = "") -> None:
    metadata = payload.get("metadata", {})
    now_iso = datetime.now(timezone.utc).isoformat()
    normalized_user = updated_by.strip().lower() or None
    row = {
        "form_key": str(metadata.get("form_key", "")),
        "equipment_code": str(metadata.get("equipment_code", "")),
        "month": int(metadata.get("month", 0)),
        "year": int(metadata.get("year", 0)),
        "payload": payload,
        "created_by": normalized_user,
        "updated_by": normalized_user,
        "updated_at": now_iso,
    }
    if row["form_key"] == "" or row["equipment_code"] == "" or row["month"] == 0 or row["year"] == 0:
        raise ValueError("El payload no tiene un periodo valido para guardarse en Supabase.")

    _table().upsert(row, on_conflict="form_key,equipment_code,year,month").execute()


def load_period_payload(form_key: str, equipment_code: str, year: int, month: int) -> dict[str, Any] | None:
    rows = (
        _table()
        .select("payload")
        .eq("form_key", form_key)
        .eq("equipment_code", equipment_code)
        .eq("year", year)
        .eq("month", month)
        .limit(1)
        .execute()
        .data
        or []
    )
    if not rows:
        return None
    payload = rows[0].get("payload")
    return payload if isinstance(payload, dict) else None


def list_periods() -> list[dict[str, Any]]:
    rows = (
        _table()
        .select("form_key,equipment_code,month,year,updated_at,updated_by")
        .order("year", desc=True)
        .order("month", desc=True)
        .order("equipment_code")
        .execute()
        .data
        or []
    )
    periods: list[dict[str, Any]] = []
    for row in rows:
        try:
            periods.append(
                {
                    "form_key": str(row["form_key"]),
                    "equipment_code": str(row["equipment_code"]),
                    "month": int(row["month"]),
                    "year": int(row["year"]),
                    "updated_at": row.get("updated_at", ""),
                    "updated_by": row.get("updated_by", "") or "",
                }
            )
        except (KeyError, TypeError, ValueError):
            continue
    return periods


def list_signature_assets() -> list[dict[str, str]]:
    if not signatures_storage_enabled():
        return []

    prefix = _CONFIG.signatures_prefix
    try:
        response = _storage_bucket().list(prefix or "")
    except TypeError:
        response = _storage_bucket().list(path=prefix or "")

    assets: list[dict[str, str]] = []
    for entry in response or []:
        name = str(entry.get("name", "")).strip()
        if not name or not name.lower().endswith(".png"):
            continue
        asset_path = _build_storage_asset_path(prefix, name)
        assets.append({"name": name, "path": asset_path})
    return assets


def download_signature_bytes(asset_path: str) -> bytes:
    if not asset_path.strip():
        raise ValueError("La ruta del archivo de firma es obligatoria.")
    return _storage_bucket().download(asset_path.strip())


def download_template_bytes(file_name: str) -> bytes:
    asset_path = _build_storage_asset_path(_CONFIG.templates_prefix, file_name)
    return _templates_storage_bucket().download(asset_path)


def list_traceability_entries(form_key: str, equipment_code: str) -> list[dict[str, Any]]:
    rows = (
        _traceability_table()
        .select("*")
        .eq("form_key", form_key)
        .eq("equipment_code", equipment_code)
        .order("scheduled_for")
        .order("created_at", desc=True)
        .execute()
        .data
        or []
    )
    entries: list[dict[str, Any]] = []
    for row in rows:
        try:
            entries.append(
                {
                    "id": str(row.get("id", "")).strip(),
                    "form_key": str(row.get("form_key", "")).strip(),
                    "equipment_code": str(row.get("equipment_code", "")).strip(),
                    "entry_type": str(row.get("entry_type", "")).strip(),
                    "status": str(row.get("status", "")).strip() or "programado",
                    "scheduled_for": str(row.get("scheduled_for", "")).strip(),
                    "completed_on": str(row.get("completed_on", "")).strip(),
                    "provider": str(row.get("provider", "")).strip(),
                    "notes": str(row.get("notes", "")).strip(),
                    "created_by": str(row.get("created_by", "")).strip(),
                    "updated_by": str(row.get("updated_by", "")).strip(),
                    "updated_at": str(row.get("updated_at", "")).strip(),
                }
            )
        except Exception:
            continue
    return entries


def save_traceability_entry(entry: dict[str, Any], updated_by: str = "") -> dict[str, Any]:
    normalized_user = updated_by.strip().lower() or None
    now_iso = datetime.now(timezone.utc).isoformat()
    normalized_entry = {
        "id": str(entry.get("id", "")).strip() or None,
        "form_key": str(entry.get("form_key", "")).strip(),
        "equipment_code": str(entry.get("equipment_code", "")).strip(),
        "entry_type": str(entry.get("entry_type", "")).strip(),
        "status": str(entry.get("status", "")).strip() or "programado",
        "scheduled_for": str(entry.get("scheduled_for", "")).strip() or None,
        "completed_on": str(entry.get("completed_on", "")).strip() or None,
        "provider": str(entry.get("provider", "")).strip() or None,
        "notes": str(entry.get("notes", "")).strip() or None,
        "updated_by": normalized_user,
        "updated_at": now_iso,
    }
    if not normalized_entry["form_key"] or not normalized_entry["equipment_code"] or not normalized_entry["entry_type"]:
        raise ValueError("La trazabilidad requiere formato, equipo y tipo de evento.")
    if normalized_entry["id"] is None:
        normalized_entry["created_by"] = normalized_user
    response = _traceability_table().upsert(normalized_entry).execute().data or []
    saved = response[0] if response else normalized_entry
    return {
        "id": str(saved.get("id", normalized_entry["id"] or "")).strip(),
        "form_key": str(saved.get("form_key", normalized_entry["form_key"])).strip(),
        "equipment_code": str(saved.get("equipment_code", normalized_entry["equipment_code"])).strip(),
        "entry_type": str(saved.get("entry_type", normalized_entry["entry_type"])).strip(),
        "status": str(saved.get("status", normalized_entry["status"])).strip(),
        "scheduled_for": str(saved.get("scheduled_for", normalized_entry.get("scheduled_for") or "")).strip(),
        "completed_on": str(saved.get("completed_on", normalized_entry.get("completed_on") or "")).strip(),
        "provider": str(saved.get("provider", normalized_entry.get("provider") or "")).strip(),
        "notes": str(saved.get("notes", normalized_entry.get("notes") or "")).strip(),
        "created_by": str(saved.get("created_by", normalized_user or "")).strip(),
        "updated_by": str(saved.get("updated_by", normalized_user or "")).strip(),
        "updated_at": str(saved.get("updated_at", now_iso)).strip(),
    }


def delete_traceability_entry(entry_id: str) -> None:
    normalized_id = entry_id.strip()
    if not normalized_id:
        raise ValueError("El identificador de la trazabilidad es obligatorio.")
    _traceability_table().delete().eq("id", normalized_id).execute()
