from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
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
) -> None:
    global _CONFIG, _CLIENT
    _CONFIG = SupabaseStorageConfig(
        url=(url or "").strip(),
        key=(key or "").strip(),
        enabled=bool(enabled and url and key),
        table_name=table_name,
        signatures_bucket=(signatures_bucket or "").strip(),
        signatures_prefix=(signatures_prefix or "").strip().strip("/"),
    )
    _CLIENT = None


def supabase_storage_enabled() -> bool:
    return bool(SUPABASE_SDK_AVAILABLE and _CONFIG.enabled and _CONFIG.url and _CONFIG.key)


def get_storage_backend_label() -> str:
    return "Supabase" if supabase_storage_enabled() else "Local JSON"


def signatures_storage_enabled() -> bool:
    return bool(supabase_storage_enabled() and _CONFIG.signatures_bucket)


def get_signatures_storage_cache_key() -> tuple[bool, str, str]:
    return (
        signatures_storage_enabled(),
        _CONFIG.signatures_bucket,
        _CONFIG.signatures_prefix,
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
        asset_path = f"{prefix}/{name}" if prefix else name
        assets.append({"name": name, "path": asset_path})
    return assets


def download_signature_bytes(asset_path: str) -> bytes:
    if not asset_path.strip():
        raise ValueError("La ruta del archivo de firma es obligatoria.")
    return _storage_bucket().download(asset_path.strip())
