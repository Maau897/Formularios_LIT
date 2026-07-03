from __future__ import annotations

from dataclasses import dataclass
import base64
import json
import time
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class MicrosoftGraphError(RuntimeError):
    pass


@dataclass(frozen=True)
class MicrosoftGraphConfig:
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    refresh_token: str = ""
    shared_url: str = ""
    table_name: str = "ListaMaestra"


_CONFIG = MicrosoftGraphConfig()
_TOKEN_CACHE: dict[str, Any] = {}
_DRIVE_ITEM_CACHE: dict[str, Any] = {}


def configure_microsoft_graph(
    *,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    refresh_token: str = "",
    shared_url: str,
    table_name: str = "ListaMaestra",
) -> None:
    global _CONFIG, _TOKEN_CACHE, _DRIVE_ITEM_CACHE
    next_config = MicrosoftGraphConfig(
        tenant_id=tenant_id.strip(),
        client_id=client_id.strip(),
        client_secret=client_secret.strip(),
        refresh_token=refresh_token.strip(),
        shared_url=shared_url.strip(),
        table_name=(table_name or "ListaMaestra").strip() or "ListaMaestra",
    )
    if next_config != _CONFIG:
        _TOKEN_CACHE = {}
        _DRIVE_ITEM_CACHE = {}
    _CONFIG = next_config


def microsoft_graph_enabled() -> bool:
    return all(
        [
            _CONFIG.tenant_id,
            _CONFIG.client_id,
            _CONFIG.client_secret,
            _CONFIG.shared_url,
            _CONFIG.table_name,
        ]
    )


def _json_request(method: str, url: str, *, token: str = "", body: Any = None) -> dict[str, Any]:
    headers = {"Accept": "application/json"}
    data = None
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if body is not None:
        headers["Content-Type"] = "application/json"
        data = json.dumps(body).encode("utf-8")

    request = Request(url, data=data, headers=headers, method=method)
    try:
        with urlopen(request, timeout=30) as response:
            payload = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise MicrosoftGraphError(f"Microsoft Graph respondio {exc.code}: {detail}") from exc
    except URLError as exc:
        raise MicrosoftGraphError(f"No se pudo conectar con Microsoft Graph: {exc.reason}") from exc

    if not payload:
        return {}
    try:
        return json.loads(payload)
    except json.JSONDecodeError as exc:
        raise MicrosoftGraphError("Microsoft Graph regreso una respuesta no valida.") from exc


def _authenticated_json_request(method: str, url: str, *, body: Any = None) -> dict[str, Any]:
    token = _get_access_token()
    try:
        return _json_request(method, url, token=token, body=body)
    except MicrosoftGraphError as exc:
        if "Microsoft Graph respondio 401" not in str(exc):
            raise
        _TOKEN_CACHE.pop("access_token", None)
        _TOKEN_CACHE.pop("expires_at", None)
        return _json_request(method, url, token=_get_access_token(), body=body)


def _form_request(url: str, form: dict[str, str]) -> dict[str, Any]:
    request = Request(
        url,
        data=urlencode(form).encode("utf-8"),
        headers={
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=30) as response:
            payload = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise MicrosoftGraphError(f"Microsoft Identity respondio {exc.code}: {detail}") from exc
    except URLError as exc:
        raise MicrosoftGraphError(f"No se pudo conectar con Microsoft Identity: {exc.reason}") from exc
    return json.loads(payload)


def _get_access_token() -> str:
    if not microsoft_graph_enabled():
        raise MicrosoftGraphError("La integracion de Microsoft Graph no esta configurada.")
    cached_token = str(_TOKEN_CACHE.get("access_token", ""))
    expires_at = float(_TOKEN_CACHE.get("expires_at", 0))
    if cached_token and expires_at > time.time() + 60:
        return cached_token

    token_url = f"https://login.microsoftonline.com/{quote(_CONFIG.tenant_id)}/oauth2/v2.0/token"
    if _CONFIG.refresh_token:
        token_response = _form_request(
            token_url,
            {
                "client_id": _CONFIG.client_id,
                "client_secret": _CONFIG.client_secret,
                "refresh_token": _CONFIG.refresh_token,
                "scope": "https://graph.microsoft.com/Files.ReadWrite.All https://graph.microsoft.com/Sites.ReadWrite.All offline_access",
                "grant_type": "refresh_token",
            },
        )
    else:
        token_response = _form_request(
            token_url,
            {
                "client_id": _CONFIG.client_id,
                "client_secret": _CONFIG.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
        )
    access_token = str(token_response.get("access_token", ""))
    if not access_token:
        raise MicrosoftGraphError("Microsoft Identity no regreso access_token.")
    expires_in = int(token_response.get("expires_in", 3600) or 3600)
    _TOKEN_CACHE["access_token"] = access_token
    _TOKEN_CACHE["expires_at"] = time.time() + max(expires_in, 60)
    if token_response.get("refresh_token"):
        _TOKEN_CACHE["refresh_token"] = token_response["refresh_token"]
    return access_token


def _sharing_token(shared_url: str) -> str:
    encoded = base64.urlsafe_b64encode(shared_url.encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{encoded}"


def _resolve_drive_item() -> dict[str, str]:
    if _DRIVE_ITEM_CACHE:
        return dict(_DRIVE_ITEM_CACHE)

    share_id = quote(_sharing_token(_CONFIG.shared_url), safe="!")
    response = _authenticated_json_request("GET", f"{GRAPH_BASE_URL}/shares/{share_id}/driveItem")
    parent_reference = response.get("parentReference") or {}
    drive_id = str(parent_reference.get("driveId", ""))
    item_id = str(response.get("id", ""))
    if not drive_id or not item_id:
        raise MicrosoftGraphError("No se pudo resolver el archivo de OneDrive desde el link compartido.")

    _DRIVE_ITEM_CACHE.update(
        {
            "drive_id": drive_id,
            "item_id": item_id,
            "web_url": str(response.get("webUrl", _CONFIG.shared_url)),
        }
    )
    return dict(_DRIVE_ITEM_CACHE)


def _table_base_url(drive_id: str, item_id: str) -> str:
    return (
        f"{GRAPH_BASE_URL}/drives/{quote(drive_id)}/items/{quote(item_id)}"
        f"/workbook/tables/{quote(_CONFIG.table_name, safe='')}"
    )


def get_master_table() -> dict[str, Any]:
    item = _resolve_drive_item()
    table_url = _table_base_url(item["drive_id"], item["item_id"])
    columns_response = _authenticated_json_request("GET", f"{table_url}/columns")
    rows_response = _authenticated_json_request("GET", f"{table_url}/rows")

    columns = [str(column.get("name", "")) for column in columns_response.get("value", [])]
    rows = []
    for row in rows_response.get("value", []):
        values = row.get("values") or [[]]
        rows.append(
            {
                "index": int(row.get("index", len(rows))),
                "values": list(values[0] if values else []),
            }
        )

    return {
        "columns": columns,
        "rows": rows,
        "drive_id": item["drive_id"],
        "item_id": item["item_id"],
        "web_url": item.get("web_url", _CONFIG.shared_url),
        "table_name": _CONFIG.table_name,
    }


def _split_excel_address(address: str) -> tuple[str, str]:
    text = address.strip()
    if not text:
        raise MicrosoftGraphError("La fila de Excel no regreso una direccion valida.")
    if text.startswith("'"):
        end = text.find("'!")
        if end < 0:
            raise MicrosoftGraphError(f"No se pudo interpretar la direccion de Excel: {address}")
        sheet_name = text[1:end].replace("''", "'")
        cell_range = text[end + 2 :]
        return sheet_name, cell_range
    if "!" not in text:
        raise MicrosoftGraphError(f"No se pudo interpretar la direccion de Excel: {address}")
    sheet_name, cell_range = text.split("!", 1)
    return sheet_name.strip("'"), cell_range


def _odata_string_literal(value: str) -> str:
    return quote(value.replace("'", "''"), safe="$:")


def update_master_table_row(row_index: int, values: list[Any]) -> None:
    item = _resolve_drive_item()
    table_url = _table_base_url(item["drive_id"], item["item_id"])
    row_range_url = f"{table_url}/rows/itemAt(index={int(row_index)})/range"
    range_response = _authenticated_json_request("GET", row_range_url)
    address = str(range_response.get("address", ""))
    sheet_name, cell_range = _split_excel_address(address)
    update_url = (
        f"{GRAPH_BASE_URL}/drives/{quote(item['drive_id'])}/items/{quote(item['item_id'])}"
        f"/workbook/worksheets('{_odata_string_literal(sheet_name)}')"
        f"/range(address='{_odata_string_literal(cell_range)}')"
    )
    _authenticated_json_request("PATCH", update_url, body={"values": [values]})


def add_master_table_row(values: list[Any]) -> None:
    item = _resolve_drive_item()
    table_url = _table_base_url(item["drive_id"], item["item_id"])
    _authenticated_json_request("POST", f"{table_url}/rows/add", body={"values": [values]})
