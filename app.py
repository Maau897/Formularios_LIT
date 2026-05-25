from __future__ import annotations

from calendar import monthrange
from dataclasses import asdict, dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import json
import os
import re
import shutil
import time
from typing import Any
import unicodedata
from zoneinfo import ZoneInfo

import streamlit as st
try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv() -> bool:
        return False
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage
from supabase_users import (
    aprobar_usuario,
    actualizar_rol_usuario,
    autenticar_usuario,
    configure_supabase_users,
    crear_admin_inicial,
    eliminar_usuario,
    listar_eventos_auditoria,
    listar_usuarios,
    registrar_evento_auditoria,
    obtener_usuarios_pendientes,
    registrar_usuario,
    supabase_users_enabled,
)
from supabase_storage import (
    configure_supabase_storage,
    download_template_bytes,
    download_signature_bytes,
    get_signatures_storage_cache_key,
    get_templates_storage_cache_key,
    list_signature_assets as list_remote_signature_assets,
    load_period_payload as load_remote_period_payload,
    save_period_payload as save_remote_period_payload,
    signatures_storage_enabled,
    supabase_storage_enabled,
    templates_storage_enabled,
)


load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "F-LIT-21-03.xlsx"
WORKING_TEMPLATE_PATH = BASE_DIR / "template_cong1.xlsx"
DATA_DIR = BASE_DIR / "data"
SIGNATURES_DIR = BASE_DIR / "firmas digitales"

MONTHS = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

TIME_SLOTS = [
    "7:00 - 10:00",
    "11:00 - 14:00",
    "15:00 - 18:00",
]

DAY_BLOCK_START_COLUMNS = {
    day: 5 + ((day - 1) * 3)
    for day in range(1, 17)
}
DAY_BLOCK_START_COLUMNS.update(
    {day: 5 + ((day - 17) * 3) for day in range(17, 32)}
)

ROW_GROUPS = {
    "top": {"temperature": 16, "performed_by": 19, "verified_by": 20, "date": 21},
    "bottom": {"temperature": 28, "performed_by": 31, "verified_by": 32, "date": 33},
}

CORRECTION_CELL_MAP = {
    "range_1": "AU7",
    "range_2": "AW7",
    "range_3": "AY7",
}

HEADER_CELL_MAP = {
    "month": "P6",
    "year": "X6",
}

FOOTER_CELL_MAP = {
    "observations": "H34",
    "reviewed_by": "H35",
    "reviewed_on": "AI35",
}

FORM_DEFINITIONS: dict[str, dict[str, Any]] = {
    "congeladores": {
        "label": "F-LIT-21-03 Congeladores",
        "source_file": "F-LIT-21-03.xlsx",
        "working_file": "template_cong1.xlsx",
        "sheet_exclusions": {"CONG"},
        "default_equipment": "CONG-1",
        "supports_corrections": True,
        "metrics": [
            {"key": "measured_temperatures", "label": "Temperatura medida", "unit": "°C", "corrected": True},
        ],
        "layout": {
            "top": {"metric_1": 16, "performed_by": 19, "verified_by": 20, "date": 21},
            "bottom": {"metric_1": 28, "performed_by": 31, "verified_by": 32, "date": 33},
            "footer": {"observations": "H34", "reviewed_by": "H35", "reviewed_on": "AI35"},
            "status_rows_to_merge": [20, 21, 32, 33],
        },
        "extractor": "cold_equipment",
    },
    "ultracongeladores": {
        "label": "F-LIT-20-03 Ultracongeladores",
        "source_file": "F-LIT-20-03 ulcos.xlsx",
        "working_file": "template_ulcos.xlsx",
        "sheet_exclusions": {"ULCO"},
        "default_equipment": "ULCO-1",
        "supports_corrections": True,
        "metrics": [
            {"key": "measured_temperatures", "label": "Temperatura medida", "unit": "°C", "corrected": True},
        ],
        "layout": {
            "top": {"metric_1": 16, "performed_by": 19, "verified_by": 20, "date": 21},
            "bottom": {"metric_1": 28, "performed_by": 31, "verified_by": 32, "date": 33},
            "footer": {"observations": "H34", "reviewed_by": "H35", "reviewed_on": "AI35"},
            "status_rows_to_merge": [20, 21, 32, 33],
        },
        "extractor": "cold_equipment",
    },
    "refrigeradores": {
        "label": "F-LIT-22-03 Refrigeradores",
        "source_file": "F-LIT-22-03 regrigeradores.xlsx",
        "working_file": "template_refrigeradores.xlsx",
        "sheet_exclusions": {"REFR"},
        "default_equipment": "REFR-1",
        "supports_corrections": True,
        "metrics": [
            {"key": "measured_temperatures", "label": "Temperatura medida", "unit": "°C", "corrected": True},
        ],
        "layout": {
            "top": {"metric_1": 16, "performed_by": 19, "verified_by": 20, "date": 21},
            "bottom": {"metric_1": 28, "performed_by": 31, "verified_by": 32, "date": 33},
            "footer": {"observations": "H34", "reviewed_by": "H35", "reviewed_on": "AI35"},
            "status_rows_to_merge": [20, 21, 32, 33],
        },
        "extractor": "cold_equipment",
    },
    "incubadoras": {
        "label": "F-LIT-23-03 Incubadoras",
        "source_file": "F-LIT-23-03 incubadoras.xlsx",
        "working_file": "template_incubadoras.xlsx",
        "sheet_exclusions": {"ICO2"},
        "default_equipment": "ICO2-1",
        "supports_corrections": True,
        "metrics": [
            {"key": "measured_temperatures", "label": "Temperatura medida", "unit": "°C", "corrected": True},
            {"key": "secondary_measurements", "label": "%CO2", "unit": "", "corrected": False},
        ],
        "layout": {
            "top": {"metric_1": 16, "metric_2": 19, "performed_by": 22, "verified_by": 23, "date": 24},
            "bottom": {"metric_1": 31, "metric_2": 34, "performed_by": 37, "verified_by": 38, "date": 39},
            "footer": {"observations": "K40", "reviewed_by": "K41", "reviewed_on": "AL41"},
            "status_rows_to_merge": [23, 24, 38, 39],
        },
        "extractor": "incubators",
    },
    "condiciones_ambientales": {
        "label": "F-LIT-09-04 Condiciones Ambientales",
        "source_file": "F-LIT-09-04 condiciones ambientales.xlsx",
        "working_file": "template_condiciones_ambientales.xlsx",
        "sheet_exclusions": {"TEMPERATURA", "HUMEDAD"},
        "default_equipment": "TEMPERATURA 504",
        "supports_corrections": True,
        "metrics": [
            {"key": "measured_temperatures", "label": "Lectura medida", "unit": "", "corrected": True},
        ],
        "layout": {
            "top": {"metric_1": 17, "performed_by": 20, "verified_by": 21, "date": 22},
            "bottom": {"metric_1": 29, "performed_by": 32, "verified_by": 33, "date": 34},
            "footer": {"observations": "E35", "reviewed_by": "E36", "reviewed_on": "AE36"},
            "status_rows_to_merge": [21, 22, 33, 34],
        },
        "extractor": "ambient",
    },
}

DEFAULT_FORM_KEY = "congeladores"
DEFAULT_EQUIPMENT_CODE = FORM_DEFINITIONS[DEFAULT_FORM_KEY]["default_equipment"]
ROLES_USUARIO = ["captura", "responsable", "auditor", "calidad", "admin"]
SENSITIVE_EDITOR_ROLES = {"calidad", "admin"}
AUTOSAVE_DEBOUNCE_SECONDS = 3.0
LOCAL_TIMEZONE = ZoneInfo("America/Mexico_City")


@dataclass
class DailyCapture:
    active: bool
    measured_temperatures: list[str]
    corrected_temperatures: list[str]
    secondary_measurements: list[str]
    performed_by_slots: list[str]
    verified_by: str
    recorded_on: str
    notes: str = ""


def default_daily_capture(day: int) -> DailyCapture:
    is_weekday = date.today().replace(day=min(day, 28)).weekday() < 5
    return DailyCapture(
        active=is_weekday,
        measured_temperatures=["", "", ""],
        corrected_temperatures=["", "", ""],
        secondary_measurements=["", "", ""],
        performed_by_slots=["", "", ""],
        verified_by="",
        recorded_on="",
        notes="",
    )


def get_config_value(secret_key: str, env_key: str, default: Any = "") -> Any:
    try:
        return st.secrets.get(secret_key, os.getenv(env_key, default))
    except Exception:
        return os.getenv(env_key, default)


def as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "si", "on"}
    return bool(value)


def configure_users_backend() -> None:
    configure_supabase_users(
        url=str(get_config_value("supabase_url", "SUPABASE_URL", "")),
        key=str(get_config_value("supabase_key", "SUPABASE_KEY", "")),
        enabled=as_bool(get_config_value("use_supabase_users", "USE_SUPABASE_USERS", False)),
        table_name=str(get_config_value("supabase_users_table", "SUPABASE_USERS_TABLE", "usuarios_app")),
        audit_table_name=str(get_config_value("supabase_audit_table", "SUPABASE_AUDIT_TABLE", "formularios_auditoria")),
    )

    admin_email = str(get_config_value("admin_email", "ADMIN_EMAIL", "")).strip()
    admin_password = str(get_config_value("admin_password", "ADMIN_PASSWORD", "")).strip()
    if supabase_users_enabled() and admin_email and admin_password:
        try:
            crear_admin_inicial(admin_email, admin_password)
        except Exception:
            pass


def configure_storage_backend() -> None:
    configure_supabase_storage(
        url=str(get_config_value("supabase_url", "SUPABASE_URL", "")),
        key=str(get_config_value("supabase_key", "SUPABASE_KEY", "")),
        enabled=as_bool(get_config_value("use_supabase_storage", "USE_SUPABASE_STORAGE", True)),
        table_name=str(get_config_value("supabase_storage_table", "SUPABASE_STORAGE_TABLE", "formularios_periodos")),
        signatures_bucket=str(get_config_value("supabase_signatures_bucket", "SUPABASE_SIGNATURES_BUCKET", "firmas-digitales")),
        signatures_prefix=str(get_config_value("supabase_signatures_prefix", "SUPABASE_SIGNATURES_PREFIX", "")),
        templates_bucket=str(get_config_value("supabase_templates_bucket", "SUPABASE_TEMPLATES_BUCKET", "")),
        templates_prefix=str(get_config_value("supabase_templates_prefix", "SUPABASE_TEMPLATES_PREFIX", "")),
    )


def normalize_user_role(rol: Any, es_admin: bool) -> str:
    if es_admin:
        return "admin"
    normalized = str(rol or "captura").strip().lower()
    return normalized if normalized in ROLES_USUARIO else "captura"


def initialize_auth_state() -> None:
    defaults = {
        "autenticado": False,
        "usuario_email": "",
        "es_admin": False,
        "rol_usuario": "captura",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def current_user_role() -> str:
    return str(st.session_state.get("rol_usuario", "captura"))


def log_activity(accion: str, detalle: str = "", payload: dict[str, Any] | None = None) -> None:
    if not supabase_users_enabled():
        return

    form_key = ""
    equipment_code = ""
    month: int | None = None
    year: int | None = None
    target_payload = payload or st.session_state.get("payload")
    if isinstance(target_payload, dict):
        metadata = target_payload.get("metadata", {})
        form_key = str(metadata.get("form_key", "")).strip()
        equipment_code = str(metadata.get("equipment_code", "")).strip()
        try:
            month = int(metadata.get("month"))
            year = int(metadata.get("year"))
        except (TypeError, ValueError):
            month = None
            year = None

    try:
        registrar_evento_auditoria(
            email=str(st.session_state.get("usuario_email", "")).strip(),
            accion=accion,
            detalle=detalle,
            form_key=form_key,
            equipment_code=equipment_code,
            month=month,
            year=year,
        )
    except Exception:
        pass


def get_local_now() -> datetime:
    return datetime.now(LOCAL_TIMEZONE)


def format_local_timestamp(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=LOCAL_TIMEZONE)
        return parsed.astimezone(LOCAL_TIMEZONE).strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return text


def can_edit_sensitive_configuration() -> bool:
    return current_user_role() in SENSITIVE_EDITOR_ROLES


def can_edit_schedule() -> bool:
    return current_user_role() in {"responsable", "calidad", "admin"}


def can_edit_daily_records() -> bool:
    return current_user_role() in {"captura", "responsable", "calidad", "admin"}


def can_close_period() -> bool:
    return current_user_role() in {"responsable", "calidad", "admin"}


def can_export_period() -> bool:
    return current_user_role() in {"responsable", "calidad", "admin"}


def render_auth_screen() -> None:
    st.title("Acceso al sistema")

    if not supabase_users_enabled():
        st.error("La autenticacion no esta configurada en esta app.")
        st.stop()

    login_tab, register_tab = st.tabs(["Iniciar sesion", "Crear cuenta"])

    with login_tab:
        email_login = st.text_input("Correo", key="login_email")
        password_login = st.text_input("Contrasena", type="password", key="login_password")
        if st.button("Ingresar", use_container_width=True):
            try:
                result = autenticar_usuario(email_login, password_login, normalize_user_role)
                if result["ok"]:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario_email"] = result["email"]
                    st.session_state["es_admin"] = result["es_admin"]
                    st.session_state["rol_usuario"] = result.get("rol", "captura")
                    log_activity("inicio_sesion", "Ingreso a la app")
                    st.rerun()
                else:
                    st.error(result["mensaje"])
            except Exception as exc:
                st.error(f"No se pudo iniciar sesion: {exc}")

    with register_tab:
        email_register = st.text_input("Correo institucional o personal", key="register_email")
        password_register = st.text_input("Contrasena", type="password", key="register_password")
        password_register_2 = st.text_input("Confirmar contrasena", type="password", key="register_password_2")
        requested_role = st.selectbox(
            "Perfil solicitado",
            ["captura", "responsable", "auditor", "calidad"],
            format_func=lambda value: value.capitalize(),
            key="register_role",
        )
        if st.button("Crear cuenta", use_container_width=True):
            try:
                if not email_register or not password_register:
                    st.warning("Completa correo y contrasena.")
                elif password_register != password_register_2:
                    st.warning("Las contrasenas no coinciden.")
                else:
                    registrar_usuario(email_register, password_register, requested_role)
                    st.success("Cuenta creada. Queda pendiente de aprobacion.")
            except Exception as exc:
                st.error(f"No se pudo crear la cuenta: {exc}")


def render_user_admin_sidebar() -> None:
    if not st.session_state.get("es_admin", False) or not supabase_users_enabled():
        return

    st.sidebar.divider()
    st.sidebar.subheader("Administracion")
    try:
        pending_users = obtener_usuarios_pendientes()
        st.sidebar.markdown("**Solicitudes pendientes**")
        if pending_users:
            for user_id, email, registered_at in pending_users:
                st.sidebar.write(f"{email} - {registered_at}")
                approval_role = st.sidebar.selectbox(
                    f"Rol para {email}",
                    ROLES_USUARIO,
                    index=ROLES_USUARIO.index("captura"),
                    format_func=lambda value: value.capitalize(),
                    key=f"approval_role_{user_id}",
                )
                if st.sidebar.button("Aprobar", key=f"approve_{user_id}", use_container_width=True):
                    aprobar_usuario(user_id, approval_role)
                    log_activity("aprobar_usuario", f"{email} -> {approval_role}")
                    st.sidebar.success(f"Usuario {email} aprobado.")
                    st.rerun()
        else:
            st.sidebar.caption("No hay usuarios pendientes.")

        st.sidebar.markdown("**Roles activos**")
        approved_users = [row for row in listar_usuarios() if row[2] == 1]
        admin_count = sum(1 for row in approved_users if row[3] == 1)
        if approved_users:
            for user_id, email, _, _, role, _ in approved_users:
                new_role = st.sidebar.selectbox(
                    email,
                    ROLES_USUARIO,
                    index=ROLES_USUARIO.index(role if role in ROLES_USUARIO else "captura"),
                    format_func=lambda value: value.capitalize(),
                    key=f"role_user_{user_id}",
                )
                if st.sidebar.button("Actualizar rol", key=f"update_role_{user_id}", use_container_width=True):
                    actualizar_rol_usuario(user_id, new_role)
                    log_activity("actualizar_rol", f"{email} -> {new_role}")
                    st.sidebar.success(f"Rol de {email} actualizado a {new_role}.")
                    st.rerun()
                can_delete_user = email != str(st.session_state.get("usuario_email", "")).strip().lower()
                would_remove_last_admin = role == "admin" and admin_count <= 1
                if st.sidebar.button("Quitar acceso", key=f"delete_user_{user_id}", use_container_width=True):
                    if not can_delete_user:
                        st.sidebar.warning("No puedes quitar tu propio acceso desde aqui.")
                    elif would_remove_last_admin:
                        st.sidebar.warning("No puedes quitar al ultimo admin activo.")
                    else:
                        eliminar_usuario(user_id)
                        log_activity("quitar_acceso", email)
                        st.sidebar.success(f"Se quito el acceso de {email}.")
                        st.rerun()
        else:
            st.sidebar.caption("No hay usuarios aprobados para administrar.")

        with st.sidebar.expander("Historial de actividad", expanded=False):
            try:
                eventos = listar_eventos_auditoria(limit=40)
                if eventos:
                    for evento in eventos:
                        marca_tiempo = format_local_timestamp(str(evento.get("created_at", "")))
                        accion = str(evento.get("accion", "")).replace("_", " ").capitalize()
                        email = str(evento.get("email", ""))
                        detalle = str(evento.get("detalle", "")).strip()
                        contexto = []
                        if evento.get("form_key"):
                            contexto.append(str(evento["form_key"]))
                        if evento.get("equipment_code"):
                            contexto.append(str(evento["equipment_code"]))
                        if evento.get("year") and evento.get("month"):
                            contexto.append(f"{int(evento['year'])}-{int(evento['month']):02d}")
                        context_label = " | ".join(contexto)
                        st.write(f"{marca_tiempo} - {email}")
                        st.caption(f"{accion}{' | ' + detalle if detalle else ''}{' | ' + context_label if context_label else ''}")
                else:
                    st.caption("Sin actividad registrada todavia.")
            except Exception:
                st.caption("El historial aun no esta disponible.")
    except Exception as exc:
        st.sidebar.error(f"No se pudo cargar la administracion de usuarios: {exc}")


def coerce_factor_value(raw_value: Any) -> tuple[str, float]:
    try:
        numeric_value = float(str(raw_value).replace(",", ".").strip())
    except (TypeError, ValueError, AttributeError):
        numeric_value = 0.0

    operation = "+" if numeric_value >= 0 else "-"
    return operation, abs(numeric_value)


def parse_range_bounds(label_text: str) -> tuple[float, float] | None:
    match = re.search(r"(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)", label_text)
    if not match:
        return None
    return float(match.group(1)), float(match.group(2))


def find_value_after_label(
    worksheet: Any,
    label: str,
    rows: list[int],
    max_columns: int | None = None,
) -> Any:
    max_col = max_columns or worksheet.max_column
    normalized_label = label.strip().lower()

    for row in rows:
        for column in range(1, max_col + 1):
            cell_value = worksheet.cell(row, column).value
            if cell_value is None:
                continue
            if str(cell_value).strip().lower() != normalized_label:
                continue

            for next_col in range(column + 1, max_col + 1):
                next_value = worksheet.cell(row, next_col).value
                if next_value not in (None, ""):
                    return next_value

    return None


def find_cell_after_label(
    worksheet: Any,
    label: str,
    rows: list[int],
    max_columns: int | None = None,
) -> str | None:
    max_col = max_columns or worksheet.max_column
    normalized_label = label.strip().lower()

    for row in rows:
        for column in range(1, max_col + 1):
            cell_value = worksheet.cell(row, column).value
            if cell_value is None:
                continue
            if str(cell_value).strip().lower() != normalized_label:
                continue

            for next_col in range(column + 1, max_col + 1):
                next_value = worksheet.cell(row, next_col).value
                if next_value not in (None, ""):
                    return worksheet.cell(row, next_col).coordinate

    return None


def get_form_definition(form_key: str) -> dict[str, Any]:
    return FORM_DEFINITIONS[form_key]


def get_template_paths(form_key: str) -> tuple[Path, Path]:
    definition = get_form_definition(form_key)
    return BASE_DIR / definition["source_file"], BASE_DIR / definition["working_file"]


def _read_local_template_bytes(source_path: Path, working_path: Path) -> bytes:
    try:
        return source_path.read_bytes()
    except PermissionError:
        if not working_path.exists() and source_path.exists():
            try:
                shutil.copy2(source_path, working_path)
            except PermissionError:
                workbook = load_workbook(source_path)
                workbook.save(working_path)
        return working_path.read_bytes()
    except FileNotFoundError:
        if working_path.exists():
            return working_path.read_bytes()
        raise


@st.cache_data(show_spinner=False, ttl=300)
def _get_template_bytes_cached(
    form_key: str,
    source_file_name: str,
    _templates_storage_cache_key: tuple[bool, str, str],
) -> bytes:
    source_path, working_path = get_template_paths(form_key)
    if templates_storage_enabled():
        try:
            return download_template_bytes(source_file_name)
        except Exception:
            pass
    return _read_local_template_bytes(source_path, working_path)


def get_template_bytes(form_key: str) -> bytes:
    definition = get_form_definition(form_key)
    return _get_template_bytes_cached(
        form_key,
        str(definition["source_file"]),
        get_templates_storage_cache_key(),
    )


def template_source_available(form_key: str) -> bool:
    source_path, working_path = get_template_paths(form_key)
    if templates_storage_enabled():
        return True
    return source_path.exists() or working_path.exists()


def extract_inline_corrections(
    worksheet: Any,
    range_row: int,
    factor_row: int,
) -> tuple[dict[str, dict[str, Any]], dict[str, str], dict[str, float], dict[str, str]]:
    correction_bands: dict[str, dict[str, Any]] = {}
    correction_cells: dict[str, str] = {}
    correction_operations: dict[str, str] = {}
    correction_factors: dict[str, float] = {}
    range_index = 1
    range_header_col = None

    for column in range(1, worksheet.max_column + 1):
        header_value = worksheet.cell(range_row, column).value
        if header_value and "Rango" in str(header_value):
            range_header_col = column
            break

    if range_header_col is None:
        return correction_bands, correction_cells, correction_operations, correction_factors

    for column in range(range_header_col + 1, worksheet.max_column + 1):
        range_value = worksheet.cell(range_row, column).value
        factor_value = worksheet.cell(factor_row, column).value
        if range_value in (None, "") or factor_value in (None, ""):
            continue

        label_text = str(range_value).strip()
        if label_text.upper() == "N/A":
            continue
        if "-" not in label_text:
            continue

        bounds = parse_range_bounds(label_text)
        if bounds is None:
            continue
        min_value, max_value = bounds

        key = f"range_{range_index}"
        range_index += 1
        operation, numeric_value = coerce_factor_value(factor_value)
        correction_bands[key] = {"label": label_text, "min": min_value, "max": max_value}
        correction_cells[key] = worksheet.cell(factor_row, column).coordinate
        correction_operations[key] = operation
        correction_factors[key] = numeric_value

    return correction_bands, correction_cells, correction_operations, correction_factors


def extract_cold_equipment_config(sheet_name: str, worksheet: Any) -> dict[str, Any]:
    correction_bands, correction_cells, correction_operations, correction_factors = extract_inline_corrections(
        worksheet,
        range_row=6,
        factor_row=7,
    )
    return {
        "sheet_name": sheet_name,
        "equipment_code": sheet_name,
        "laboratory": find_value_after_label(worksheet, "Laboratorio", [6]) or "",
        "equipment_name": find_value_after_label(worksheet, "Equipo", [7]) or "",
        "brand": find_value_after_label(worksheet, "Marca", [7]) or "",
        "model": find_value_after_label(worksheet, "Modelo", [7]) or "",
        "serial_number": find_value_after_label(worksheet, "No. Serie", [6]) or "",
        "inventory_code": find_value_after_label(worksheet, "Inventario / Código", [7])
        or find_value_after_label(worksheet, "Inventario/Código", [7])
        or "",
        "primary_label": find_value_after_label(worksheet, "Temperatura (ºC)", [8]) or "",
        "minimum_label": find_value_after_label(worksheet, "Mínima", [8]) or "",
        "maximum_label": find_value_after_label(worksheet, "Máxima", [8]) or "",
        "metadata_cells": {
            "equipment_name": find_cell_after_label(worksheet, "Equipo", [7]),
            "brand": find_cell_after_label(worksheet, "Marca", [7]),
            "model": find_cell_after_label(worksheet, "Modelo", [7]),
            "serial_number": find_cell_after_label(worksheet, "No. Serie", [6]),
            "inventory_code": find_cell_after_label(worksheet, "Inventario / Código", [7])
            or find_cell_after_label(worksheet, "Inventario/Código", [7]),
            "temperature_label": find_cell_after_label(worksheet, "Temperatura (ºC)", [8]),
            "minimum_label": find_cell_after_label(worksheet, "Mínima", [8]),
            "maximum_label": find_cell_after_label(worksheet, "Máxima", [8]),
        },
        "correction_bands": correction_bands,
        "correction_cells": correction_cells,
        "correction_factors": correction_factors,
        "correction_operations": correction_operations,
    }


def extract_incubator_config(sheet_name: str, worksheet: Any) -> dict[str, Any]:
    correction_bands, correction_cells, correction_operations, correction_factors = extract_inline_corrections(
        worksheet,
        range_row=6,
        factor_row=7,
    )
    return {
        "sheet_name": sheet_name,
        "equipment_code": sheet_name,
        "laboratory": find_value_after_label(worksheet, "Laboratorio", [6]) or "",
        "equipment_name": find_value_after_label(worksheet, "Equipo", [7]) or "",
        "brand": find_value_after_label(worksheet, "Marca", [7]) or "",
        "model": find_value_after_label(worksheet, "Modelo", [7]) or "",
        "serial_number": find_value_after_label(worksheet, "No. Serie", [6]) or "",
        "inventory_code": find_value_after_label(worksheet, "Inventario / Código", [7])
        or find_value_after_label(worksheet, "Inventario/Código", [7])
        or "",
        "primary_label": find_value_after_label(worksheet, "Temperatura (ºC)", [8]) or "",
        "minimum_label": find_value_after_label(worksheet, "Mínima", [8]) or "",
        "maximum_label": find_value_after_label(worksheet, "Máxima", [8]) or "",
        "secondary_label": find_value_after_label(worksheet, "%CO2", [9]) or "",
        "secondary_minimum_label": find_value_after_label(worksheet, "Mínima", [9]) or "",
        "secondary_maximum_label": find_value_after_label(worksheet, "Máxima", [9]) or "",
        "metadata_cells": {
            "equipment_name": find_cell_after_label(worksheet, "Equipo", [7]),
            "brand": find_cell_after_label(worksheet, "Marca", [7]),
            "model": find_cell_after_label(worksheet, "Modelo", [7]),
            "serial_number": find_cell_after_label(worksheet, "No. Serie", [6]),
            "inventory_code": find_cell_after_label(worksheet, "Inventario / Código", [7])
            or find_cell_after_label(worksheet, "Inventario/Código", [7]),
            "temperature_label": find_cell_after_label(worksheet, "Temperatura (ºC)", [8]),
            "minimum_label": find_cell_after_label(worksheet, "Mínima", [8]),
            "maximum_label": find_cell_after_label(worksheet, "Máxima", [8]),
            "secondary_label": find_cell_after_label(worksheet, "%CO2", [9]),
            "secondary_minimum_label": find_cell_after_label(worksheet, "Mínima", [9]),
            "secondary_maximum_label": find_cell_after_label(worksheet, "Máxima", [9]),
        },
        "correction_bands": correction_bands,
        "correction_cells": correction_cells,
        "correction_factors": correction_factors,
        "correction_operations": correction_operations,
    }


def extract_ambient_config(sheet_name: str, worksheet: Any) -> dict[str, Any]:
    is_humidity_sheet = "HUMEDAD" in sheet_name.upper()
    if is_humidity_sheet:
        uses_shifted_header = worksheet["O8"].value in (None, "") and worksheet["S8"].value in (None, "")
        if uses_shifted_header:
            brand_value = worksheet["Q8"].value or ""
            model_value = worksheet["W8"].value or ""
            serial_value = worksheet["AC8"].value or ""
            brand_cell = "Q8"
            model_cell = "W8"
            serial_cell = "AC8"
        else:
            brand_value = worksheet["O8"].value or ""
            model_value = worksheet["S8"].value or ""
            serial_value = worksheet["W8"].value or ""
            brand_cell = "O8"
            model_cell = "S8"
            serial_cell = "W8"
    else:
        brand_value = worksheet["Q8"].value or ""
        model_value = worksheet["W8"].value or ""
        serial_value = worksheet["AC8"].value or ""
        brand_cell = "Q8"
        model_cell = "W8"
        serial_cell = "AC8"
    primary_label_value = (
        find_value_after_label(worksheet, "TEMPERATURA (°C)", [9])
        or find_value_after_label(worksheet, "TEMPERATURA (Â°C)", [9])
        or find_value_after_label(worksheet, "% HUMEDAD", [9])
        or find_value_after_label(worksheet, "% Humedad Relativa", [9])
        or ""
    )
    primary_label_cell = (
        find_cell_after_label(worksheet, "TEMPERATURA (°C)", [9])
        or find_cell_after_label(worksheet, "TEMPERATURA (Â°C)", [9])
        or find_cell_after_label(worksheet, "% HUMEDAD", [9])
        or find_cell_after_label(worksheet, "% Humedad Relativa", [9])
    )
    correction_bands: dict[str, dict[str, Any]] = {}
    correction_cells: dict[str, str] = {}
    operations: dict[str, str] = {}
    factors: dict[str, float] = {}
    range_index = 1
    for column in range(1, worksheet.max_column + 1):
        range_label = worksheet.cell(7, column).value
        factor_value = worksheet.cell(8, column).value
        if not range_label or factor_value in (None, ""):
            continue
        label_text = str(range_label).strip()
        if "-" not in label_text:
            continue
        bounds = parse_range_bounds(label_text)
        if bounds is None:
            continue
        min_value, max_value = bounds

        key = f"range_{range_index}"
        range_index += 1
        op, num = coerce_factor_value(factor_value)
        operations[key] = op
        factors[key] = num
        correction_bands[key] = {"label": label_text, "min": min_value, "max": max_value}
        correction_cells[key] = worksheet.cell(8, column).coordinate
    return {
        "sheet_name": sheet_name,
        "equipment_code": sheet_name,
        "laboratory": find_value_after_label(worksheet, "Laboratorio", [6]) or "",
        "equipment_name": worksheet["E8"].value or sheet_name,
        "brand": brand_value,
        "model": model_value,
        "serial_number": serial_value,
        "inventory_code": worksheet["K8"].value or "",
        "primary_label": primary_label_value,
        "minimum_label": find_value_after_label(worksheet, "Mínima", [9]) or "",
        "maximum_label": find_value_after_label(worksheet, "Máxima", [9]) or "",
        "metadata_cells": {
            "equipment_name": "E8",
            "brand": brand_cell,
            "model": model_cell,
            "serial_number": serial_cell,
            "inventory_code": "K8",
            "temperature_label": primary_label_cell,
            "minimum_label": find_cell_after_label(worksheet, "Mínima", [9]),
            "maximum_label": find_cell_after_label(worksheet, "Máxima", [9]),
        },
        "correction_bands": correction_bands,
        "correction_cells": correction_cells,
        "correction_factors": factors,
        "correction_operations": operations,
    }


@st.cache_data(show_spinner=False)
def load_equipment_configs(form_key: str) -> dict[str, dict[str, Any]]:
    workbook = load_template_workbook(form_key, data_only=True)
    equipment_configs: dict[str, dict[str, Any]] = {}
    definition = get_form_definition(form_key)

    for sheet_name in workbook.sheetnames:
        if sheet_name in definition["sheet_exclusions"]:
            continue

        worksheet = workbook[sheet_name]
        extractor = definition["extractor"]
        if extractor == "cold_equipment":
            equipment_configs[sheet_name] = extract_cold_equipment_config(sheet_name, worksheet)
        elif extractor == "incubators":
            equipment_configs[sheet_name] = extract_incubator_config(sheet_name, worksheet)
        else:
            equipment_configs[sheet_name] = extract_ambient_config(sheet_name, worksheet)

        if not definition["supports_corrections"]:
            equipment_configs[sheet_name]["correction_bands"] = {}
            equipment_configs[sheet_name]["correction_cells"] = {}
            equipment_configs[sheet_name]["correction_factors"] = {}
            equipment_configs[sheet_name]["correction_operations"] = {}

    return equipment_configs


def build_default_payload(
    equipment_code: str = DEFAULT_EQUIPMENT_CODE,
    form_key: str = DEFAULT_FORM_KEY,
) -> dict[str, Any]:
    today = date.today()
    equipment_configs = load_equipment_configs(form_key)
    equipment_config = equipment_configs[equipment_code]
    definition = get_form_definition(form_key)
    return {
        "metadata": {
            "form_key": form_key,
            "form_label": definition["label"],
            "month": today.month,
            "year": today.year,
            "equipment_code": equipment_config["equipment_code"],
            "laboratory": equipment_config["laboratory"],
            "equipment_name": equipment_config["equipment_name"],
            "brand": equipment_config["brand"],
            "model": equipment_config["model"],
            "serial_number": equipment_config["serial_number"],
            "inventory_code": equipment_config["inventory_code"],
            "temperature_label": equipment_config["primary_label"],
            "minimum_label": equipment_config["minimum_label"],
            "maximum_label": equipment_config["maximum_label"],
            "secondary_label": equipment_config.get("secondary_label", ""),
            "secondary_minimum_label": equipment_config.get("secondary_minimum_label", ""),
            "secondary_maximum_label": equipment_config.get("secondary_maximum_label", ""),
        },
        "metadata_cells": dict(equipment_config.get("metadata_cells", {})),
        "correction_bands": dict(equipment_config.get("correction_bands", {})),
        "correction_cells": dict(equipment_config.get("correction_cells", {})),
        "correction_factors": dict(equipment_config["correction_factors"]),
        "correction_operations": dict(equipment_config["correction_operations"]),
        "non_working_days": [],
        "daily_records": {
            str(day): asdict(default_daily_capture(day)) for day in range(1, 32)
        },
        "change_log": [],
        "monthly_closure": {
            "observations": "",
            "reviewed_by": "",
            "reviewed_on": today.isoformat(),
        },
    }


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def merge_payload_with_saved_data(
    data: dict[str, Any],
    equipment_code: str = DEFAULT_EQUIPMENT_CODE,
    form_key: str = DEFAULT_FORM_KEY,
) -> dict[str, Any]:
    default_payload = build_default_payload(equipment_code=equipment_code, form_key=form_key)
    default_payload["metadata"].update(data.get("metadata", {}))
    default_payload["metadata_cells"] = data.get("metadata_cells", default_payload.get("metadata_cells", {}))
    default_payload["correction_factors"].update(data.get("correction_factors", {}))
    default_payload["correction_operations"].update(data.get("correction_operations", {}))
    default_payload["non_working_days"] = data.get("non_working_days", [])
    default_payload["change_log"] = data.get("change_log", [])
    default_payload["monthly_closure"].update(data.get("monthly_closure", {}))

    incoming_records = data.get("daily_records", {})
    for day in range(1, 32):
        day_key = str(day)
        default_payload["daily_records"][day_key].update(incoming_records.get(day_key, {}))
        record = default_payload["daily_records"][day_key]
        if "temperatures" in record and "measured_temperatures" not in incoming_records.get(day_key, {}):
            record["measured_temperatures"] = record.pop("temperatures")
        record.setdefault("measured_temperatures", ["", "", ""])
        record.setdefault("corrected_temperatures", ["", "", ""])
        record.setdefault("secondary_measurements", ["", "", ""])
        record.setdefault("performed_by_slots", ["", "", ""])

    return default_payload


def get_period_key(payload: dict[str, Any]) -> str:
    return (
        f"{payload['metadata']['form_key']}_"
        f"{payload['metadata']['equipment_code']}_"
        f"{int(payload['metadata']['year'])}_{int(payload['metadata']['month']):02d}"
    )


def get_period_file_from_values(form_key: str, equipment_code: str, year: int, month: int) -> Path:
    ensure_data_dir()
    normalized_form = form_key.lower().replace("-", "_")
    normalized_code = equipment_code.lower().replace("-", "_")
    return DATA_DIR / f"{normalized_form}_{normalized_code}_{year}_{month:02d}.json"


def get_period_file(payload: dict[str, Any]) -> Path:
    return get_period_file_from_values(
        str(payload["metadata"]["form_key"]),
        str(payload["metadata"]["equipment_code"]),
        int(payload["metadata"]["year"]),
        int(payload["metadata"]["month"]),
    )


def clear_period_widget_state(period_key: str) -> None:
    keys_to_delete = [
        key for key in st.session_state.keys() if period_key in str(key)
    ]
    for key in keys_to_delete:
        del st.session_state[key]


def load_saved_payload(
    form_key: str = DEFAULT_FORM_KEY,
    equipment_code: str = DEFAULT_EQUIPMENT_CODE,
    year: int | None = None,
    month: int | None = None,
) -> dict[str, Any]:
    default_payload = build_default_payload(equipment_code=equipment_code, form_key=form_key)
    if year is None:
        year = int(default_payload["metadata"]["year"])
    if month is None:
        month = int(default_payload["metadata"]["month"])

    if supabase_storage_enabled():
        try:
            remote_payload = load_remote_period_payload(form_key, equipment_code, year, month)
            if remote_payload:
                return merge_payload_with_saved_data(remote_payload, equipment_code=equipment_code, form_key=form_key)
        except Exception:
            pass

    data_file = get_period_file_from_values(form_key, equipment_code, year, month)
    if not data_file.exists():
        default_payload["metadata"]["year"] = year
        default_payload["metadata"]["month"] = month
        return default_payload

    with data_file.open("r", encoding="utf-8") as file:
        data = json.load(file)

    return merge_payload_with_saved_data(data, equipment_code=equipment_code, form_key=form_key)


def get_comparable_payload(payload: dict[str, Any]) -> dict[str, Any]:
    comparable_payload = json.loads(json.dumps(payload, ensure_ascii=False))
    comparable_payload.pop("change_log", None)
    return comparable_payload


def get_payload_signature(payload: dict[str, Any]) -> str:
    return json.dumps(get_comparable_payload(payload), ensure_ascii=False, sort_keys=True)


def remember_saved_snapshot(payload: dict[str, Any]) -> None:
    st.session_state["last_saved_payload_snapshot"] = get_comparable_payload(payload)
    st.session_state["last_saved_payload_signature"] = get_payload_signature(payload)


def format_change_log_item(payload: dict[str, Any], metric: dict[str, Any], value: str) -> str:
    formatted = format_metric_value(payload, metric, value)
    return formatted or "vacio"


def build_change_log_entry(previous_payload: dict[str, Any], current_payload: dict[str, Any]) -> dict[str, Any] | None:
    definition = get_form_definition(current_payload["metadata"]["form_key"])
    items: list[str] = []

    for day in range(1, 32):
        day_key = str(day)
        previous_record = previous_payload.get("daily_records", {}).get(day_key, {})
        current_record = current_payload.get("daily_records", {}).get(day_key, {})

        for metric in definition["metrics"]:
            metric_label = get_primary_metric_display_label(current_payload, metric)
            previous_values = previous_record.get(metric["key"], ["", "", ""])
            current_values = current_record.get(metric["key"], ["", "", ""])
            for index, slot_label in enumerate(TIME_SLOTS):
                previous_value = str(previous_values[index] if index < len(previous_values) else "").strip()
                current_value = str(current_values[index] if index < len(current_values) else "").strip()
                if not previous_value or previous_value == current_value:
                    continue
                old_display = format_change_log_item(current_payload, metric, previous_value)
                new_display = format_change_log_item(current_payload, metric, current_value)
                items.append(f"Cambio: Dia {day} {metric_label} {slot_label}: {old_display} -> {new_display}")

        previous_verified = str(previous_record.get("verified_by", "")).strip()
        current_verified = str(current_record.get("verified_by", "")).strip()
        if previous_verified and previous_verified != current_verified:
            items.append(f"Cambio: Dia {day} Verifico: {previous_verified} -> {current_verified or 'vacio'}")

        previous_date = str(previous_record.get("recorded_on", "")).strip()
        current_date = str(current_record.get("recorded_on", "")).strip()
        if previous_date and previous_date != current_date:
            items.append(
                f"Cambio: Dia {day} Fecha de verificacion: {normalize_excel_date(previous_date)} -> {normalize_excel_date(current_date)}"
            )

    if not items:
        return None

    return {
        "timestamp": get_local_now().isoformat(timespec="seconds"),
        "user": str(st.session_state.get("usuario_email", "")).strip(),
        "items": items,
    }


def compose_observations_export_text(payload: dict[str, Any]) -> str:
    base_observations = str(payload["monthly_closure"].get("observations", "")).strip()
    non_working_days = sorted(int(day) for day in payload.get("non_working_days", []))
    change_log = payload.get("change_log", [])
    sections: list[str] = []

    if base_observations:
        sections.append(base_observations)

    if non_working_days:
        days_text = ", ".join(str(day) for day in non_working_days)
        sections.append(f"* Los dias {days_text} no aplican por marcarse como no laborados.")

    if not change_log:
        return "\n\n".join(sections).strip()

    audit_lines: list[str] = []
    for entry in change_log[-6:]:
        items = entry.get("items", [])
        if not items:
            continue
        visible_items = items[:3]
        detail = "; ".join(str(item) for item in visible_items)
        if len(items) > len(visible_items):
            detail += f"; y {len(items) - len(visible_items)} cambio(s) mas"
        audit_lines.append(detail)

    if not audit_lines:
        return "\n\n".join(sections).strip()

    audit_block = "Registro de correcciones:\n" + "\n".join(audit_lines)
    sections.append(audit_block)
    return "\n\n".join(section for section in sections if section).strip()


def save_payload(payload: dict[str, Any]) -> str:
    previous_snapshot = st.session_state.get("last_saved_payload_snapshot")
    if isinstance(previous_snapshot, dict):
        change_entry = build_change_log_entry(previous_snapshot, get_comparable_payload(payload))
        if change_entry is not None:
            payload.setdefault("change_log", []).append(change_entry)
            payload["change_log"] = payload["change_log"][-80:]

    if supabase_storage_enabled():
        try:
            save_remote_period_payload(
                payload,
                updated_by=str(st.session_state.get("usuario_email", "")).strip(),
            )
            remember_saved_snapshot(payload)
            return "Supabase"
        except Exception:
            pass

    data_file = get_period_file(payload)
    with data_file.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
    remember_saved_snapshot(payload)
    return "Local JSON"


def get_supabase_status() -> tuple[bool, str]:
    if supabase_storage_enabled():
        return True, "Supabase"
    return False, "Local JSON"


def get_row_group(payload: dict[str, Any], day: int) -> dict[str, int]:
    layout = get_form_definition(payload["metadata"]["form_key"])["layout"]
    return layout["top"] if day <= 16 else layout["bottom"]


def get_day_column(day: int, slot_index: int) -> int:
    if day <= 16:
        return DAY_BLOCK_START_COLUMNS[day] + slot_index
    return DAY_BLOCK_START_COLUMNS[day] + slot_index


def normalize_excel_date(date_value: str) -> str:
    if not date_value:
        return ""

    try:
        parsed = datetime.fromisoformat(date_value)
        return parsed.strftime("%d/%m/%Y")
    except ValueError:
        return date_value


def get_effective_record_date(record: dict[str, Any]) -> str:
    recorded_on = record.get("recorded_on", "").strip()
    if recorded_on:
        return normalize_excel_date(recorded_on)
    return ""


def get_format_specific_copy(payload: dict[str, Any]) -> dict[str, str]:
    form_key = payload["metadata"]["form_key"]
    if form_key == "incubadoras":
        return {
            "config_intro": "Este formato registra temperatura y %CO2. La corrección aplica solo a la temperatura.",
            "daily_intro": "Cada día laborado captura temperatura, %CO2, tres responsables por horario y un verificador del bloque.",
            "notes_label": "Incidencia del día (opcional)",
        }
    if form_key == "condiciones_ambientales":
        return {
            "config_intro": "Este formato monitorea condiciones ambientales con un thermohigrómetro y rangos de corrección específicos por hoja.",
            "daily_intro": "Cada día laborado registra lecturas ambientales por horario, responsables, verificación y fecha del bloque.",
            "notes_label": "Incidencia ambiental del día (opcional)",
        }
    if form_key == "ultracongeladores":
        return {
            "config_intro": "Este formato controla ultracongeladores. Algunos equipos usan factores reales y otros están marcados como N/A.",
            "daily_intro": "Cada día laborado captura temperatura por horario, responsables, verificación y fecha del bloque.",
            "notes_label": "Incidencia del ultracongelador (opcional)",
        }
    if form_key == "refrigeradores":
        return {
            "config_intro": "Este formato controla refrigeradores con rangos y factores específicos por equipo.",
            "daily_intro": "Cada día laborado captura temperatura por horario, responsables, verificación y fecha del bloque.",
            "notes_label": "Incidencia del refrigerador (opcional)",
        }
    return {
        "config_intro": "Este formato controla congeladores con rangos y factores específicos por equipo.",
        "daily_intro": "Cada día laborado captura temperatura por horario, responsables, verificación y fecha del bloque.",
        "notes_label": "Incidencia del congelador (opcional)",
    }


def format_percentage_display(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if "%" in text:
        return text

    normalized = text.replace(",", ".")
    try:
        numeric = float(normalized)
    except ValueError:
        return text

    percentage_value = numeric * 100 if abs(numeric) <= 1 else numeric
    if percentage_value.is_integer():
        return f"{int(percentage_value)}%"
    return f"{percentage_value:.2f}".rstrip("0").rstrip(".") + "%"


def is_ambient_form(payload: dict[str, Any]) -> bool:
    return payload["metadata"]["form_key"] == "condiciones_ambientales"


def is_ambient_humidity_payload(payload: dict[str, Any]) -> bool:
    equipment_code = str(payload["metadata"].get("equipment_code", "")).upper()
    return is_ambient_form(payload) and equipment_code.startswith("HUMEDAD")


def get_ambient_variable_name(payload: dict[str, Any]) -> str:
    return "% Humedad" if is_ambient_humidity_payload(payload) else "Temperatura"


def get_primary_metric_display_label(payload: dict[str, Any], metric: dict[str, Any]) -> str:
    if is_ambient_form(payload):
        return "Humedad medida" if is_ambient_humidity_payload(payload) else "Temperatura medida"
    return str(metric["label"])


def is_humidity_metric(payload: dict[str, Any], metric: dict[str, Any]) -> bool:
    return is_ambient_humidity_payload(payload) and metric["key"] == "measured_temperatures"


def is_incubator_co2_metric(payload: dict[str, Any], metric: dict[str, Any]) -> bool:
    return (
        payload["metadata"]["form_key"] == "incubadoras"
        and metric["key"] == "secondary_measurements"
    )


def parse_measurement_number(raw_value: str) -> float | None:
    cleaned = (
        raw_value.strip()
        .replace(",", ".")
        .replace("%", "")
        .replace("≤", "")
        .replace("≥", "")
    )
    if not cleaned:
        return None

    try:
        numeric = float(cleaned)
    except ValueError:
        return None

    return numeric * 100 if 0 < abs(numeric) <= 1 else numeric


def format_metric_value(payload: dict[str, Any], metric: dict[str, Any], value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if is_humidity_metric(payload, metric):
        return format_percentage_display(text)
    if is_incubator_co2_metric(payload, metric):
        return format_percentage_display(text)
    return text


def render_sidebar(
    payload: dict[str, Any],
    form_keys: list[str],
    equipment_codes: list[str],
) -> tuple[str, str]:
    st.sidebar.title("Formatos")
    st.sidebar.write(f"Sesion: `{st.session_state.get('usuario_email', '')}`")
    st.sidebar.write(f"Perfil: `{st.session_state.get('rol_usuario', 'captura')}`")
    if st.sidebar.button("Cerrar sesion", use_container_width=True):
        st.session_state["autenticado"] = False
        st.session_state["usuario_email"] = ""
        st.session_state["es_admin"] = False
        st.session_state["rol_usuario"] = "captura"
        st.rerun()

    selected_form_key = st.sidebar.selectbox(
        "Formato",
        options=form_keys,
        index=form_keys.index(payload["metadata"]["form_key"]),
        format_func=lambda key: FORM_DEFINITIONS[key]["label"],
        key="sidebar_form_key",
    )
    selected_equipment = st.sidebar.selectbox(
        "Equipo",
        options=equipment_codes,
        index=equipment_codes.index(payload["metadata"]["equipment_code"]),
        key="sidebar_equipment",
    )
    st.sidebar.caption(payload["metadata"]["form_label"])
    st.sidebar.write(
        f"Mes de trabajo: `{MONTHS[payload['metadata']['month']]} {payload['metadata']['year']}`"
    )
    st.sidebar.write(f"Laboratorio: `{payload['metadata']['laboratory']}`")
    st.sidebar.write(f"Equipo / instrumento: `{payload['metadata']['equipment_name']}`")
    render_user_admin_sidebar()
    return selected_form_key, selected_equipment


def render_configuration(payload: dict[str, Any]) -> None:
    st.subheader("1. Configuracion del mes")
    metadata = payload["metadata"]
    correction_bands = payload["correction_bands"]
    correction_factors = payload["correction_factors"]
    correction_operations = payload["correction_operations"]
    period_key = st.session_state.get("period_key", get_period_key(payload))
    copy = get_format_specific_copy(payload)
    allow_sensitive_edits = can_edit_sensitive_configuration()

    month_col, year_col = st.columns(2)
    metadata["month"] = month_col.selectbox(
        "Mes",
        options=list(MONTHS.keys()),
        index=list(MONTHS.keys()).index(metadata["month"]),
        format_func=lambda value: MONTHS[value],
    )
    metadata["year"] = year_col.number_input(
        "Año",
        min_value=2024,
        max_value=2100,
        value=int(metadata["year"]),
        step=1,
    )
    st.caption(copy["config_intro"])
    if allow_sensitive_edits:
        st.caption("Edicion sensible habilitada para este perfil.")
    else:
        st.caption("Rangos, inventario y metadata del equipo estan en solo lectura para este perfil.")

    definition = get_form_definition(metadata["form_key"])
    is_incubator_form = metadata["form_key"] == "incubadoras"
    ambient_variable_name = get_ambient_variable_name(payload) if is_ambient_form(payload) else ""
    is_ambient_humidity = is_ambient_humidity_payload(payload)
    summary_cols = st.columns(4)
    equipment_name_key = f"equipment_name_{period_key}"
    if equipment_name_key not in st.session_state:
        st.session_state[equipment_name_key] = str(metadata["equipment_name"])
    metadata["equipment_name"] = summary_cols[0].text_input(
        "Equipo / instrumento",
        key=equipment_name_key,
        disabled=not allow_sensitive_edits,
    )
    brand_key = f"brand_{period_key}"
    if brand_key not in st.session_state:
        st.session_state[brand_key] = str(metadata["brand"])
    metadata["brand"] = summary_cols[1].text_input(
        "Marca",
        key=brand_key,
        disabled=not allow_sensitive_edits,
    )
    model_key = f"model_{period_key}"
    if model_key not in st.session_state:
        st.session_state[model_key] = str(metadata["model"])
    metadata["model"] = summary_cols[2].text_input(
        "Modelo",
        key=model_key,
        disabled=not allow_sensitive_edits,
    )
    inventory_key = f"inventory_{period_key}"
    if inventory_key not in st.session_state:
        st.session_state[inventory_key] = str(metadata["inventory_code"])
    metadata["inventory_code"] = summary_cols[3].text_input(
        "Inventario / código",
        key=inventory_key,
        disabled=not allow_sensitive_edits,
    )

    detail_cols = st.columns(4)
    serial_key = f"serial_{period_key}"
    if serial_key not in st.session_state:
        st.session_state[serial_key] = str(metadata["serial_number"])
    metadata["serial_number"] = detail_cols[0].text_input(
        "Serie",
        key=serial_key,
        disabled=not allow_sensitive_edits,
    )
    normal_key = f"normal_{period_key}"
    if normal_key not in st.session_state:
        st.session_state[normal_key] = (
            format_percentage_display(metadata["temperature_label"])
            if is_ambient_humidity
            else str(metadata["temperature_label"])
        )
    metadata["temperature_label"] = detail_cols[1].text_input(
        f"{ambient_variable_name} normal" if ambient_variable_name else "Valor normal",
        key=normal_key,
        disabled=not allow_sensitive_edits,
    )
    minimum_key = f"minimum_{period_key}"
    if minimum_key not in st.session_state:
        st.session_state[minimum_key] = (
            format_percentage_display(metadata["minimum_label"])
            if is_ambient_humidity
            else str(metadata["minimum_label"])
        )
    metadata["minimum_label"] = detail_cols[2].text_input(
        f"Mínima {ambient_variable_name.lower()}" if ambient_variable_name else "Mínima",
        key=minimum_key,
        disabled=not allow_sensitive_edits,
    )
    maximum_key = f"maximum_{period_key}"
    if maximum_key not in st.session_state:
        st.session_state[maximum_key] = (
            format_percentage_display(metadata["maximum_label"])
            if is_ambient_humidity
            else str(metadata["maximum_label"])
        )
    metadata["maximum_label"] = detail_cols[3].text_input(
        f"Máxima {ambient_variable_name.lower()}" if ambient_variable_name else "Máxima",
        key=maximum_key,
        disabled=not allow_sensitive_edits,
    )

    if metadata.get("secondary_label"):
        secondary_section_title = "%CO2 visible" if is_incubator_form else "Variable secundaria visible"
        secondary_primary_label = "%CO2 normal" if is_incubator_form else "Nombre de la segunda variable"
        secondary_minimum_label = "Mínima %CO2" if is_incubator_form else "Mínima secundaria"
        secondary_maximum_label = "Máxima %CO2" if is_incubator_form else "Máxima secundaria"
        st.markdown(f"**{secondary_section_title}**")
        secondary_cols = st.columns(3)
        secondary_label_key = f"secondary_label_{period_key}"
        if secondary_label_key not in st.session_state:
            st.session_state[secondary_label_key] = (
                format_percentage_display(metadata["secondary_label"])
                if is_incubator_form
                else str(metadata["secondary_label"])
            )
        elif is_incubator_form:
            st.session_state[secondary_label_key] = format_percentage_display(st.session_state[secondary_label_key])
        metadata["secondary_label"] = secondary_cols[0].text_input(
            secondary_primary_label,
            key=secondary_label_key,
            disabled=not allow_sensitive_edits,
        )
        secondary_minimum_key = f"secondary_minimum_{period_key}"
        if secondary_minimum_key not in st.session_state:
            st.session_state[secondary_minimum_key] = (
                format_percentage_display(metadata.get("secondary_minimum_label", ""))
                if is_incubator_form
                else str(metadata.get("secondary_minimum_label", ""))
            )
        elif is_incubator_form:
            st.session_state[secondary_minimum_key] = format_percentage_display(st.session_state[secondary_minimum_key])
        metadata["secondary_minimum_label"] = secondary_cols[1].text_input(
            secondary_minimum_label,
            key=secondary_minimum_key,
            disabled=not allow_sensitive_edits,
        )
        secondary_maximum_key = f"secondary_maximum_{period_key}"
        if secondary_maximum_key not in st.session_state:
            st.session_state[secondary_maximum_key] = (
                format_percentage_display(metadata.get("secondary_maximum_label", ""))
                if is_incubator_form
                else str(metadata.get("secondary_maximum_label", ""))
            )
        elif is_incubator_form:
            st.session_state[secondary_maximum_key] = format_percentage_display(st.session_state[secondary_maximum_key])
        metadata["secondary_maximum_label"] = secondary_cols[2].text_input(
            secondary_maximum_label,
            key=secondary_maximum_key,
            disabled=not allow_sensitive_edits,
        )

    if definition["supports_corrections"] and correction_factors:
        st.caption("Factores de correccion editables por rango para este equipo.")
        factor_labels = list(correction_factors.keys())
        if is_ambient_form(payload):
            ranges_per_row = 3
            for start_index in range(0, len(factor_labels), ranges_per_row):
                row_keys = factor_labels[start_index:start_index + ranges_per_row]
                factor_cols = st.columns(len(row_keys))
                for factor_col, factor_key in zip(factor_cols, row_keys):
                    label = correction_bands.get(factor_key, {}).get("label", factor_key.replace("_", " ").title())
                    factor_col.markdown(f"**{ambient_variable_name} {label}**")
                    operation_col, value_col = factor_col.columns([1, 2])
                    correction_operations[factor_key] = operation_col.selectbox(
                        "Operacion",
                        options=["+", "-"],
                        index=0 if correction_operations[factor_key] == "+" else 1,
                        key=f"operation_{period_key}_{factor_key}",
                        disabled=not allow_sensitive_edits,
                    )
                    correction_factors[factor_key] = value_col.number_input(
                        "Factor de correccion",
                        value=float(correction_factors[factor_key]),
                        step=0.01,
                        format="%.2f",
                        key=f"factor_{period_key}_{factor_key}",
                        disabled=not allow_sensitive_edits,
                    )
        else:
            factor_cols = st.columns(len(factor_labels))
            for factor_col, factor_key in zip(factor_cols, factor_labels):
                operation_col, value_col = factor_col.columns([1, 2])
                label = correction_bands.get(factor_key, {}).get("label", factor_key.replace("_", " ").title())
                operation_label = f"Operacion temperatura {label}" if is_incubator_form else f"Operacion {label}"
                factor_value_label = f"Temperatura {label}" if is_incubator_form else label
                correction_operations[factor_key] = operation_col.selectbox(
                    operation_label,
                    options=["+", "-"],
                    index=0 if correction_operations[factor_key] == "+" else 1,
                    key=f"operation_{period_key}_{factor_key}",
                    disabled=not allow_sensitive_edits,
                )
                correction_factors[factor_key] = value_col.number_input(
                    factor_value_label,
                    value=float(correction_factors[factor_key]),
                    step=0.01,
                    format="%.2f",
                    key=f"factor_{period_key}_{factor_key}",
                    disabled=not allow_sensitive_edits,
                )
    else:
        st.caption(
            "Este equipo no usa factores de corrección editables en la plantilla o están marcados como N/A."
        )


def render_non_working_days(payload: dict[str, Any]) -> None:
    st.subheader("2. Dias no laborados")
    st.caption("Marca manualmente los dias que no aplican para la toma. Cada boton activa o desactiva el dia.")
    allow_schedule_edits = can_edit_schedule()
    if not allow_schedule_edits:
        st.caption("Este apartado esta en solo lectura para tu perfil.")
    period_key = st.session_state.get("period_key", get_period_key(payload))
    selected_days = set(int(day) for day in payload["non_working_days"])
    days = list(range(1, 32))
    for week_start in range(0, len(days), 7):
        cols = st.columns(7)
        for offset, day in enumerate(days[week_start:week_start + 7]):
            checkbox_key = f"non_working_{period_key}_{day}"
            if checkbox_key not in st.session_state:
                st.session_state[checkbox_key] = day in selected_days
            is_selected = cols[offset].checkbox(
                f"{day}",
                key=checkbox_key,
                disabled=not allow_schedule_edits,
            )
            if is_selected:
                selected_days.add(day)
            else:
                selected_days.discard(day)
    payload["non_working_days"] = sorted(selected_days)

    for day in range(1, 32):
        payload["daily_records"][str(day)]["active"] = day not in payload["non_working_days"]


def render_daily_capture(payload: dict[str, Any]) -> None:
    st.subheader("3. Captura diaria")
    copy = get_format_specific_copy(payload)
    st.caption(copy["daily_intro"])
    allow_daily_edits = can_edit_daily_records()
    allow_verification_edits = can_close_period()
    if not allow_daily_edits:
        st.caption("La captura diaria esta en solo lectura para tu perfil.")
    period_key = st.session_state.get("period_key", get_period_key(payload))
    definition = get_form_definition(payload["metadata"]["form_key"])
    metrics = definition["metrics"]

    active_days = [
        day for day in range(1, 32) if payload["daily_records"][str(day)]["active"]
    ]
    if not active_days:
        st.info("No hay dias activos. Marca al menos un dia laborado para capturar informacion.")
        return

    preferred_day = get_preferred_capture_day(payload, active_days)
    for day in active_days:
        record = payload["daily_records"][str(day)]
        with st.expander(f"Dia {day}", expanded=day == preferred_day):
            for metric in metrics:
                metric_label = get_primary_metric_display_label(payload, metric)
                if len(metrics) > 1 or is_ambient_form(payload):
                    if metric["key"] == "measured_temperatures":
                        if is_ambient_form(payload):
                            heading = get_ambient_variable_name(payload)
                        elif payload["metadata"]["form_key"] == "incubadoras":
                            heading = "Temperatura medida"
                        else:
                            heading = payload["metadata"].get("temperature_label", "Temperatura")
                    else:
                        heading = "%CO2" if payload["metadata"]["form_key"] == "incubadoras" else payload["metadata"].get("secondary_label", "Variable secundaria")
                    st.markdown(f"**{heading}**")
                metric_cols = st.columns(3)
                metric_values: list[str] = []
                corrected_values: list[str] = []
                for index, label in enumerate(TIME_SLOTS):
                    input_key = f"{metric['key']}_{period_key}_{day}_{index}"
                    if input_key not in st.session_state:
                        st.session_state[input_key] = record[metric["key"]][index]
                    value = metric_cols[index].text_input(
                        f"{metric_label} {label}",
                        key=input_key,
                        disabled=not allow_daily_edits,
                        placeholder="-20.3" if metric["unit"] == "°C" else "",
                    )
                    metric_values.append(value)
                    if metric.get("corrected", False):
                        corrected_value = calculate_corrected_temperature(
                            value,
                            payload["correction_bands"],
                            payload["correction_factors"],
                            payload["correction_operations"],
                        )
                        corrected_values.append(corrected_value)
                        corrected_display = format_metric_value(payload, metric, corrected_value)
                        metric_cols[index].caption(
                            f"Corregida: {corrected_display}" if corrected_display else "Corregida: pendiente"
                        )
                record[metric["key"]] = metric_values
                if metric["key"] == "measured_temperatures":
                    record["corrected_temperatures"] = corrected_values

            actor_cols = st.columns(3)
            performed_by_slots = []
            for index, label in enumerate(TIME_SLOTS):
                input_key = f"performed_{period_key}_{day}_{index}"
                if input_key not in st.session_state:
                    st.session_state[input_key] = record["performed_by_slots"][index]
                performed_value = actor_cols[index].text_input(
                    f"Realizo {label}",
                    key=input_key,
                    disabled=not allow_daily_edits,
                )
                performed_by_slots.append(performed_value)
                performed_signature_name = get_signature_display_name(performed_value)
                if performed_value.strip():
                    if performed_signature_name:
                        actor_cols[index].caption(f"Firma detectada: {performed_signature_name}")
                    else:
                        actor_cols[index].caption("Firma digital: sin coincidencia")
                else:
                    actor_cols[index].caption("Firma digital: pendiente")
            record["performed_by_slots"] = performed_by_slots

            verifier_cols = st.columns(3)
            verified_key = f"verified_{period_key}_{day}"
            if verified_key not in st.session_state:
                st.session_state[verified_key] = record["verified_by"]
            record["verified_by"] = verifier_cols[0].text_input(
                "Verifico bloque del dia",
                key=verified_key,
                disabled=not allow_verification_edits,
            )
            verified_signature_name = get_signature_display_name(record["verified_by"])
            if record["verified_by"].strip():
                if verified_signature_name:
                    verifier_cols[0].caption(f"Firma detectada: {verified_signature_name}")
                else:
                    verifier_cols[0].caption("Firma digital: sin coincidencia")
            else:
                verifier_cols[0].caption("Firma digital: pendiente")
            current_date = get_period_default_date(payload, day, record["recorded_on"])
            date_key = f"date_{period_key}_{day}"
            if date_key not in st.session_state:
                st.session_state[date_key] = current_date
            elif is_current_period(payload) and uses_default_daily_record_date(payload, day, record["recorded_on"]):
                st.session_state[date_key] = current_date
            record["recorded_on"] = verifier_cols[1].date_input(
                "Fecha de verificacion",
                min_value=date(
                    int(payload["metadata"]["year"]),
                    int(payload["metadata"]["month"]),
                    1,
                ),
                max_value=date(
                    int(payload["metadata"]["year"]),
                    int(payload["metadata"]["month"]),
                    monthrange(
                        int(payload["metadata"]["year"]),
                        int(payload["metadata"]["month"]),
                    )[1],
                ),
                key=date_key,
                disabled=not allow_verification_edits,
            ).isoformat()

            notes_key = f"notes_{period_key}_{day}"
            if notes_key not in st.session_state:
                st.session_state[notes_key] = record.get("notes", "")
            record["notes"] = st.text_input(
                copy["notes_label"],
                key=notes_key,
                disabled=not allow_daily_edits,
            )


def parse_streamlit_date(value: str) -> date:
    try:
        return datetime.fromisoformat(value).date()
    except ValueError:
        return date.today()


def is_current_period(payload: dict[str, Any]) -> bool:
    today = date.today()
    return (
        int(payload["metadata"]["year"]) == today.year
        and int(payload["metadata"]["month"]) == today.month
    )


def get_preferred_capture_day(payload: dict[str, Any], active_days: list[int]) -> int:
    today = date.today()
    if is_current_period(payload) and today.day in active_days:
        return today.day
    return active_days[0]


def get_row_period_date(payload: dict[str, Any], day: int) -> date:
    year = int(payload["metadata"]["year"])
    month = int(payload["metadata"]["month"])
    last_day = monthrange(year, month)[1]
    return date(year, month, min(day, last_day))


def uses_default_daily_record_date(payload: dict[str, Any], day: int, recorded_on: str) -> bool:
    if not recorded_on.strip():
        return True
    parsed = parse_streamlit_date(recorded_on)
    return parsed == get_row_period_date(payload, day)


def get_period_default_date(payload: dict[str, Any], day: int, recorded_on: str) -> date:
    year = int(payload["metadata"]["year"])
    month = int(payload["metadata"]["month"])
    if is_current_period(payload) and uses_default_daily_record_date(payload, day, recorded_on):
        return date.today()

    if recorded_on.strip():
        parsed = parse_streamlit_date(recorded_on)
        if parsed.year == year and parsed.month == month:
            return parsed

    return get_row_period_date(payload, day)


def get_period_end_date(payload: dict[str, Any]) -> date:
    year = int(payload["metadata"]["year"])
    month = int(payload["metadata"]["month"])
    last_day = monthrange(year, month)[1]
    return date(year, month, last_day)


def uses_default_closure_date(payload: dict[str, Any], reviewed_on: str) -> bool:
    if not reviewed_on.strip():
        return True
    parsed = parse_streamlit_date(reviewed_on)
    return parsed == get_period_end_date(payload)


def get_closure_default_date(payload: dict[str, Any]) -> date:
    closure = payload["monthly_closure"]
    reviewed_on = str(closure.get("reviewed_on", "")).strip()
    reviewed_by = str(closure.get("reviewed_by", "")).strip()

    if is_current_period(payload) and uses_default_closure_date(payload, reviewed_on):
        return date.today()

    if reviewed_on:
        parsed = parse_streamlit_date(reviewed_on)
        if parsed.year == int(payload["metadata"]["year"]) and parsed.month == int(payload["metadata"]["month"]):
            return parsed

    if reviewed_by:
        return parse_streamlit_date(reviewed_on)

    return get_period_end_date(payload)


def render_monthly_closure(payload: dict[str, Any]) -> None:
    st.subheader("4. Cierre del formato")
    closure = payload["monthly_closure"]
    period_key = st.session_state.get("period_key", get_period_key(payload))
    allow_closure_edits = can_close_period()
    if not allow_closure_edits:
        st.caption("El cierre del formato esta en solo lectura para tu perfil.")

    observations_key = f"observations_{period_key}"
    if observations_key not in st.session_state:
        st.session_state[observations_key] = closure["observations"]
    closure["observations"] = st.text_area(
        "Observaciones",
        key=observations_key,
        height=120,
        placeholder="Anota incidencias, mantenimiento, ajustes o aclaraciones del periodo.",
        disabled=not allow_closure_edits,
    )
    if payload.get("change_log"):
        st.caption("Las correcciones sobre capturas previas se anexan automaticamente en Observaciones al exportar.")
    review_cols = st.columns(2)
    reviewed_by_key = f"reviewed_by_{period_key}"
    if reviewed_by_key not in st.session_state:
        st.session_state[reviewed_by_key] = closure["reviewed_by"]
    closure["reviewed_by"] = review_cols[0].text_input(
        "Reviso",
        key=reviewed_by_key,
        disabled=not allow_closure_edits,
    )
    reviewed_signature_name = get_signature_display_name(closure["reviewed_by"])
    if closure["reviewed_by"].strip():
        if reviewed_signature_name:
            review_cols[0].caption(f"Firma detectada: {reviewed_signature_name}")
        else:
            review_cols[0].caption("Firma digital: sin coincidencia")
    else:
        review_cols[0].caption("Firma digital: pendiente")
    reviewed_on_key = f"reviewed_on_{period_key}"
    default_review_date = get_closure_default_date(payload)
    if reviewed_on_key not in st.session_state:
        st.session_state[reviewed_on_key] = default_review_date
    elif is_current_period(payload) and uses_default_closure_date(payload, str(closure.get("reviewed_on", ""))):
        st.session_state[reviewed_on_key] = default_review_date
    closure["reviewed_on"] = review_cols[1].date_input(
        "Fecha de revision",
        key=reviewed_on_key,
        disabled=not allow_closure_edits,
    ).isoformat()


def validate_payload(payload: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    definition = get_form_definition(payload["metadata"]["form_key"])
    for day in range(1, 32):
        record = payload["daily_records"][str(day)]
        if not record["active"]:
            continue
        for metric in definition["metrics"]:
            if not all(value.strip() for value in record[metric["key"]]):
                metric_label = get_primary_metric_display_label(payload, metric).lower()
                errors.append(f"Dia {day}: faltan capturas de {metric_label}.")
        if not all(value.strip() for value in record["performed_by_slots"]):
            errors.append(f"Dia {day}: faltan responsables en una o mas horas.")
        if not record["verified_by"].strip():
            errors.append(f"Dia {day}: falta 'Verifico'.")
        if not record["recorded_on"].strip():
            errors.append(f"Dia {day}: falta la fecha.")

    if not payload["monthly_closure"]["reviewed_by"].strip():
        errors.append("Falta capturar quien reviso el formato.")

    return errors


def populate_template(payload: dict[str, Any]) -> BytesIO:
    form_key = payload["metadata"]["form_key"]
    definition = get_form_definition(form_key)
    workbook = load_template_workbook(form_key)
    target_sheet_name = str(payload["metadata"]["equipment_code"])
    worksheet = workbook[target_sheet_name]
    ensure_status_row_merges(worksheet, payload)

    metadata = payload["metadata"]
    metadata_cells = payload.get("metadata_cells", {})
    editable_metadata_map = {
        "equipment_name": metadata.get("equipment_name", ""),
        "brand": metadata.get("brand", ""),
        "model": metadata.get("model", ""),
        "inventory_code": metadata.get("inventory_code", ""),
        "serial_number": metadata.get("serial_number", ""),
        "temperature_label": metadata.get("temperature_label", ""),
        "minimum_label": metadata.get("minimum_label", ""),
        "maximum_label": metadata.get("maximum_label", ""),
        "secondary_label": metadata.get("secondary_label", ""),
        "secondary_minimum_label": metadata.get("secondary_minimum_label", ""),
        "secondary_maximum_label": metadata.get("secondary_maximum_label", ""),
    }
    for field_key, value in editable_metadata_map.items():
        cell = metadata_cells.get(field_key)
        if cell:
            write_template_cell(worksheet, cell, value)

    month_cell = HEADER_CELL_MAP["month"]
    year_cell = HEADER_CELL_MAP["year"]
    if form_key == "condiciones_ambientales":
        month_cell = "V6"
        year_cell = "AO6"
    write_template_cell(worksheet, month_cell, MONTHS[payload["metadata"]["month"]])
    write_template_cell(worksheet, year_cell, payload["metadata"]["year"])

    if definition["supports_corrections"]:
        factor_cells = payload.get("correction_cells", {})
        for factor_key, cell in factor_cells.items():
            if factor_key not in payload["correction_factors"]:
                continue
            factor_value = float(payload["correction_factors"][factor_key])
            if payload["correction_operations"][factor_key] == "-":
                factor_value *= -1
            write_template_cell(worksheet, cell, factor_value)

    footer_map = definition["layout"]["footer"]

    for day in range(1, 32):
        record = payload["daily_records"][str(day)]
        row_group = get_row_group(payload, day)
        start_col = get_day_column(day, 0)

        if not record["active"]:
            for index in range(3):
                write_slot_value(
                    worksheet,
                    row_group["metric_1"],
                    start_col + index,
                    "N/A",
                    font_size=18,
                    rotate_like_hours=True,
                )
                if "metric_2" in row_group:
                    write_slot_value(
                        worksheet,
                        row_group["metric_2"],
                        start_col + index,
                        "N/A",
                        font_size=18,
                        rotate_like_hours=True,
                    )
                write_slot_value(
                    worksheet,
                    row_group["performed_by"],
                    start_col + index,
                    "NO LABORADO",
                    font_size=18,
                    rotate_like_hours=True,
                )
            write_day_status(worksheet, row_group["verified_by"], start_col, "")
            write_day_status(
                worksheet,
                row_group["date"],
                start_col,
                "",
            )
            continue

        primary_values = (
            record["corrected_temperatures"]
            if any(record["corrected_temperatures"]) and definition["metrics"][0].get("corrected")
            else record["measured_temperatures"]
        )
        primary_metric = definition["metrics"][0]
        primary_unit = primary_metric["unit"]
        for index, temperature in enumerate(primary_values):
            formatted_value = format_metric_value(payload, primary_metric, temperature)
            write_slot_value(
                worksheet,
                row_group["metric_1"],
                start_col + index,
                f"{formatted_value} {primary_unit}".strip() if formatted_value else "",
                font_size=18,
                rotate_like_hours=True,
            )

        if len(definition["metrics"]) > 1 and "metric_2" in row_group:
            second_metric = definition["metrics"][1]
            for index, value in enumerate(record[second_metric["key"]]):
                formatted_second_value = format_metric_value(payload, second_metric, value)
                write_slot_value(
                    worksheet,
                    row_group["metric_2"],
                    start_col + index,
                    f"{formatted_second_value} {second_metric['unit']}".strip() if formatted_second_value else "",
                    font_size=18,
                    rotate_like_hours=True,
                )

        for index, performed_by in enumerate(record["performed_by_slots"]):
            write_signature_or_text_slot(
                worksheet,
                row_group["performed_by"],
                start_col + index,
                performed_by,
                width=70,
                height=170,
            )
        write_signature_or_text_status(
            worksheet,
            row_group["verified_by"],
            start_col,
            record["verified_by"],
            width=120,
            height=42,
        )
        write_day_status(
            worksheet,
            row_group["date"],
            start_col,
            get_effective_record_date(record),
        )

    write_observations_cell(worksheet, footer_map["observations"], compose_observations_export_text(payload))
    write_signature_or_text_cell(
        worksheet,
        footer_map["reviewed_by"],
        payload["monthly_closure"]["reviewed_by"],
        width=180,
        height=52,
    )
    write_template_cell(
        worksheet,
        footer_map["reviewed_on"],
        normalize_excel_date(payload["monthly_closure"]["reviewed_on"]),
    )

    for index, sheet_name in enumerate(workbook.sheetnames):
        sheet = workbook[sheet_name]
        if sheet_name == target_sheet_name:
            sheet.sheet_state = "visible"
            workbook.active = index
        else:
            sheet.sheet_state = "hidden"

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def load_template_workbook(form_key: str, data_only: bool = False):
    return load_workbook(BytesIO(get_template_bytes(form_key)), data_only=data_only)


def resolve_writable_coordinate(worksheet: Any, coordinate: str) -> str:
    for merged_range in worksheet.merged_cells.ranges:
        if coordinate in merged_range:
            return worksheet.cell(merged_range.min_row, merged_range.min_col).coordinate
    return coordinate


def write_template_cell(worksheet: Any, coordinate: str, value: Any) -> None:
    worksheet[resolve_writable_coordinate(worksheet, coordinate)] = value


def write_observations_cell(worksheet: Any, coordinate: str, value: str) -> None:
    writable_coordinate = resolve_writable_coordinate(worksheet, coordinate)
    cell = worksheet[writable_coordinate]
    cell.value = value
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    merged_range = get_merged_range_for_coordinate(worksheet, writable_coordinate)
    if merged_range is None:
        min_col = max_col = cell.column
        min_row = max_row = cell.row
    else:
        min_col = merged_range.min_col
        max_col = merged_range.max_col
        min_row = merged_range.min_row
        max_row = merged_range.max_row

    total_width_pixels = 0.0
    for column in range(min_col, max_col + 1):
        total_width_pixels += column_width_to_pixels(
            worksheet.column_dimensions[get_column_letter(column)].width
        )

    approx_chars_per_line = max(int(total_width_pixels / 18), 45)
    content_lines = str(value or "").splitlines() or [""]
    estimated_lines = 0
    for line in content_lines:
        width_based_lines = max(1, (len(line) + approx_chars_per_line - 1) // approx_chars_per_line)
        conservative_lines = max(1, (len(line) + 79) // 80)
        estimated_lines += max(width_based_lines, conservative_lines)

    total_rows = max_row - min_row + 1
    total_height_points = max(60.0, estimated_lines * 18.0)
    row_height_points = total_height_points / total_rows
    for row in range(min_row, max_row + 1):
        current_height = worksheet.row_dimensions[row].height
        worksheet.row_dimensions[row].height = max(current_height or 0, row_height_points)


def normalize_signature_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_text = "".join(char for char in normalized if not unicodedata.combining(char))
    ascii_text = re.sub(r"[_\-.,/]+", " ", ascii_text.lower())
    ascii_text = re.sub(r"\s+", " ", ascii_text).strip()
    return ascii_text


def strip_signature_suffix(file_stem: str) -> str:
    cleaned = re.sub(r"_(azul|negro|firma)$", "", file_stem, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+\d+$", "", cleaned).strip()
    return cleaned


def build_signature_aliases(tokens: list[str], normalized_name: str) -> dict[str, int]:
    aliases: dict[str, int] = {normalized_name: 1000}
    collapsed_name = normalized_name.replace(" ", "")
    if collapsed_name:
        aliases[collapsed_name] = 980

    for token in tokens:
        aliases.setdefault(token, 120)
        aliases.setdefault(f"{token[0]} {token}", 890)
        aliases.setdefault(f"{token[0]}.{token}", 890)

    for index, token in enumerate(tokens):
        for other_index, other_token in enumerate(tokens):
            if index == other_index:
                continue

            aliases.setdefault(f"{token} {other_token}", 340)
            aliases.setdefault(f"{other_token} {token}", 330)

            initial = other_token[0]
            aliases.setdefault(f"{initial} {token}", 900)
            aliases.setdefault(f"{initial}.{token}", 900)
            aliases.setdefault(f"{token} {initial}", 860)
            aliases.setdefault(f"{token}.{initial}", 860)

            aliases.setdefault(f"{other_token} {token}", 780)
            aliases.setdefault(f"{token} {other_token}", 780)

    if len(tokens) >= 2:
        primary_surname = tokens[0]
        for given_token in tokens[1:]:
            initial = given_token[0]
            aliases[f"{initial} {primary_surname}"] = max(aliases.get(f"{initial} {primary_surname}", 0), 950)
            aliases[f"{initial}.{primary_surname}"] = max(aliases.get(f"{initial}.{primary_surname}", 0), 950)
            aliases[f"{given_token} {primary_surname}"] = max(aliases.get(f"{given_token} {primary_surname}", 0), 920)

    return aliases


def get_local_signature_cache_key() -> tuple[str, ...]:
    if not SIGNATURES_DIR.exists():
        return ()
    return tuple(sorted(path.name.lower() for path in SIGNATURES_DIR.glob("*.png")))


@st.cache_data(show_spinner=False, ttl=300)
def _load_signature_catalog_cached(
    _storage_cache_key: tuple[bool, str, str],
    _local_cache_key: tuple[str, ...],
) -> list[dict[str, Any]]:
    catalog: list[dict[str, Any]] = []
    remote_assets: list[dict[str, str]] = []
    if signatures_storage_enabled():
        try:
            remote_assets = sorted(list_remote_signature_assets(), key=lambda asset: asset["name"].lower())
        except Exception:
            remote_assets = []

    for asset in remote_assets:
        asset_name = asset["name"]
        display_name = strip_signature_suffix(Path(asset_name).stem)
        normalized_name = normalize_signature_text(display_name)
        tokens = normalized_name.split()
        catalog.append(
            {
                "asset_name": asset_name,
                "storage_path": asset["path"],
                "local_path": None,
                "display_name": display_name,
                "normalized_name": normalized_name,
                "tokens": tokens,
                "aliases": build_signature_aliases(tokens, normalized_name),
            }
        )

    if SIGNATURES_DIR.exists():
        existing_names = {candidate["asset_name"].lower() for candidate in catalog}
        for path in sorted(SIGNATURES_DIR.glob("*.png")):
            if path.name.lower() in existing_names:
                continue
            display_name = strip_signature_suffix(path.stem)
            normalized_name = normalize_signature_text(display_name)
            tokens = normalized_name.split()
            catalog.append(
                {
                    "asset_name": path.name,
                    "storage_path": None,
                    "local_path": path,
                    "display_name": display_name,
                    "normalized_name": normalized_name,
                    "tokens": tokens,
                    "aliases": build_signature_aliases(tokens, normalized_name),
                }
            )
    return catalog


def load_signature_catalog() -> list[dict[str, Any]]:
    return _load_signature_catalog_cached(
        get_signatures_storage_cache_key(),
        get_local_signature_cache_key(),
    )


def find_signature_candidate(person_name: str) -> dict[str, Any] | None:
    normalized_input = normalize_signature_text(person_name)
    if not normalized_input:
        return None

    catalog = load_signature_catalog()
    if not catalog:
        return None

    raw_input = unicodedata.normalize("NFKD", person_name)
    ascii_input = "".join(char for char in raw_input if not unicodedata.combining(char))
    dotted_match = re.search(r"\.\s*([A-Za-z]+)", ascii_input)
    if dotted_match:
        surname_after_dot = normalize_signature_text(dotted_match.group(1))
        surname_matches = [
            candidate
            for candidate in catalog
            if surname_after_dot and surname_after_dot in candidate["tokens"]
        ]
        unique_surname_matches = {candidate["asset_name"].lower(): candidate for candidate in surname_matches}
        if len(unique_surname_matches) == 1:
            return next(iter(unique_surname_matches.values()))

    best_score = -1
    best_matches: list[dict[str, Any]] = []
    collapsed_input = normalized_input.replace(" ", "")
    input_tokens = normalized_input.split()

    for candidate in catalog:
        score = 0
        aliases = candidate["aliases"]
        if normalized_input in aliases:
            score = max(score, aliases[normalized_input])
        if collapsed_input in aliases:
            score = max(score, aliases[collapsed_input])
        if collapsed_input and collapsed_input == candidate["normalized_name"].replace(" ", ""):
            score = max(score, 970)
        if input_tokens and all(token in candidate["tokens"] for token in input_tokens):
            score = max(score, 400 + (len(input_tokens) * 10))
        if len(input_tokens) == 1 and input_tokens[0] in candidate["tokens"]:
            score = max(score, 180)

        if score > best_score:
            best_score = score
            best_matches = [candidate]
        elif score > 0 and score == best_score:
            best_matches.append(candidate)

    unique_matches = {candidate["asset_name"].lower(): candidate for candidate in best_matches}
    if best_score <= 0 or len(unique_matches) != 1:
        return None
    return next(iter(unique_matches.values()))


def get_signature_display_name(person_name: str) -> str | None:
    signature_candidate = find_signature_candidate(person_name)
    if signature_candidate is None:
        return None
    return str(signature_candidate["display_name"])


def get_signature_image_source(signature_candidate: dict[str, Any]) -> BytesIO | Path | None:
    storage_path = signature_candidate.get("storage_path")
    if storage_path:
        try:
            signature_bytes = download_signature_bytes(str(storage_path))
            image_buffer = BytesIO(signature_bytes)
            image_buffer.seek(0)
            return image_buffer
        except Exception:
            pass

    local_path = signature_candidate.get("local_path")
    if isinstance(local_path, Path) and local_path.exists():
        return local_path
    return None


def add_signature_image(
    worksheet: Any,
    coordinate: str,
    image_source: BytesIO | Path,
    width: int,
    height: int,
    rotate_vertical: bool = False,
) -> None:
    writable_coordinate = resolve_writable_coordinate(worksheet, coordinate)
    worksheet[writable_coordinate] = None
    image_buffer = prepare_signature_image_buffer(image_source, rotate_vertical=rotate_vertical)
    image = XLImage(image_buffer)
    fitted_width, fitted_height = fit_signature_dimensions(
        worksheet,
        writable_coordinate,
        original_width=image.width,
        original_height=image.height,
        requested_width=width,
        requested_height=height,
    )
    image.width = fitted_width
    image.height = fitted_height
    image.anchor = build_signature_anchor(worksheet, writable_coordinate, fitted_width, fitted_height)
    worksheet.add_image(image)


def prepare_signature_image_buffer(image_source: BytesIO | Path, rotate_vertical: bool = False) -> BytesIO:
    if isinstance(image_source, Path):
        image = PILImage.open(image_source).convert("RGBA")
    else:
        image_source.seek(0)
        image = PILImage.open(image_source).convert("RGBA")
    alpha_bbox = image.getchannel("A").getbbox() if "A" in image.getbands() else None
    if alpha_bbox is not None:
        image = image.crop(alpha_bbox)
    if rotate_vertical:
        image = image.rotate(90, expand=True)
    image_buffer = BytesIO()
    image.save(image_buffer, format="PNG")
    image_buffer.seek(0)
    return image_buffer


def get_merged_range_for_coordinate(worksheet: Any, coordinate: str) -> Any | None:
    for merged_range in worksheet.merged_cells.ranges:
        if coordinate in merged_range:
            return merged_range
    return None


def column_width_to_pixels(width: float | None) -> float:
    effective_width = 8.43 if width in (None, 0) else float(width)
    return max((effective_width * 7) + 5, 24)


def row_height_to_pixels(height: float | None) -> float:
    effective_height = 15 if height in (None, 0) else float(height)
    return max(effective_height * (96 / 72), 20)


def get_signature_bounds(worksheet: Any, coordinate: str) -> tuple[float, float]:
    merged_range = get_merged_range_for_coordinate(worksheet, coordinate)
    if merged_range is None:
        cell = worksheet[coordinate]
        column_width = worksheet.column_dimensions[get_column_letter(cell.column)].width
        row_height = worksheet.row_dimensions[cell.row].height
        return column_width_to_pixels(column_width), row_height_to_pixels(row_height)

    total_width = 0.0
    for column in range(merged_range.min_col, merged_range.max_col + 1):
        column_letter = get_column_letter(column)
        total_width += column_width_to_pixels(worksheet.column_dimensions[column_letter].width)

    total_height = 0.0
    for row in range(merged_range.min_row, merged_range.max_row + 1):
        total_height += row_height_to_pixels(worksheet.row_dimensions[row].height)

    return total_width, total_height


def get_signature_anchor_origin(worksheet: Any, coordinate: str) -> tuple[int, int]:
    merged_range = get_merged_range_for_coordinate(worksheet, coordinate)
    if merged_range is None:
        cell = worksheet[coordinate]
        return cell.column, cell.row
    return merged_range.min_col, merged_range.min_row


def fit_signature_dimensions(
    worksheet: Any,
    coordinate: str,
    original_width: float,
    original_height: float,
    requested_width: int,
    requested_height: int,
) -> tuple[int, int]:
    max_width, max_height = get_signature_bounds(worksheet, coordinate)
    max_width = max(min(max_width - 8, requested_width), 24)
    max_height = max(min(max_height - 4, requested_height), 16)

    if original_width <= 0 or original_height <= 0:
        return int(max_width), int(max_height)

    scale = min(max_width / original_width, max_height / original_height)
    scale = min(scale, 1.0)
    return max(int(original_width * scale), 1), max(int(original_height * scale), 1)


def build_signature_anchor(
    worksheet: Any,
    coordinate: str,
    width: int,
    height: int,
) -> OneCellAnchor:
    anchor_col, anchor_row = get_signature_anchor_origin(worksheet, coordinate)
    bounds_width, bounds_height = get_signature_bounds(worksheet, coordinate)
    offset_x = max(int((bounds_width - width) / 2), 0)
    offset_y = max(int((bounds_height - height) / 2), 0)
    marker = AnchorMarker(
        col=anchor_col - 1,
        colOff=pixels_to_EMU(offset_x),
        row=anchor_row - 1,
        rowOff=pixels_to_EMU(offset_y),
    )
    ext = XDRPositiveSize2D(pixels_to_EMU(width), pixels_to_EMU(height))
    return OneCellAnchor(_from=marker, ext=ext)


def write_signature_or_text_cell(
    worksheet: Any,
    coordinate: str,
    value: str,
    width: int,
    height: int,
) -> None:
    signature_candidate = find_signature_candidate(value)
    if signature_candidate is not None:
        image_source = get_signature_image_source(signature_candidate)
        if image_source is not None:
            add_signature_image(worksheet, coordinate, image_source, width=width, height=height)
            return
    write_template_cell(worksheet, coordinate, value)


def write_signature_or_text_slot(
    worksheet: Any,
    row: int,
    column: int,
    value: str,
    width: int,
    height: int,
) -> None:
    coordinate = worksheet.cell(row=row, column=column).coordinate
    signature_candidate = find_signature_candidate(value)
    if signature_candidate is not None:
        image_source = get_signature_image_source(signature_candidate)
        if image_source is not None:
            add_signature_image(
                worksheet,
                coordinate,
                image_source,
                width=width,
                height=height,
                rotate_vertical=True,
            )
            return
    write_slot_value(
        worksheet,
        row,
        column,
        value,
        font_size=18,
        rotate_like_hours=True,
    )


def write_signature_or_text_status(
    worksheet: Any,
    row: int,
    start_col: int,
    value: str,
    width: int,
    height: int,
) -> None:
    coordinate = worksheet.cell(row=row, column=start_col).coordinate
    signature_candidate = find_signature_candidate(value)
    if signature_candidate is not None:
        image_source = get_signature_image_source(signature_candidate)
        if image_source is not None:
            add_signature_image(worksheet, coordinate, image_source, width=width, height=height)
            return
    write_day_status(worksheet, row, start_col, value)


def write_day_status(worksheet: Any, row: int, start_col: int, value: str) -> None:
    cell = worksheet.cell(row=row, column=start_col)
    cell.value = value
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    cell.font = Font(size=20)


def write_slot_value(
    worksheet: Any,
    row: int,
    column: int,
    value: str,
    font_size: int = 9,
    rotate_like_hours: bool = False,
) -> None:
    cell = worksheet.cell(row=row, column=column)
    cell.value = value
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        text_rotation=90 if rotate_like_hours else 0,
        wrap_text=False,
    )
    cell.font = Font(size=font_size)


def ensure_status_row_merges(worksheet: Any, payload: dict[str, Any]) -> None:
    rows_to_merge = get_form_definition(payload["metadata"]["form_key"])["layout"]["status_rows_to_merge"]
    for day in range(1, 32):
        start_col = DAY_BLOCK_START_COLUMNS[day]
        for row in rows_to_merge:
            cell_range = f"{worksheet.cell(row=row, column=start_col).coordinate}:{worksheet.cell(row=row, column=start_col + 2).coordinate}"
            if not is_range_already_merged(worksheet, cell_range):
                worksheet.merge_cells(cell_range)


def is_range_already_merged(worksheet: Any, cell_range: str) -> bool:
    return any(str(merged_range) == cell_range for merged_range in worksheet.merged_cells.ranges)


def calculate_corrected_temperature(
    raw_value: str,
    correction_bands: dict[str, Any],
    correction_factors: dict[str, Any],
    correction_operations: dict[str, str],
) -> str:
    measured = parse_measurement_number(raw_value)
    if measured is None:
        return ""

    factor_key = None
    band_items = list(correction_bands.items())
    for index, (key, band) in enumerate(band_items):
        min_value = float(band["min"])
        max_value = float(band["max"])
        if max_value < min_value:
            reparsed_bounds = parse_range_bounds(str(band.get("label", "")))
            if reparsed_bounds is not None:
                min_value, max_value = reparsed_bounds
        is_last = index == len(band_items) - 1
        if min_value <= measured <= max_value if is_last else min_value <= measured < max_value:
            factor_key = key
            break
    if factor_key is None or factor_key not in correction_factors:
        return f"{measured:.2f}"

    factor = float(correction_factors[factor_key])
    operation = correction_operations[factor_key]
    corrected = measured + factor if operation == "+" else measured - factor
    return f"{corrected:.2f}"


def maybe_autosave_payload(payload: dict[str, Any]) -> None:
    if not can_edit_daily_records():
        return

    last_saved_signature = st.session_state.get("last_saved_payload_signature")
    current_signature = get_payload_signature(payload)
    if last_saved_signature is None:
        remember_saved_snapshot(payload)
        return
    if current_signature == last_saved_signature:
        return

    now = time.time()
    last_attempt = float(st.session_state.get("last_autosave_attempt_at", 0.0))
    if now - last_attempt < AUTOSAVE_DEBOUNCE_SECONDS:
        return

    st.session_state["last_autosave_attempt_at"] = now
    try:
        saved_backend = save_payload(payload)
        st.session_state["last_autosave_at"] = get_local_now().strftime("%H:%M:%S")
        st.session_state["last_autosave_backend"] = saved_backend
        st.session_state["last_autosave_error"] = ""
    except Exception as exc:
        st.session_state["last_autosave_error"] = str(exc)


def render_actions(payload: dict[str, Any]) -> None:
    st.subheader("5. Guardado y exportacion")
    allow_daily_edits = can_edit_daily_records()
    allow_export = can_export_period()
    maybe_autosave_payload(payload)
    errors = validate_payload(payload)
    if errors:
        st.warning("Hay datos pendientes antes de exportar.")
        for error in errors[:8]:
            st.write(f"- {error}")
        if len(errors) > 8:
            st.write(f"- ... y {len(errors) - 8} mas.")
    else:
        st.success("La captura esta completa para exportar la plantilla.")

    autosave_error = str(st.session_state.get("last_autosave_error", "")).strip()
    autosave_at = str(st.session_state.get("last_autosave_at", "")).strip()
    autosave_backend = str(st.session_state.get("last_autosave_backend", "")).strip()
    if autosave_error:
        st.caption(f"Autoguardado con problema: {autosave_error}")
    elif autosave_at:
        autosave_label = "respaldo local" if autosave_backend == "Local JSON" else "correcto"
        st.caption(f"Autoguardado {autosave_label}: {autosave_at}")

    save_col, export_col, reset_col = st.columns(3)
    if save_col.button("Guardar borrador", use_container_width=True, disabled=not allow_daily_edits):
        try:
            saved_backend = save_payload(payload)
            log_activity(
                "guardar_borrador",
                "Respaldo local" if saved_backend == "Local JSON" else "Guardado principal",
                payload,
            )
            if saved_backend == "Local JSON":
                st.warning("Se guardo un respaldo local del borrador.")
            else:
                st.success("Se guardo el borrador.")
        except Exception as exc:
            st.error(f"No se pudo guardar el borrador: {exc}")

    if export_col.button("Preparar Excel", use_container_width=True, disabled=not allow_export):
        if errors:
            st.error("Completa los campos pendientes antes de exportar.")
            return
        try:
            save_payload(payload)
        except Exception:
            pass
        excel_file = populate_template(payload)
        log_activity("exportar_excel", "Genero el archivo Excel", payload)
        form_definition = get_form_definition(payload["metadata"]["form_key"])
        form_code = form_definition["source_file"].split()[0].replace(".xlsx", "")
        filename = (
            f"{form_code}_{payload['metadata']['equipment_code']}_{payload['metadata']['year']}_{payload['metadata']['month']:02d}.xlsx"
        )
        st.download_button(
            "Descargar formato llenado",
            data=excel_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if reset_col.button("Limpiar periodo", use_container_width=True, disabled=not allow_daily_edits):
        log_activity("limpiar_periodo", "Restablecio el periodo actual", payload)
        current_period_key = get_period_key(payload)
        clear_period_widget_state(current_period_key)
        st.session_state.payload = build_default_payload(
            equipment_code=str(payload["metadata"]["equipment_code"]),
            form_key=str(payload["metadata"]["form_key"]),
        )
        st.session_state.payload["metadata"]["month"] = int(payload["metadata"]["month"])
        st.session_state.payload["metadata"]["year"] = int(payload["metadata"]["year"])
        st.session_state.period_key = get_period_key(st.session_state.payload)
        st.rerun()

    if not allow_export:
        st.caption("Solo responsable, calidad o admin pueden cerrar y exportar el formato.")


def main() -> None:
    form_keys = list(FORM_DEFINITIONS.keys())

    st.set_page_config(
        page_title="Formularios Digitales",
        layout="wide",
    )
    configure_users_backend()
    configure_storage_backend()
    initialize_auth_state()

    if not st.session_state["autenticado"]:
        render_auth_screen()
        st.stop()

    for form_key in form_keys:
        definition = get_form_definition(form_key)
        if not template_source_available(form_key):
            st.error(f"No se encontro la plantilla {definition['source_file']}.")
            st.stop()

    if "payload" not in st.session_state:
        st.session_state.payload = load_saved_payload(
            form_key=DEFAULT_FORM_KEY,
            equipment_code=FORM_DEFINITIONS[DEFAULT_FORM_KEY]["default_equipment"],
        )
        remember_saved_snapshot(st.session_state.payload)

    payload = st.session_state.payload
    previous_period_key = st.session_state.get("period_key", get_period_key(payload))
    equipment_configs = load_equipment_configs(payload["metadata"]["form_key"])
    equipment_codes = list(equipment_configs.keys())

    st.title("Digitalizacion de formatos")
    st.caption(
        f"Captura guiada y exportacion automatica para {payload['metadata']['form_label']} en {payload['metadata']['equipment_code']}."
    )

    selected_form_key, selected_equipment = render_sidebar(payload, form_keys, equipment_codes)
    if selected_form_key != payload["metadata"]["form_key"]:
        target_equipment = FORM_DEFINITIONS[selected_form_key]["default_equipment"]
        st.session_state.payload = load_saved_payload(
            form_key=selected_form_key,
            equipment_code=target_equipment,
            year=int(payload["metadata"]["year"]),
            month=int(payload["metadata"]["month"]),
        )
        remember_saved_snapshot(st.session_state.payload)
        st.session_state.period_key = get_period_key(st.session_state.payload)
        st.rerun()

    if selected_equipment != payload["metadata"]["equipment_code"]:
        target_period_key = (
            f"{payload['metadata']['form_key']}_{selected_equipment}_{int(payload['metadata']['year'])}_{int(payload['metadata']['month']):02d}"
        )
        clear_period_widget_state(target_period_key)
        st.session_state.payload = load_saved_payload(
            form_key=str(payload["metadata"]["form_key"]),
            equipment_code=selected_equipment,
            year=int(payload["metadata"]["year"]),
            month=int(payload["metadata"]["month"]),
        )
        remember_saved_snapshot(st.session_state.payload)
        st.session_state.period_key = get_period_key(st.session_state.payload)
        st.rerun()

    payload = st.session_state.payload
    render_configuration(payload)

    current_period_key = get_period_key(payload)
    if current_period_key != previous_period_key:
        clear_period_widget_state(current_period_key)
        st.session_state.payload = load_saved_payload(
            form_key=str(payload["metadata"]["form_key"]),
            equipment_code=str(payload["metadata"]["equipment_code"]),
            year=int(payload["metadata"]["year"]),
            month=int(payload["metadata"]["month"]),
        )
        remember_saved_snapshot(st.session_state.payload)
        st.session_state.period_key = current_period_key
        st.rerun()

    st.session_state.period_key = current_period_key
    payload = st.session_state.payload

    render_non_working_days(payload)
    render_daily_capture(payload)
    render_monthly_closure(payload)
    render_actions(payload)


if __name__ == "__main__":
    main()
