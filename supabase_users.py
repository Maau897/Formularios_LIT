from __future__ import annotations

from dataclasses import dataclass

import bcrypt

try:
    from supabase import Client, create_client
    SUPABASE_SDK_AVAILABLE = True
except ImportError:  # pragma: no cover
    Client = object  # type: ignore[assignment]
    create_client = None
    SUPABASE_SDK_AVAILABLE = False


@dataclass
class SupabaseUsersConfig:
    url: str
    key: str
    enabled: bool = False
    table_name: str = "usuarios_app"


_CONFIG = SupabaseUsersConfig(url="", key="", enabled=False)
_CLIENT: Client | None = None


def configure_supabase_users(
    *,
    url: str | None,
    key: str | None,
    enabled: bool = False,
    table_name: str = "usuarios_app",
) -> None:
    global _CONFIG, _CLIENT
    _CONFIG = SupabaseUsersConfig(
        url=(url or "").strip(),
        key=(key or "").strip(),
        enabled=bool(enabled and url and key),
        table_name=table_name,
    )
    _CLIENT = None


def supabase_users_enabled() -> bool:
    return bool(SUPABASE_SDK_AVAILABLE and _CONFIG.enabled and _CONFIG.url and _CONFIG.key)


def get_users_backend_label() -> str:
    return "Supabase" if supabase_users_enabled() else "No configurado"


def _client() -> Client:
    global _CLIENT
    if not supabase_users_enabled():
        raise RuntimeError("Supabase de usuarios no esta configurado.")
    if _CLIENT is None:
        _CLIENT = create_client(_CONFIG.url, _CONFIG.key)
    return _CLIENT


def _table():
    return _client().table(_CONFIG.table_name)


def _hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def _verify_password(password: str, password_hash: str) -> bool:
    return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))


def registrar_usuario(email: str, password: str, rol: str) -> None:
    _table().insert(
        {
            "email": email.strip().lower(),
            "password_hash": _hash_password(password),
            "aprobado": False,
            "es_admin": False,
            "rol": rol,
        }
    ).execute()


def autenticar_usuario(email: str, password: str, normalizar_rol_fn):
    response = _table().select("*").eq("email", email.strip().lower()).limit(1).execute()
    rows = response.data or []
    if not rows:
        return {"ok": False, "mensaje": "Usuario no encontrado."}

    usuario = rows[0]
    if not _verify_password(password, usuario["password_hash"]):
        return {"ok": False, "mensaje": "Contrasena incorrecta."}

    if not usuario.get("aprobado", False):
        return {"ok": False, "mensaje": "Tu cuenta aun no ha sido aprobada."}

    es_admin = bool(usuario.get("es_admin", False))
    return {
        "ok": True,
        "id_usuario": usuario["id_usuario"],
        "email": usuario["email"],
        "es_admin": es_admin,
        "rol": normalizar_rol_fn(usuario.get("rol"), es_admin),
    }


def obtener_usuarios_pendientes() -> list[tuple]:
    rows = (
        _table()
        .select("id_usuario,email,fecha_registro")
        .eq("aprobado", False)
        .order("fecha_registro")
        .execute()
        .data
        or []
    )
    return [(row["id_usuario"], row["email"], row["fecha_registro"]) for row in rows]


def listar_usuarios() -> list[tuple]:
    rows = (
        _table()
        .select("id_usuario,email,aprobado,es_admin,rol,fecha_registro")
        .order("es_admin", desc=True)
        .order("aprobado", desc=True)
        .order("email")
        .execute()
        .data
        or []
    )
    return [
        (
            row["id_usuario"],
            row["email"],
            1 if row.get("aprobado", False) else 0,
            1 if row.get("es_admin", False) else 0,
            row.get("rol") or "captura",
            row.get("fecha_registro"),
        )
        for row in rows
    ]


def aprobar_usuario(id_usuario: int, rol: str) -> None:
    es_admin = rol == "admin"
    _table().update({"aprobado": True, "rol": rol, "es_admin": es_admin}).eq("id_usuario", id_usuario).execute()


def actualizar_rol_usuario(id_usuario: int, rol: str) -> None:
    es_admin = rol == "admin"
    _table().update({"rol": rol, "es_admin": es_admin}).eq("id_usuario", id_usuario).execute()


def crear_admin_inicial(email: str, password: str) -> None:
    correo = email.strip().lower()
    existente = (_table().select("id_usuario").eq("email", correo).limit(1).execute().data or [])
    if existente:
        _table().update({"aprobado": True, "es_admin": True, "rol": "admin"}).eq("email", correo).execute()
        return

    _table().insert(
        {
            "email": correo,
            "password_hash": _hash_password(password),
            "aprobado": True,
            "es_admin": True,
            "rol": "admin",
        }
    ).execute()
