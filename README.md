# Formularios

App en Streamlit para digitalizar formatos `F-LIT`, con exportacion automatica a Excel, firmas digitales y acceso por usuarios.

## Alcance actual

- Captura guiada de multiples familias `F-LIT`.
- Dias no laborados marcados manualmente.
- Factores de correccion editables por mes y equipo.
- Exportacion automatica a una copia de la plantilla oficial de Excel.
- Firmas digitales para `Realizo`, `Verifico` y `Reviso`.
- Acceso con usuarios compartidos desde Supabase.

## Puesta en marcha

1. Crear un entorno virtual.
2. Instalar dependencias:

```bash
pip install -r requirements.txt
```

3. Configurar `.streamlit/secrets.toml` a partir de `.streamlit/secrets.toml.example` o usar variables de entorno.
4. Ejecutar la app:

```bash
streamlit run app.py
```

## Configuracion de acceso

Para reutilizar la misma base de usuarios de `iner_db`, apuntar esta app a la tabla `usuarios_app`.

Tambien se pueden usar variables de entorno:

```env
SUPABASE_URL=https://tu-proyecto.supabase.co
SUPABASE_KEY=tu-clave
USE_SUPABASE_USERS=true
SUPABASE_USERS_TABLE=usuarios_app
ADMIN_EMAIL=tu-admin@ejemplo.com
ADMIN_PASSWORD=tu-password
```

## Nota

La app trabaja sobre copias de las plantillas y genera un archivo nuevo con los datos capturados, sin modificar la plantilla base.
