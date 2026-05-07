# Formularios

App en Streamlit para digitalizar formatos `F-LIT`, con exportacion automatica a Excel, firmas digitales y acceso por usuarios.

## Alcance actual

- Captura guiada de multiples familias `F-LIT`.
- Dias no laborados marcados manualmente.
- Factores de correccion editables por mes y equipo.
- Exportacion automatica a una copia de la plantilla oficial de Excel.
- Firmas digitales para `Realizo`, `Verifico` y `Reviso`.
- Acceso con usuarios compartidos desde Supabase.
- Persistencia de borradores y periodos mensuales en Supabase.
- Soporte para firmas digitales desde Supabase Storage con fallback local.

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
USE_SUPABASE_STORAGE=true
SUPABASE_USERS_TABLE=usuarios_app
SUPABASE_STORAGE_TABLE=formularios_periodos
SUPABASE_SIGNATURES_BUCKET=firmas-digitales
SUPABASE_SIGNATURES_PREFIX=
ADMIN_EMAIL=tu-admin@ejemplo.com
ADMIN_PASSWORD=tu-password
```

## Nota

La app trabaja sobre copias de las plantillas y genera un archivo nuevo con los datos capturados, sin modificar la plantilla base.

Antes de usar la persistencia remota, ejecutar el esquema de [supabase_schema.sql](C:/Users/mauri/OneDrive/Desktop/Formularios/supabase_schema.sql) en el proyecto de Supabase.

## Firmas digitales en despliegue

En produccion, las firmas deben subirse a un bucket de Supabase Storage, por defecto `firmas-digitales`.

- Subir los archivos `.png` con el mismo nombre que ya usas localmente.
- Si quieres guardarlos dentro de una carpeta del bucket, configurar `SUPABASE_SIGNATURES_PREFIX`.
- La app intenta resolver primero desde Supabase Storage y, si no encuentra el archivo o falla la descarga, usa la carpeta local `firmas digitales/` como fallback.
