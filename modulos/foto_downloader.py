"""
foto_downloader.py
==================
Módulo para descargar fotos desde Google Sheets (Google Forms).
Busca una matrícula en la columna H, obtiene el nombre de la columna I
y descarga la foto de la columna K con ese nombre.

Autenticación: cuenta de servicio (archivo JSON) — más confiable que API Key.
Lectura dinámica de IDs de formularios desde config.json.
"""
import os
import re
import sys
import json
import requests
from pathlib import Path

# ─────────────────────────────────────────
# RUTA BASE — compatible con PyInstaller
# ─────────────────────────────────────────

def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    # IMPORTANTE: Le agregamos un .parent extra. 
    # Como el archivo está en la carpeta "modulos", esto lo hace subir a la carpeta principal.
    return Path(__file__).resolve().parent.parent

BASE_DIR = get_base_dir()

# ─────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────

# Archivo JSON de configuración donde están los IDs de los formularios
CONFIG_FILE = BASE_DIR / "config.json"

# Ruta al archivo JSON de la cuenta de servicio
CREDENTIALS_JSON = BASE_DIR / "educacion-continua-490016-7ff6875c1d8f.json"

SHEET_NAME = "Respuestas de formulario 1"  # cambia si tu hoja tiene otro nombre
FILA_INICIO = 2

# Columnas (letras)
COL_MATRICULA = "H"
COL_NOMBRE    = "I"
COL_FOTO      = "K"

# Carpeta de descarga por defecto
CARPETA_DEFAULT = str(Path.home() / "Downloads")

# Scopes necesarios
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ─────────────────────────────────────────
# AUTENTICACIÓN Y LECTURA DE CONFIG
# ─────────────────────────────────────────

def _get_form_ids() -> list:
    """Obtiene la lista de IDs de formularios desde el archivo config.json."""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                return config.get("forms_ids", [])
        except Exception as e:
            print(f"  ❌ Error leyendo config.json: {e}")
    return []

def _get_credentials():
    """Carga las credenciales desde el archivo JSON de cuenta de servicio."""
    from google.oauth2 import service_account
    if not CREDENTIALS_JSON.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de credenciales: {CREDENTIALS_JSON}\n"
            f"Coloca tu archivo JSON de cuenta de servicio en: {BASE_DIR}"
        )
    return service_account.Credentials.from_service_account_file(
        str(CREDENTIALS_JSON), scopes=SCOPES
    )

def _get_sheets_service():
    """Retorna el cliente autenticado de Google Sheets."""
    from googleapiclient.discovery import build
    creds = _get_credentials()
    return build("sheets", "v4", credentials=creds)

def _get_drive_service():
    """Retorna el cliente autenticado de Google Drive."""
    from googleapiclient.discovery import build
    creds = _get_credentials()
    return build("drive", "v3", credentials=creds)

# ─────────────────────────────────────────
# FUNCIONES INTERNAS
# ─────────────────────────────────────────

def _obtener_rango(sheet_id: str, rango: str) -> list:
    """Consulta un rango de un Google Sheets específico y retorna los valores."""
    service = _get_sheets_service()
    resultado = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=f"{SHEET_NAME}!{rango}")
        .execute()
    )
    return resultado.get("values", [])

def _extraer_id_drive(link: str) -> str | None:
    """Extrae el ID del archivo de un link de Google Drive."""
    patrones = [
        r"(?:id=)([a-zA-Z0-9_-]+)",
        r"(?:/d/)([a-zA-Z0-9_-]+)",
    ]
    for patron in patrones:
        match = re.search(patron, link)
        if match:
            return match.group(1)
    return None

def _descargar_drive(file_id: str, ruta_destino: str) -> bool:
    """Descarga un archivo de Google Drive usando la cuenta de servicio."""
    try:
        from googleapiclient.http import MediaIoBaseDownload
        import io

        service = _get_drive_service()
        request = service.files().get_media(fileId=file_id)

        with open(ruta_destino, "wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        return True

    except Exception as e:
        # Fallback: descarga directa por requests si falla la API
        print(f"  ⚠️ API Drive falló ({e}), intentando descarga directa...")
        return _descargar_drive_directo(file_id, ruta_destino)

def _descargar_drive_directo(file_id: str, ruta_destino: str) -> bool:
    """Descarga directa desde Google Drive sin autenticación (fallback)."""
    url    = f"https://drive.google.com/uc?export=download&id={file_id}"
    sesion = requests.Session()
    resp   = sesion.get(url, stream=True, timeout=30)

    # Confirmación para archivos grandes
    for key, value in resp.cookies.items():
        if "download_warning" in key:
            resp = sesion.get(f"{url}&confirm={value}", stream=True, timeout=30)
            break

    if resp.status_code != 200:
        return False

    with open(ruta_destino, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
    return True

# ─────────────────────────────────────────
# FUNCIÓN PRINCIPAL DEL MÓDULO
# ─────────────────────────────────────────

def descargar_foto(matricula, carpeta_destino: str = CARPETA_DEFAULT) -> tuple:
    """
    Busca TODAS las filas con esa matrícula en la columna H a través de 
    TODOS los formularios registrados en config.json, toma la MÁS RECIENTE 
    y descarga su foto.
    """
    matricula_str = str(matricula).strip()
    form_ids = _get_form_ids()

    if not form_ids:
        print("  ❌ No se encontraron IDs de formularios en config.json")
        return None, None

    filas_encontradas = []

    # Iterar por todos los formularios configurados
    for sheet_id in form_ids:
        try:
            valores = _obtener_rango(sheet_id, f"{COL_MATRICULA}{FILA_INICIO}:{COL_FOTO}5000")
            for fila in valores:
                if len(fila) > 0 and str(fila[0]).strip() == matricula_str:
                    filas_encontradas.append(fila)
        except Exception as e:
            print(f"  ❌ Error leyendo Sheet ID {sheet_id}: {e}")

    if not filas_encontradas:
        print(f"  ⚠️ Matrícula {matricula_str} no encontrada en ningún formulario")
        return None, None

    # Tomar la más reciente (última en la lista general)
    fila_encontrada = filas_encontradas[-1]
    print(f"  📋 {len(filas_encontradas)} respuesta(s) encontrada(s) en total — usando la más reciente")

    # Obtener nombre (columna I = índice 1 del rango H:K)
    if len(fila_encontrada) < 2 or not fila_encontrada[1]:
        print(f"  ⚠️ Sin nombre para matrícula {matricula_str}")
        return None, None
    nombre = str(fila_encontrada[1]).strip()

    # Obtener link de foto (columna K = índice 3 del rango H:K)
    if len(fila_encontrada) < 4 or not fila_encontrada[3]:
        print(f"  ⚠️ Sin foto para {nombre}")
        return None, None
    link_foto = str(fila_encontrada[3]).strip()

    # Extraer ID de Google Drive
    file_id = _extraer_id_drive(link_foto)
    if not file_id:
        print(f"  ❌ No se pudo extraer el ID del link: {link_foto}")
        return None, None

    os.makedirs(carpeta_destino, exist_ok=True)
    nombre_archivo = f"{nombre}.jpg"
    ruta_destino   = os.path.join(carpeta_destino, nombre_archivo)

    print(f"  📥 Descargando foto de {nombre} (respuesta más reciente)...")
    exito = _descargar_drive(file_id, ruta_destino)

    if exito:
        print(f"  ✅ Foto guardada: {ruta_destino}")
        return ruta_destino, nombre
    else:
        print(f"  ❌ Error descargando foto de {nombre}")
        return None, None

# ─────────────────────────────────────────
# EJECUCIÓN DIRECTA (prueba)
# ─────────────────────────────────────────

if __name__ == "__main__":
    # Prueba manual. Reemplaza con una matrícula que sepas que existe.
    resultado = descargar_foto(
        matricula="1234567", 
        carpeta_destino=str(Path.home() / "Downloads")
    )
    if resultado:
        print(f"\n✅ Descarga completada: {resultado}")
    else:
        print("\n❌ No se pudo descargar la foto")