"""
AGENTE DRIVE - Gestion de archivos en Google Drive (API)
========================================================
Guarda archivos Excel y reportes HTML en las carpetas del Shared Drive
usando la API de Google Drive. Funciona en la nube (GitHub Actions)
sin depender de que la PC este encendida.

POLITICA DE ARCHIVOS:
  - NUNCA se sobreescribe un archivo existente.
  - Si ya existe un archivo con el mismo nombre, se agrega un sufijo
    numerico al nuevo: "archivo (2).xlsx", "archivo (3).xlsx", etc.
  - Los correos ya procesados se filtran por Gmail Label, asi que
    normalmente no deberian llegar duplicados.
"""
import io
from datetime import datetime, timezone, timedelta

from googleapiclient.http import MediaIoBaseUpload

from config import OBRA_FOLDER_IDS, MONTH_NAMES, MONTH_ABBREVS

# Zona horaria Peru (UTC-5)
PERU_TZ = timezone(timedelta(hours=-5))


def guardar_excel(drive_service, excel_bytes, obra_key, month_name, filename):
    """
    Guarda el archivo Excel en la carpeta del Shared Drive.
    Si ya existe uno con el mismo nombre, renombra el nuevo con sufijo.

    Args:
        drive_service: Google Drive API service
        excel_bytes: bytes del archivo Excel
        obra_key: clave de la obra (ej: 'BEETHOVEN')
        month_name: nombre del mes (ej: 'Febrero')
        filename: nombre del archivo

    Returns:
        dict con 'file_id', 'web_link' y 'filename_final', o None si falla
    """
    parent_folder_id = OBRA_FOLDER_IDS.get(obra_key)
    if not parent_folder_id or parent_folder_id == "PEGAR_FOLDER_ID_AQUI":
        print(f"  [DRIVE] No hay Folder ID configurado para obra '{obra_key}'")
        return None

    try:
        # Buscar o crear subcarpeta del mes
        month_folder_id = _get_or_create_month_folder(drive_service, parent_folder_id, month_name)

        # Determinar mime type
        if filename.lower().endswith(".xlsm"):
            mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12"
        else:
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        # Generar nombre unico (no sobreescribir)
        final_name = _generar_nombre_unico(drive_service, month_folder_id, filename)

        # Subir archivo
        file_metadata = {
            "name": final_name,
            "parents": [month_folder_id],
        }

        media = MediaIoBaseUpload(
            io.BytesIO(excel_bytes),
            mimetype=mime_type,
            resumable=True
        )

        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()

        if final_name != filename:
            print(f"  [DRIVE] Excel guardado (renombrado): {obra_key}/{month_name}/{final_name}")
        else:
            print(f"  [DRIVE] Excel guardado: {obra_key}/{month_name}/{final_name}")

        return {
            "file_id": file["id"],
            "web_link": file.get("webViewLink", ""),
            "filename_final": final_name,
        }

    except Exception as e:
        print(f"  [DRIVE] Error guardando Excel: {e}")
        return None


def guardar_reporte(drive_service, html_content, obra_key, month_name, filename):
    """
    Guarda el reporte HTML en la carpeta del Shared Drive.
    Si ya existe uno con el mismo nombre, renombra el nuevo con sufijo.

    Args:
        drive_service: Google Drive API service
        html_content: string con el HTML completo del reporte
        obra_key: clave de la obra
        month_name: nombre del mes
        filename: nombre del archivo HTML

    Returns:
        dict con 'file_id', 'web_link' y 'filename_final', o None si falla
    """
    parent_folder_id = OBRA_FOLDER_IDS.get(obra_key)
    if not parent_folder_id or parent_folder_id == "PEGAR_FOLDER_ID_AQUI":
        print(f"  [DRIVE] No hay Folder ID configurado para obra '{obra_key}'")
        return None

    try:
        # Buscar o crear subcarpeta del mes
        month_folder_id = _get_or_create_month_folder(drive_service, parent_folder_id, month_name)

        # Generar nombre unico (no sobreescribir)
        final_name = _generar_nombre_unico(drive_service, month_folder_id, filename)

        # Crear archivo nuevo
        file_metadata = {
            "name": final_name,
            "parents": [month_folder_id],
        }

        media = MediaIoBaseUpload(
            io.BytesIO(html_content.encode("utf-8")),
            mimetype="text/html",
            resumable=True
        )

        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()

        if final_name != filename:
            print(f"  [DRIVE] Reporte guardado (renombrado): {obra_key}/{month_name}/{final_name}")
        else:
            print(f"  [DRIVE] Reporte guardado: {obra_key}/{month_name}/{final_name}")

        return {
            "file_id": file["id"],
            "web_link": file.get("webViewLink", ""),
            "filename_final": final_name,
        }

    except Exception as e:
        print(f"  [DRIVE] Error guardando reporte: {e}")
        return None


def obtener_link_archivo(file_id):
    """Genera el link de visualizacion de un archivo en Drive."""
    return f"https://drive.google.com/file/d/{file_id}/view"


def obtener_mes_actual():
    """Retorna el nombre del mes actual (hora Peru)."""
    now = datetime.now(PERU_TZ)
    return MONTH_NAMES[now.month - 1]


def obtener_fecha_actual():
    """Retorna la fecha actual (hora Peru)."""
    return datetime.now(PERU_TZ)


def verificar_folder_ids():
    """
    Verifica cuales obras tienen Folder ID configurado.
    Util para diagnostico.

    Returns:
        dict con obra -> {'configurado': bool, 'folder_id': str}
    """
    resultado = {}
    for obra_key, folder_id in OBRA_FOLDER_IDS.items():
        configurado = folder_id and folder_id != "PEGAR_FOLDER_ID_AQUI"
        resultado[obra_key] = {"configurado": configurado, "folder_id": folder_id}
    return resultado


# ============================================================
# FUNCIONES INTERNAS
# ============================================================

def _generar_nombre_unico(drive_service, folder_id, filename):
    """
    Si ya existe un archivo con el mismo nombre en la carpeta,
    genera un nombre unico agregando sufijo numerico.

    Ejemplo:
      "reporte.html"  -> si existe -> "reporte (2).html"
      "reporte (2).html" -> si existe -> "reporte (3).html"
      "datos.xlsm" -> si existe -> "datos (2).xlsm"

    Returns:
        nombre de archivo unico (puede ser el original si no existe conflicto)
    """
    # Verificar si existe el nombre original
    existing = _buscar_archivo(drive_service, folder_id, filename)
    if not existing:
        return filename  # No hay conflicto, usar nombre original

    # Separar nombre base y extension
    if "." in filename:
        dot_pos = filename.rfind(".")
        base = filename[:dot_pos]
        ext = filename[dot_pos:]  # incluye el punto: ".html", ".xlsx"
    else:
        base = filename
        ext = ""

    # Probar con sufijos (2), (3), (4)... hasta encontrar uno libre
    for n in range(2, 100):
        candidate = f"{base} ({n}){ext}"
        existing = _buscar_archivo(drive_service, folder_id, candidate)
        if not existing:
            print(f"  [DRIVE] Archivo '{filename}' ya existe, renombrando a '{candidate}'")
            return candidate

    # Fallback extremo: usar timestamp
    ts = datetime.now(PERU_TZ).strftime("%Y%m%d_%H%M%S")
    fallback = f"{base}_{ts}{ext}"
    print(f"  [DRIVE] Muchas versiones, usando timestamp: '{fallback}'")
    return fallback


def _get_or_create_month_folder(drive_service, parent_folder_id, month_name):
    """
    Busca la subcarpeta del mes dentro de la carpeta padre.
    Las carpetas existentes usan el formato: "2.Feb-26", "3.Mar-26", etc.

    Busca en este orden:
      1. Patron existente: "2.Feb-26" (numero.Abrev-anio)
      2. Nombre completo: "Febrero"
      3. Si no existe, crea con el formato existente: "2.Feb-26"

    Returns:
        folder_id de la subcarpeta del mes
    """
    now = datetime.now(PERU_TZ)
    month_idx = MONTH_NAMES.index(month_name)  # 0-based
    month_num = month_idx + 1  # 1-based
    month_abbrev = MONTH_ABBREVS[month_idx]
    year_short = str(now.year)[-2:]  # "26"

    # Nombre esperado en formato existente: "2.Feb-26"
    expected_name = f"{month_num}.{month_abbrev}-{year_short}"

    try:
        # Listar TODAS las carpetas hijas para buscar por patron
        query = (
            f"'{parent_folder_id}' in parents "
            f"and mimeType = 'application/vnd.google-apps.folder' "
            f"and trashed = false"
        )

        results = drive_service.files().list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageSize=50,
        ).execute()

        folders = results.get("files", [])

        # Busqueda 1: nombre exacto esperado "2.Feb-26"
        for f in folders:
            if f["name"] == expected_name:
                print(f"  [DRIVE] Carpeta encontrada: {f['name']}")
                return f["id"]

        # Busqueda 2: contiene la abreviatura del mes (ej: "Feb")
        for f in folders:
            if month_abbrev.lower() in f["name"].lower():
                print(f"  [DRIVE] Carpeta encontrada (por abrev): {f['name']}")
                return f["id"]

        # Busqueda 3: nombre completo "Febrero"
        for f in folders:
            if f["name"].lower() == month_name.lower():
                print(f"  [DRIVE] Carpeta encontrada (nombre completo): {f['name']}")
                return f["id"]

        # No existe -> crear con formato existente
        folder_metadata = {
            "name": expected_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_folder_id],
        }

        folder = drive_service.files().create(
            body=folder_metadata,
            fields="id",
            supportsAllDrives=True,
        ).execute()

        print(f"  [DRIVE] Carpeta creada: {expected_name}")
        return folder["id"]

    except Exception as e:
        print(f"  [DRIVE] Error con carpeta {month_name}: {e}")
        raise


def _buscar_archivo(drive_service, folder_id, filename):
    """Busca un archivo por nombre en una carpeta de Shared Drive."""
    # Escapar comillas simples en el nombre
    safe_name = filename.replace("'", "\\'")
    query = (
        f"'{folder_id}' in parents "
        f"and name = '{safe_name}' "
        f"and trashed = false"
    )

    try:
        results = drive_service.files().list(
            q=query,
            fields="files(id, name, webViewLink)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        files = results.get("files", [])
        return files[0] if files else None

    except Exception:
        return None
