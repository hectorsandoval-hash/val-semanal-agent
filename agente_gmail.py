"""
AGENTE GMAIL - Busqueda de correos con valorizaciones semanales
===============================================================
Busca correos nuevos que contengan archivos Excel de valorizacion,
los descarga y detecta a que obra pertenecen.

Deteccion de obra:
  1. Primario: por email del remitente (mapeado en SENDER_TO_OBRA)
  2. Fallback: por keywords en el asunto del correo
  3. Ultimo recurso: por contenido del Excel (nombre del proyecto)
"""
import base64
import re
from datetime import datetime, timezone, timedelta

from config import (
    EMAIL_SENDERS, EMAIL_SUBJECT_KEYWORDS, DIAS_BUSQUEDA,
    GMAIL_LABEL, OBRA_KEYWORDS, SENDER_TO_OBRA,
)

# Zona horaria Peru (UTC-5)
PERU_TZ = timezone(timedelta(hours=-5))


def buscar_correos_valorizacion(service, max_results=20):
    """
    Busca correos nuevos de los remitentes de costos con adjuntos Excel.

    La busqueda usa los remitentes configurados en SENDER_TO_OBRA
    y opcionalmente keywords en el asunto. Excluye correos ya procesados.

    Returns:
        Lista de dicts con info de cada correo encontrado.
    """
    # Construir query: remitentes conocidos + keywords en asunto + con adjunto
    sender_query = " OR ".join([f"from:{s}" for s in EMAIL_SENDERS])
    subject_query = " OR ".join([f'subject:"{kw}"' for kw in EMAIL_SUBJECT_KEYWORDS])

    # Query precisa: remitente + asunto con keyword + adjunto + no procesado
    query = (
        f"({sender_query}) ({subject_query}) "
        f"has:attachment newer_than:{DIAS_BUSQUEDA}d -label:{GMAIL_LABEL}"
    )

    print(f"  [GMAIL] Query: {query[:100]}...")

    try:
        results = service.users().messages().list(
            userId="me", q=query, maxResults=max_results
        ).execute()
    except Exception as e:
        print(f"  [GMAIL] Error buscando correos: {e}")
        return []

    messages = results.get("messages", [])
    if not messages:
        print("  [GMAIL] No se encontraron correos nuevos.")
        return []

    print(f"  [GMAIL] Encontrados {len(messages)} correo(s).")

    correos = []
    for msg_ref in messages:
        try:
            msg = service.users().messages().get(
                userId="me", id=msg_ref["id"], format="full"
            ).execute()

            headers = {h["name"].lower(): h["value"] for h in msg["payload"].get("headers", [])}
            asunto = headers.get("subject", "Sin asunto")
            de = headers.get("from", "Desconocido")
            fecha = headers.get("date", "")

            # Extraer email del remitente (puede venir como "Nombre <email>")
            de_email = _extraer_email(de)

            # Buscar adjuntos Excel
            adjuntos = _buscar_adjuntos_excel(msg["payload"])

            if not adjuntos:
                print(f"  [GMAIL] '{asunto[:40]}...' sin adjunto Excel, saltando.")
                continue

            # Detectar obra: PRIMERO por email, luego por asunto
            obra = detectar_obra_de_sender(de_email)
            if not obra:
                obra = detectar_obra_de_asunto(asunto)

            correos.append({
                "id": msg_ref["id"],
                "thread_id": msg.get("threadId", ""),
                "asunto": asunto,
                "de": de,
                "de_email": de_email,
                "fecha": fecha,
                "obra_detectada": obra,
                "adjuntos": adjuntos,
                "label_ids": msg.get("labelIds", []),
            })

        except Exception as e:
            print(f"  [GMAIL] Error procesando mensaje {msg_ref['id']}: {e}")

    return correos


def descargar_adjunto_excel(service, mensaje_id, attachment_id):
    """
    Descarga un adjunto de Gmail y retorna sus bytes.

    Returns:
        bytes del archivo, o None si falla
    """
    try:
        attachment = service.users().messages().attachments().get(
            userId="me", messageId=mensaje_id, id=attachment_id
        ).execute()
        file_data = base64.urlsafe_b64decode(attachment["data"])
        return file_data
    except Exception as e:
        print(f"  [GMAIL] Error descargando adjunto: {e}")
        return None


def marcar_procesado(service, correo):
    """
    Agrega el label de procesado al correo para no reprocesarlo.
    """
    try:
        label_id = _get_or_create_label(service, GMAIL_LABEL)

        service.users().messages().modify(
            userId="me",
            id=correo["id"],
            body={"addLabelIds": [label_id]}
        ).execute()

        print(f"  [GMAIL] Correo marcado como procesado: {correo['asunto'][:40]}")

    except Exception as e:
        print(f"  [GMAIL] Error marcando correo: {e}")


def detectar_obra_de_sender(email):
    """
    Detecta la obra a partir del email del remitente.
    Forma MAS CONFIABLE de deteccion.

    Returns:
        Nombre corto de la obra o None.
    """
    if not email:
        return None
    email_lower = email.lower().strip()
    return SENDER_TO_OBRA.get(email_lower)


def detectar_obra_de_asunto(asunto):
    """
    Detecta la obra a partir de keywords en el asunto.
    Fallback cuando no se puede detectar por email.

    Returns:
        Nombre corto de la obra o None.
    """
    if not asunto:
        return None
    asunto_lower = asunto.lower()

    for obra_key, keywords in OBRA_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in asunto_lower:
                return obra_key

    return None


def detectar_obra_de_texto(texto):
    """
    Detecta la obra a partir de un texto generico
    (nombre del proyecto extraido del Excel).

    Returns:
        Nombre corto de la obra o None.
    """
    if not texto:
        return None
    texto_lower = texto.lower()

    for obra_key, keywords in OBRA_KEYWORDS.items():
        for kw in keywords:
            if kw.lower() in texto_lower:
                return obra_key

    return None


# ============================================================
# FUNCIONES INTERNAS
# ============================================================

def _extraer_email(campo_from):
    """
    Extrae el email de un campo From que puede venir como:
    - "email@ejemplo.com"
    - "Nombre <email@ejemplo.com>"
    - "Nombre Apellido <email@ejemplo.com>"
    """
    if not campo_from:
        return ""

    # Buscar entre < >
    match = re.search(r"<([^>]+)>", campo_from)
    if match:
        return match.group(1).lower().strip()

    # Si no hay < >, asumir que todo es el email
    return campo_from.lower().strip()


def _buscar_adjuntos_excel(payload):
    """
    Busca recursivamente adjuntos Excel (.xlsx, .xlsm, .xls) en el payload.

    Returns:
        Lista de dicts con filename y attachmentId.
    """
    adjuntos = []
    _buscar_adjuntos_recursivo(payload, adjuntos)
    return adjuntos


def _buscar_adjuntos_recursivo(payload, adjuntos):
    """Recorre recursivamente el payload buscando adjuntos Excel."""
    filename = payload.get("filename", "")
    attachment_id = payload.get("body", {}).get("attachmentId")

    if filename and attachment_id:
        ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
        if ext in ("xlsx", "xlsm", "xls"):
            adjuntos.append({
                "filename": filename,
                "attachmentId": attachment_id,
                "mimeType": payload.get("mimeType", ""),
            })

    for part in payload.get("parts", []):
        _buscar_adjuntos_recursivo(part, adjuntos)


def _get_or_create_label(service, label_name):
    """
    Obtiene el ID de un label de Gmail, o lo crea si no existe.
    """
    labels = service.users().labels().list(userId="me").execute()
    for label in labels.get("labels", []):
        if label["name"] == label_name:
            return label["id"]

    label_body = {
        "name": label_name,
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show",
    }
    created = service.users().labels().create(userId="me", body=label_body).execute()
    print(f"  [GMAIL] Label '{label_name}' creado.")
    return created["id"]
