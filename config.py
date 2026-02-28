"""
Configuracion central del agente de valorizacion semanal.
=====================================================
Disenado para ejecucion en LA NUBE (GitHub Actions).
Todos los datos sensibles se cargan desde variables de entorno / secretos.
"""
import os
import json as _json

# Ruta base del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# CREDENCIALES OAUTH2
# ============================================================
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "token.json")

# Scopes: Gmail (lectura + labels) + Drive (lectura + escritura)
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
]

# ============================================================
# MAPEO PRINCIPAL: EMAIL REMITENTE -> OBRA
# ============================================================
# Se carga desde variable de entorno SENDER_TO_OBRA (JSON)
# Formato: {"email1@dominio.com":"OBRA1","email2@dominio.com":"OBRA2",...}
_sender_raw = os.environ.get("SENDER_TO_OBRA", "")
if _sender_raw:
    SENDER_TO_OBRA = _json.loads(_sender_raw)
else:
    SENDER_TO_OBRA = {}

# Lista de TODOS los remitentes (para la query de Gmail)
EMAIL_SENDERS = list(SENDER_TO_OBRA.keys())

# ============================================================
# PATRON DE BUSQUEDA EN ASUNTO
# ============================================================
EMAIL_SUBJECT_KEYWORDS = [
    "reporte semanal",
    "val semanal",
    "val. semanal",
    "valorizacion semanal",
    "COS-PR02-FR02",
    "COS-PRO2-FR02",
    "corte de venta y costo",
]

# Rango de busqueda en dias
DIAS_BUSQUEDA = 7

# Label de Gmail para marcar correos ya procesados
GMAIL_LABEL = "ValSemanal-Procesado"

# ============================================================
# FOLDER IDS EN GOOGLE DRIVE
# ============================================================
# Se carga desde variable de entorno OBRA_FOLDER_IDS (JSON)
# Formato: {"OBRA1":"folder_id_1","OBRA2":"folder_id_2",...}
_folder_ids_raw = os.environ.get("OBRA_FOLDER_IDS", "")
if _folder_ids_raw:
    OBRA_FOLDER_IDS = _json.loads(_folder_ids_raw)
else:
    OBRA_FOLDER_IDS = {}

# Keywords para detectar obra desde asunto o contenido del Excel
# (fallback si no se puede detectar desde el email del remitente)
OBRA_KEYWORDS = {
    "BEETHOVEN": ["beethoven", "btv"],
    "ALMA MATER": ["alma mater", "mater"],
    "MARA": ["mara", "cema mara"],
    "CENEPA": ["cenepa", "heroes cenep"],
    "BIOMEDICAS": ["biomedicas", "biomedic", "biomédicas"],
    "ROOSEVELT": ["roosevelt", "rooselvet"],
}

# ============================================================
# NOMBRES DE MESES (para subcarpetas dentro de Val. Semanal)
# ============================================================
MONTH_NAMES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre",
]
MONTH_ABBREVS = [
    "Ene", "Feb", "Mar", "Abr", "May", "Jun",
    "Jul", "Ago", "Set", "Oct", "Nov", "Dic",
]

# ============================================================
# TELEGRAM BOT
# ============================================================
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# ============================================================
# DIRECTORIOS LOCALES DEL PROYECTO
# ============================================================
TEMP_DIR = os.path.join(BASE_DIR, "temp_files")
LOG_DIR = os.path.join(BASE_DIR, "logs")
REPORT_DIR = os.path.join(BASE_DIR, "reportes")
LOG_FILE = os.path.join(LOG_DIR, "ejecucion.log")
