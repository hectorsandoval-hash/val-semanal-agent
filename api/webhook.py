"""
Webhook para Vercel Serverless Function.
Recibe updates de Telegram y responde con datos de los reportes.
"""
import json
import os
import urllib.request
from http.server import BaseHTTPRequestHandler

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
GITHUB_RAW_URL = os.environ.get("GITHUB_RAW_URL", "")

# Nombres cortos validos para buscar obras
OBRA_ALIASES = {
    "beethoven": "BEETHOVEN", "btv": "BEETHOVEN",
    "alma mater": "ALMA MATER", "mater": "ALMA MATER", "alma": "ALMA MATER",
    "mara": "MARA", "cema": "MARA", "cema mara": "MARA",
    "cenepa": "CENEPA", "heroes": "CENEPA",
    "biomedicas": "BIOMEDICAS", "biomedica": "BIOMEDICAS", "bio": "BIOMEDICAS",
    "roosevelt": "ROOSEVELT", "rooselvet": "ROOSEVELT",
}


def _cargar_resumen():
    """Descarga resumen.json desde GitHub."""
    try:
        req = urllib.request.Request(GITHUB_RAW_URL)
        req.add_header("Cache-Control", "no-cache")
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"Error cargando resumen: {e}")
        return None


def _resolver_obra(texto):
    """Resuelve un nombre de obra desde texto libre."""
    texto_lower = texto.strip().lower()
    if texto_lower in OBRA_ALIASES:
        return OBRA_ALIASES[texto_lower]
    for alias, obra_key in OBRA_ALIASES.items():
        if alias in texto_lower or texto_lower in alias:
            return obra_key
    return texto.strip().upper()


def _fmt(numero):
    if not numero or numero == 0:
        return "S/ 0.00"
    return f"S/ {numero:,.2f}"


def _fmt_pct(numero):
    if not numero:
        return "0.0%"
    return f"{numero:.1f}%"


def _send(chat_id, text):
    """Envia mensaje por Telegram."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = json.dumps({
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown",
        "disable_web_page_preview": True,
    }).encode("utf-8")

    req = urllib.request.Request(url, data=payload)
    req.add_header("Content-Type", "application/json")
    try:
        urllib.request.urlopen(req, timeout=10)
    except Exception as e:
        print(f"Error enviando: {e}")


def _handle_start(chat_id):
    _send(chat_id, (
        "Hola! Soy el bot de *Val. Semanal*\n\n"
        "Puedo darte informacion sobre los reportes de valorizacion semanal.\n\n"
        "*Comandos:*\n"
        "/resumen - Resumen de todas las obras\n"
        "/montos - Tabla de montos\n"
        "/obra NOMBRE - Detalle de una obra\n"
        "/costos NOMBRE - Desglose de costos\n"
        "/ayuda - Ver esta ayuda\n\n"
        "_Tambien puedes escribir el nombre de una obra._\n"
        "_Ej: mara, beethoven, cenepa_"
    ))


def _handle_resumen(chat_id):
    data = _cargar_resumen()
    if not data or not data.get("reportes"):
        return _send(chat_id, "No hay reportes procesados aun.")

    msg = f"*Resumen Val. Semanal*\n_Actualizado: {data.get('ultima_actualizacion', '?')}_\n\n"
    for obra_key, r in sorted(data["reportes"].items()):
        total = r.get("total_valorizacion", 0)
        cd = r.get("costo_directo", 0)
        fecha = r.get("fecha_reporte", "?")
        msg += f"*{obra_key}*\n"
        msg += f"  CD: {_fmt(cd)}\n"
        msg += f"  Total: {_fmt(total)}\n"
        msg += f"  Fecha: {fecha}\n\n"
    _send(chat_id, msg)


def _handle_montos(chat_id):
    data = _cargar_resumen()
    if not data or not data.get("reportes"):
        return _send(chat_id, "No hay reportes procesados aun.")

    msg = "*Montos Val. Semanal*\n\n"
    msg += "`Obra          |    CD         |    GG         |    Total`\n"
    msg += "`--------------+--------------+--------------+--------------`\n"
    for obra_key, r in sorted(data["reportes"].items()):
        cd = r.get("costo_directo", 0)
        gg = r.get("gastos_generales", 0)
        total = r.get("total_valorizacion", 0)
        nombre = obra_key[:13].ljust(13)
        msg += f"`{nombre} | {cd:>12,.0f} | {gg:>12,.0f} | {total:>12,.0f}`\n"
    _send(chat_id, msg)


def _handle_obra(chat_id, nombre):
    if not nombre:
        return _send(chat_id, "Indica el nombre de la obra.\nEj: /obra MARA")

    obra_key = _resolver_obra(nombre)
    data = _cargar_resumen()
    if not data:
        return _send(chat_id, "Error cargando datos.")

    r = data.get("reportes", {}).get(obra_key)
    if not r:
        obras = ", ".join(data.get("reportes", {}).keys())
        return _send(chat_id, f"No se encontro '{nombre}'.\nObras disponibles: {obras}")

    msg = f"*{obra_key}*\n"
    msg += f"Proyecto: _{r.get('proyecto', '')[:80]}_\n\n"
    msg += f"*Valoracion:*\n"
    msg += f"  Costo Directo: {_fmt(r.get('costo_directo', 0))}\n"
    msg += f"  Gastos Generales ({_fmt_pct(r.get('gg_percent', 0))}): {_fmt(r.get('gastos_generales', 0))}\n"
    msg += f"  Utilidad ({_fmt_pct(r.get('util_percent', 0))}): {_fmt(r.get('utilidad', 0))}\n"
    msg += f"  *Total Valorizacion: {_fmt(r.get('total_valorizacion', 0))}*\n\n"
    msg += f"Fecha reporte: {r.get('fecha_reporte', '?')}\n"
    msg += f"Procesado: {r.get('fecha_procesado', '?')}\n"
    if r.get("drive_link"):
        msg += f"\n[Ver reporte en Drive]({r['drive_link']})"
    _send(chat_id, msg)


def _handle_costos(chat_id, nombre):
    if not nombre:
        return _send(chat_id, "Indica el nombre de la obra.\nEj: /costos BEETHOVEN")

    obra_key = _resolver_obra(nombre)
    data = _cargar_resumen()
    if not data:
        return _send(chat_id, "Error cargando datos.")

    r = data.get("reportes", {}).get(obra_key)
    if not r:
        return _send(chat_id, f"No se encontro '{nombre}'.")

    msg = f"*Desglose de Costos - {obra_key}*\n\n"
    msg += f"*Costo Directo:*\n"
    msg += f"  Personal Obrero: {_fmt(r.get('personal_obrero', 0))}\n"
    msg += f"  Materiales: {_fmt(r.get('materiales', 0))}\n"
    msg += f"  Alquileres: {_fmt(r.get('alquileres', 0))}\n"
    msg += f"  Subcontratos: {_fmt(r.get('subcontratos', 0))}\n"
    msg += f"  Costos Varios: {_fmt(r.get('costos_varios', 0))}\n"
    msg += f"  *Total CD: {_fmt(r.get('total_cd', 0))}*\n\n"
    msg += f"*Gastos Generales:*\n"
    msg += f"  Planilla Staff: {_fmt(r.get('planilla_staff', 0))}\n"
    msg += f"  Otros GG: {_fmt(r.get('otros_gg', 0))}\n"
    msg += f"  *Total GG: {_fmt(r.get('total_gg', 0))}*\n"
    _send(chat_id, msg)


def _handle_text(chat_id, texto):
    obra_key = _resolver_obra(texto)
    data = _cargar_resumen()
    if data and obra_key in data.get("reportes", {}):
        return _handle_obra(chat_id, texto)

    _send(chat_id, (
        "No entendi tu mensaje.\n"
        "Usa /ayuda para ver los comandos disponibles.\n"
        "O escribe el nombre de una obra (ej: mara, beethoven)."
    ))


def process_update(update):
    """Procesa un update de Telegram."""
    message = update.get("message")
    if not message:
        return

    chat_id = message["chat"]["id"]
    text = message.get("text", "").strip()
    if not text:
        return

    if text.startswith("/"):
        parts = text.split(maxsplit=1)
        cmd = parts[0].lower().split("@")[0]
        arg = parts[1] if len(parts) > 1 else ""

        if cmd in ("/start", "/ayuda", "/help"):
            _handle_start(chat_id)
        elif cmd == "/resumen":
            _handle_resumen(chat_id)
        elif cmd == "/montos":
            _handle_montos(chat_id)
        elif cmd in ("/obra", "/detalle"):
            _handle_obra(chat_id, arg)
        elif cmd in ("/costos", "/desglose"):
            _handle_costos(chat_id, arg)
        else:
            _handle_text(chat_id, text[1:])
    else:
        _handle_text(chat_id, text)


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get("content-length", 0))
        body = self.rfile.read(content_length)

        try:
            update = json.loads(body)
            process_update(update)
        except Exception as e:
            print(f"Error: {e}")

        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"OK")

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"Bot Val. Semanal - Running")
