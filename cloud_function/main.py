"""
BOT TELEGRAM - Webhook para Google Cloud Functions
===================================================
Responde consultas sobre los reportes de valorizacion semanal.
Lee los datos de resumen.json (desde GitHub raw URL).

Comandos disponibles:
  /start          - Bienvenida y ayuda
  /ayuda          - Lista de comandos
  /resumen        - Resumen de todas las obras
  /obra NOMBRE    - Detalle de una obra especifica
  /montos         - Tabla de montos (CD, GG, Total)
  /costos NOMBRE  - Desglose de costos de una obra

Deploy en Google Cloud Functions:
  gcloud functions deploy val-semanal-bot \
    --runtime python312 \
    --trigger-http \
    --allow-unauthenticated \
    --entry-point webhook \
    --set-env-vars TELEGRAM_BOT_TOKEN=xxx,GITHUB_RAW_URL=xxx
"""
import os
import json
import functions_framework
import requests
import urllib.request

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "")

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
    """Descarga resumen.json desde GitHub API (funciona con repos privados)."""
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/resumen.json"
        req = urllib.request.Request(api_url)
        req.add_header("Accept", "application/vnd.github.v3.raw")
        if GITHUB_TOKEN:
            req.add_header("Authorization", f"token {GITHUB_TOKEN}")
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"Error cargando resumen: {e}")
        return None


def _resolver_obra(texto):
    """Resuelve un nombre de obra desde texto libre."""
    texto_lower = texto.strip().lower()

    # Busqueda directa en aliases
    if texto_lower in OBRA_ALIASES:
        return OBRA_ALIASES[texto_lower]

    # Busqueda parcial
    for alias, obra_key in OBRA_ALIASES.items():
        if alias in texto_lower or texto_lower in alias:
            return obra_key

    return texto.strip().upper()


def _fmt(numero):
    """Formatea numero como S/ 1,234,567.89"""
    if not numero or numero == 0:
        return "S/ 0.00"
    return f"S/ {numero:,.2f}"


def _fmt_pct(numero):
    """Formatea porcentaje."""
    if not numero:
        return "0.0%"
    return f"{numero:.1f}%"


def _handle_start(chat_id):
    """Comando /start - Bienvenida."""
    msg = (
        "Hola! Soy el bot de *Val. Semanal*\n\n"
        "Puedo darte informacion sobre los reportes de valorizacion semanal.\n\n"
        "*Comandos:*\n"
        "/resumen - Resumen de todas las obras\n"
        "/montos - Tabla de montos\n"
        "/obra NOMBRE - Detalle de una obra\n"
        "/costos NOMBRE - Desglose de costos\n"
        "/ayuda - Ver esta ayuda"
    )
    return _send(chat_id, msg)


def _handle_resumen(chat_id):
    """Comando /resumen - Resumen general."""
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

    return _send(chat_id, msg)


def _handle_montos(chat_id):
    """Comando /montos - Tabla de montos."""
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

    return _send(chat_id, msg)


def _handle_obra(chat_id, nombre):
    """Comando /obra NOMBRE - Detalle de una obra."""
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

    return _send(chat_id, msg)


def _handle_costos(chat_id, nombre):
    """Comando /costos NOMBRE - Desglose de costos."""
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

    return _send(chat_id, msg)


def _handle_text(chat_id, texto):
    """Manejo de texto libre (no comando)."""
    # Intentar resolver como nombre de obra
    obra_key = _resolver_obra(texto)
    data = _cargar_resumen()
    if data and obra_key in data.get("reportes", {}):
        return _handle_obra(chat_id, texto)

    return _send(
        chat_id,
        "No entendi tu mensaje.\n"
        "Usa /ayuda para ver los comandos disponibles.\n"
        "O escribe el nombre de una obra (ej: mara, beethoven)."
    )


def _send(chat_id, text):
    """Envia mensaje por Telegram."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown",
        "disable_web_page_preview": True,
    }
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception as e:
        print(f"Error enviando: {e}")
    return "ok"


@functions_framework.http
def webhook(request):
    """Entry point para Google Cloud Functions (HTTP trigger)."""
    if request.method != "POST":
        return "OK", 200

    try:
        update = request.get_json(silent=True)
        if not update or "message" not in update:
            return "OK", 200

        message = update["message"]
        chat_id = message["chat"]["id"]
        text = message.get("text", "").strip()

        if not text:
            return "OK", 200

        # Parsear comando
        if text.startswith("/"):
            parts = text.split(maxsplit=1)
            cmd = parts[0].lower().split("@")[0]  # Remove @botname
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
                _handle_text(chat_id, text[1:])  # Remove /
        else:
            _handle_text(chat_id, text)

    except Exception as e:
        print(f"Error en webhook: {e}")

    return "OK", 200
