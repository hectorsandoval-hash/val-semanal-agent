"""
BOT TELEGRAM - Modo Polling Local
===================================
Bot conversacional que responde consultas sobre los reportes
de valorizacion semanal en TIEMPO REAL.

Lee los datos de resumen.json desde GitHub (se actualiza automaticamente
cuando GitHub Actions procesa nuevos reportes).

Uso:
  python bot_telegram.py          # Iniciar bot en modo polling
  python bot_telegram.py --test   # Enviar mensaje de prueba

Comandos del bot:
  /start          - Bienvenida y ayuda
  /ayuda          - Lista de comandos
  /resumen        - Resumen de todas las obras
  /obra NOMBRE    - Detalle de una obra especifica
  /montos         - Tabla de montos (CD, GG, Total)
  /costos NOMBRE  - Desglose de costos de una obra
"""
import os
import sys
import json
import time
import argparse
import urllib.request
import requests
from datetime import datetime, timezone, timedelta

# Zona horaria Peru
PERU_TZ = timezone(timedelta(hours=-5))

# Configuracion - prioriza variables de entorno, luego config.py
try:
    from config import TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
except ImportError:
    TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
    TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# URL del resumen.json en GitHub (se actualiza con cada ejecucion del pipeline)
GITHUB_RAW_URL = os.environ.get("GITHUB_RAW_URL", "")

# Tambien intenta leer el archivo local como fallback
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_RESUMEN = os.path.join(BASE_DIR, "resumen.json")

# Nombres cortos validos para buscar obras
OBRA_ALIASES = {
    "beethoven": "BEETHOVEN", "btv": "BEETHOVEN",
    "alma mater": "ALMA MATER", "mater": "ALMA MATER", "alma": "ALMA MATER",
    "mara": "MARA", "cema": "MARA", "cema mara": "MARA",
    "cenepa": "CENEPA", "heroes": "CENEPA",
    "biomedicas": "BIOMEDICAS", "biomedica": "BIOMEDICAS", "bio": "BIOMEDICAS",
    "roosevelt": "ROOSEVELT", "rooselvet": "ROOSEVELT",
}

# Cache del resumen (se refresca cada 5 minutos)
_cache = {"data": None, "timestamp": 0}
CACHE_TTL = 300  # 5 minutos


def _cargar_resumen():
    """Carga resumen.json - primero intenta GitHub, luego archivo local."""
    now = time.time()

    # Usar cache si es reciente
    if _cache["data"] and (now - _cache["timestamp"]) < CACHE_TTL:
        return _cache["data"]

    # Intentar desde GitHub
    try:
        req = urllib.request.Request(GITHUB_RAW_URL)
        req.add_header("Cache-Control", "no-cache")
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            _cache["data"] = data
            _cache["timestamp"] = now
            return data
    except Exception as e:
        print(f"  [WARN] No se pudo cargar desde GitHub: {e}")

    # Fallback: archivo local
    if os.path.exists(LOCAL_RESUMEN):
        try:
            with open(LOCAL_RESUMEN, "r", encoding="utf-8") as f:
                data = json.load(f)
                _cache["data"] = data
                _cache["timestamp"] = now
                return data
        except Exception:
            pass

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
        "/ayuda - Ver esta ayuda\n\n"
        "_Tambien puedes escribir el nombre de una obra directamente._\n"
        "_Ej: mara, beethoven, cenepa_"
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
        resp = requests.post(url, json=payload, timeout=10)
        if resp.status_code != 200:
            print(f"  [WARN] Telegram respuesta {resp.status_code}: {resp.text[:100]}")
    except Exception as e:
        print(f"  [ERROR] Enviando mensaje: {e}")
    return "ok"


def _process_update(update):
    """Procesa un update de Telegram."""
    message = update.get("message")
    if not message:
        return

    chat_id = message["chat"]["id"]
    text = message.get("text", "").strip()

    if not text:
        return

    print(f"  [{datetime.now(PERU_TZ).strftime('%H:%M:%S')}] "
          f"Chat {chat_id}: {text[:50]}")

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


def run_polling():
    """Ejecuta el bot en modo long-polling."""
    print("=" * 50)
    print("  BOT TELEGRAM - Val. Semanal")
    print("  Modo: Long Polling (tiempo real)")
    print(f"  Hora: {datetime.now(PERU_TZ).strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)

    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == "TU_BOT_TOKEN_AQUI":
        print("\n[ERROR] TELEGRAM_BOT_TOKEN no configurado.")
        print("Configura el token en config.py o como variable de entorno.")
        sys.exit(1)

    # Verificar que el bot funciona
    try:
        resp = requests.get(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getMe",
            timeout=10
        )
        bot_info = resp.json()
        if bot_info.get("ok"):
            bot_name = bot_info["result"].get("username", "?")
            print(f"\n  Bot conectado: @{bot_name}")
        else:
            print(f"\n[ERROR] Token invalido: {bot_info}")
            sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] No se pudo conectar: {e}")
        sys.exit(1)

    # Verificar que hay datos
    data = _cargar_resumen()
    if data and data.get("reportes"):
        n_obras = len(data["reportes"])
        print(f"  Datos cargados: {n_obras} obras")
        print(f"  Ultima actualizacion: {data.get('ultima_actualizacion', '?')}")
    else:
        print("  [WARN] No hay datos en resumen.json aun.")
        print("  El bot respondera cuando se procesen reportes.")

    print(f"\n  Escuchando mensajes... (Ctrl+C para detener)\n")

    offset = None
    consecutive_errors = 0

    while True:
        try:
            params = {"timeout": 30, "allowed_updates": ["message"]}
            if offset:
                params["offset"] = offset

            resp = requests.get(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates",
                params=params,
                timeout=35,
            )

            if resp.status_code != 200:
                print(f"  [WARN] getUpdates: {resp.status_code}")
                consecutive_errors += 1
                if consecutive_errors > 10:
                    print("  [ERROR] Demasiados errores consecutivos. Deteniendo.")
                    break
                time.sleep(5)
                continue

            result = resp.json()
            consecutive_errors = 0

            if not result.get("ok"):
                print(f"  [WARN] Respuesta no OK: {result}")
                time.sleep(5)
                continue

            updates = result.get("result", [])

            for update in updates:
                offset = update["update_id"] + 1
                try:
                    _process_update(update)
                except Exception as e:
                    print(f"  [ERROR] Procesando update: {e}")

        except KeyboardInterrupt:
            print("\n\n  Bot detenido por el usuario.")
            break
        except requests.exceptions.Timeout:
            continue  # Normal en long-polling
        except Exception as e:
            print(f"  [ERROR] {e}")
            consecutive_errors += 1
            if consecutive_errors > 10:
                print("  [ERROR] Demasiados errores. Deteniendo.")
                break
            time.sleep(5)


def send_test():
    """Envia un mensaje de prueba al chat configurado."""
    if not TELEGRAM_CHAT_ID or TELEGRAM_CHAT_ID == "TU_CHAT_ID_AQUI":
        print("[ERROR] TELEGRAM_CHAT_ID no configurado.")
        return

    data = _cargar_resumen()
    if data and data.get("reportes"):
        n = len(data["reportes"])
        msg = (
            f"*Bot Val. Semanal - Test*\n\n"
            f"Bot funcionando correctamente.\n"
            f"Datos cargados: {n} obras\n"
            f"Actualizado: {data.get('ultima_actualizacion', '?')}\n\n"
            f"Escribe /ayuda para ver los comandos."
        )
    else:
        msg = (
            f"*Bot Val. Semanal - Test*\n\n"
            f"Bot funcionando pero sin datos aun.\n"
            f"Los datos se cargaran cuando se procesen reportes."
        )

    _send(TELEGRAM_CHAT_ID, msg)
    print("Mensaje de prueba enviado.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Bot Telegram Val. Semanal")
    parser.add_argument("--test", action="store_true",
                        help="Enviar mensaje de prueba")
    args = parser.parse_args()

    if args.test:
        send_test()
    else:
        run_polling()
