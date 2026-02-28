"""
BOT TELEGRAM - Worker de Polling para GitHub Actions
=====================================================
Motor de IA con GitHub Models (gpt-4o) + memoria persistente.
- Memoria de conversacion (contexto inmediato, in-memory)
- Memoria de estilo (preferencias persistentes, guardadas en repo)

Diseñado para correr durante horario laboral (9AM-6PM Peru, Lun-Vie).
Cada instancia corre hasta MAX_RUNTIME_HOURS, luego termina limpiamente.
"""
import os
import sys
import json
import time
import signal
import re
import base64
import urllib.request
import requests
from datetime import datetime, timezone, timedelta

# ============================================================
# CONFIGURACION
# ============================================================

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "")

# GitHub Models API (gratuito con token de GitHub)
# Prioridad: GH_MODELS_TOKEN > GITHUB_TOKEN (Actions provee uno automaticamente)
GH_MODELS_TOKEN = os.environ.get("GH_MODELS_TOKEN", "") or os.environ.get("GITHUB_TOKEN", "")
AI_MODEL = "gpt-4o"
AI_MODEL_FALLBACK = "gpt-4o-mini"
AI_API_URL = "https://models.inference.ai.azure.com/chat/completions"

MAX_RUNTIME_HOURS = float(os.environ.get("MAX_RUNTIME_HOURS", "5.5"))
PERU_TZ = timezone(timedelta(hours=-5))
HORA_FIN = int(os.environ.get("HORA_FIN", "18"))

# Cache del resumen
_cache = {"data": None, "timestamp": 0}
CACHE_TTL = 120

# ============================================================
# MEMORIA
# ============================================================

# Historial de conversacion por chat (in-memory, se pierde al reiniciar)
_historiales = {}  # {chat_id: [{"role": "user"/"assistant", "content": "..."}]}
MAX_HISTORIAL = 8  # ultimos 8 mensajes (4 intercambios)

# Preferencias de estilo persistentes (se guardan en el repo)
_preferencias = []  # ["Mostrar montos en miles", "Usar tablas para comparaciones", ...]
MAX_PREFERENCIAS = 20
PREFERENCIAS_FILE = "preferencias.json"

# Control de senales
_running = True

def _signal_handler(signum, frame):
    global _running
    print(f"\n  Senal recibida ({signum}). Terminando...")
    _running = False

signal.signal(signal.SIGTERM, _signal_handler)
signal.signal(signal.SIGINT, _signal_handler)


# ============================================================
# FUNCIONES DE DATOS
# ============================================================

def _cargar_resumen():
    """Carga resumen.json: archivo local (artifact) > GitHub API > cache."""
    now = time.time()
    if _cache["data"] and (now - _cache["timestamp"]) < CACHE_TTL:
        return _cache["data"]

    # 1. Intentar archivo local (descargado como artifact del pipeline)
    local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resumen.json")
    if os.path.exists(local_path):
        try:
            with open(local_path, encoding="utf-8") as f:
                data = json.load(f)
            _cache["data"] = data
            _cache["timestamp"] = now
            return data
        except Exception as e:
            print(f"  [WARN] Error leyendo resumen local: {e}")

    # 2. Fallback: GitHub API (si el archivo existe en el repo)
    if GITHUB_REPO and GITHUB_TOKEN:
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/resumen.json"
            req = urllib.request.Request(api_url)
            req.add_header("Accept", "application/vnd.github.v3.raw")
            req.add_header("Authorization", f"token {GITHUB_TOKEN}")
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                _cache["data"] = data
                _cache["timestamp"] = now
                return data
        except Exception as e:
            print(f"  [WARN] Error cargando resumen desde API: {e}")

    return _cache.get("data")


def _fmt(numero):
    if not numero or numero == 0:
        return "S/ 0.00"
    return f"S/ {numero:,.2f}"


# ============================================================
# PREFERENCIAS PERSISTENTES
# ============================================================

def _cargar_preferencias():
    """Carga preferencias de estilo desde el repo."""
    global _preferencias
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PREFERENCIAS_FILE}"
        resp = requests.get(
            api_url,
            headers={
                "Authorization": f"token {GITHUB_TOKEN}",
                "Accept": "application/vnd.github.v3.raw",
            },
            timeout=10,
        )
        if resp.status_code == 200:
            data = resp.json()
            _preferencias = data.get("reglas", [])[:MAX_PREFERENCIAS]
            print(f"  Preferencias: {len(_preferencias)} reglas cargadas")
            return True
    except Exception as e:
        print(f"  [WARN] Error cargando preferencias: {e}")
    return False


def _guardar_preferencias():
    """Guarda preferencias en el repo via GitHub API."""
    content_str = json.dumps(
        {"reglas": _preferencias}, ensure_ascii=False, indent=2
    )
    encoded = base64.b64encode(content_str.encode("utf-8")).decode("ascii")

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PREFERENCIAS_FILE}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }

    # Obtener SHA actual (necesario para update)
    sha = None
    try:
        r = requests.get(api_url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha = r.json().get("sha")
    except Exception:
        pass

    payload = {
        "message": "Bot: actualizar preferencias de estilo",
        "content": encoded,
    }
    if sha:
        payload["sha"] = sha

    try:
        r = requests.put(api_url, json=payload, headers=headers, timeout=10)
        ok = r.status_code in (200, 201)
        if ok:
            print(f"  Preferencias guardadas ({len(_preferencias)} reglas)")
        else:
            print(f"  [WARN] Error guardando preferencias: {r.status_code}")
        return ok
    except Exception as e:
        print(f"  [ERROR] Guardando preferencias: {e}")
        return False


# ============================================================
# HISTORIAL DE CONVERSACION
# ============================================================

def _agregar_historial(chat_id, role, content):
    """Agrega un mensaje al historial del chat."""
    if chat_id not in _historiales:
        _historiales[chat_id] = []
    _historiales[chat_id].append({"role": role, "content": content})
    # Mantener solo los ultimos MAX_HISTORIAL mensajes
    if len(_historiales[chat_id]) > MAX_HISTORIAL:
        _historiales[chat_id] = _historiales[chat_id][-MAX_HISTORIAL:]


def _obtener_historial(chat_id):
    """Retorna el historial del chat."""
    return _historiales.get(chat_id, [])


# ============================================================
# TELEGRAM
# ============================================================

def _send(chat_id, text):
    """Envia mensaje a Telegram. Si falla con Markdown, reintenta sin formato."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown",
        "disable_web_page_preview": True,
    }
    try:
        resp = requests.post(url, json=payload, timeout=15)
        if resp.status_code != 200:
            # Reintentar sin Markdown (a veces la IA genera markdown invalido)
            payload["parse_mode"] = ""
            resp2 = requests.post(url, json=payload, timeout=15)
            if resp2.status_code != 200:
                print(f"  [WARN] Telegram error: {resp2.text[:150]}")
    except Exception as e:
        print(f"  [ERROR] Enviando: {e}")


# ============================================================
# MOTOR IA - GitHub Models (gpt-4o, gratuito)
# ============================================================

SYSTEM_PROMPT = """Eres el asistente experto de valorizacion semanal de una constructora en Peru.
Respondes preguntas sobre datos de 6 obras: MARA, ROOSEVELT, CENEPA, BIOMEDICAS, ALMA MATER, BEETHOVEN.

ESTRUCTURA DE DATOS (JSON que recibes por obra):
- fecha: fecha del reporte
- val: corte de valorizacion -> cd, gg (con %), util (con %), sub_total, igv, total
- ejec_cd: gastos ejecutados CD -> obrero, materiales, alquileres, subcontratos, varios (con % y monto), total_cd
- ejec_gg: gastos ejecutados GG -> planilla, otros (con % y monto), total_gg
- analisis: comparativo Val vs Ejec -> para cd, gg y total: val, ejec, var (S/), var_pct (%), estado (GANANCIA/PERDIDA)
- curva: Curva S -> prog_pct, ejec_pct, plan_pct, desvio_pct, estado (ATRASADO/ADELANTADO), total_contractual
- link: link al reporte en Drive

ALIASES: alma/mater=ALMA MATER, bio=BIOMEDICAS, cema=MARA, heroes=CENEPA, btv=BEETHOVEN

REGLAS DE FORMATO:
- Formato Telegram Markdown: *bold*, _italic_
- Montos: S/ 1,234,567.89
- Porcentajes: 1-2 decimales
- Emojis: ✅ ganancia, 🔴 perdida, 📊📈📉💰🏗
- SIEMPRE di fecha del reporte y nombre corto de obra
- NUNCA uses el nombre largo del proyecto
- Analisis: muestra Val vs Ejec con variacion y estado para CD, GG y Total
- Rankings: ordena de mayor a menor
- Comparaciones: datos lado a lado
- Se directo y conciso

MEMORIA DE ESTILO:
Si el usuario te CORRIGE el formato, te pide que RECUERDES algo, o te indica COMO quiere ver la informacion
(ej: "siempre muestra...", "recuerda que...", "quiero que...", "de ahora en adelante...", "no pongas..."),
entonces al FINAL de tu respuesta agrega en una linea aparte:
[REGLA: descripcion concisa de la preferencia]
Solo agrega [REGLA:] cuando el usuario EXPLICITAMENTE te ensena un estilo o te corrige. NO en preguntas normales."""


def _compress_data(data):
    """Comprime el JSON eliminando campos innecesarios y redondeando numeros."""
    if not data or not data.get("reportes"):
        return data

    compressed = {}
    for nombre, r in data["reportes"].items():
        def rd(v, d=2):
            return round(v, d) if isinstance(v, (int, float)) and v else v

        obra = {
            "fecha": r.get("fecha_reporte", "?"),
            "val": {
                "cd": rd(r.get("costo_directo", 0)),
                "gg": rd(r.get("gastos_generales", 0)),
                "gg_pct": rd(r.get("gg_percent", 0), 1),
                "util": rd(r.get("utilidad", 0)),
                "util_pct": rd(r.get("util_percent", 0), 1),
                "sub_total": rd(r.get("sub_total", 0)),
                "igv": rd(r.get("igv", 0)),
                "total": rd(r.get("total_valorizacion", 0)),
            },
            "ejec_cd": {
                "obrero": rd(r.get("personal_obrero", 0)),
                "obrero_pct": rd(r.get("personal_obrero_pct", 0), 1),
                "materiales": rd(r.get("materiales", 0)),
                "materiales_pct": rd(r.get("materiales_pct", 0), 1),
                "alquileres": rd(r.get("alquileres", 0)),
                "alquileres_pct": rd(r.get("alquileres_pct", 0), 1),
                "subcontratos": rd(r.get("subcontratos", 0)),
                "subcontratos_pct": rd(r.get("subcontratos_pct", 0), 1),
                "varios": rd(r.get("costos_varios", 0)),
                "varios_pct": rd(r.get("costos_varios_pct", 0), 1),
                "total_cd": rd(r.get("total_cd_ejecutado", 0)),
            },
            "ejec_gg": {
                "planilla": rd(r.get("planilla_staff", 0)),
                "planilla_pct": rd(r.get("planilla_staff_pct", 0), 1),
                "otros": rd(r.get("otros_gg", 0)),
                "otros_pct": rd(r.get("otros_gg_pct", 0), 1),
                "total_gg": rd(r.get("total_gg_ejecutado", 0)),
            },
        }

        # Analisis
        analisis = r.get("analisis", {})
        if analisis:
            obra["analisis"] = {}
            for sec in ("cd", "gg", "total"):
                s = analisis.get(sec, {})
                obra["analisis"][sec] = {
                    "val": rd(s.get("valorizacion", 0)),
                    "ejec": rd(s.get("ejecutado", 0)),
                    "var": rd(s.get("variacion", 0)),
                    "var_pct": rd(s.get("variacion_pct", 0), 1),
                    "estado": s.get("estado", ""),
                }

        # Curva S
        if r.get("tiene_curva"):
            obra["curva"] = {
                "prog_pct": rd(r.get("curva_prog_acum_pct", 0) * 100, 1),
                "ejec_pct": rd(r.get("curva_ejec_acum_pct", 0) * 100, 1),
                "plan_pct": rd(r.get("curva_plan_acum_pct", 0) * 100, 1),
                "desvio_pct": rd(r.get("curva_desvio_pct", 0) * 100, 1),
                "estado": r.get("curva_estado", ""),
                "total_contractual": rd(r.get("curva_total_contractual", 0)),
            }

        if r.get("drive_link"):
            obra["link"] = r["drive_link"]

        compressed[nombre] = obra

    return compressed


def _build_messages(pregunta, data, chat_id=None):
    """Construye el array de mensajes para la API."""
    # Comprimir datos
    datos_comprimidos = _compress_data(data)
    datos_json = json.dumps(datos_comprimidos, ensure_ascii=False, separators=(",", ":"))

    # System prompt con preferencias
    system_content = SYSTEM_PROMPT
    if _preferencias:
        prefs = "\n".join(f"- {p}" for p in _preferencias)
        system_content += f"\n\nPREFERENCIAS DEL USUARIO (respeta SIEMPRE estas reglas de estilo):\n{prefs}"

    messages = [
        {"role": "system", "content": system_content},
        {"role": "user", "content": f"[DATOS ACTUALES DE LAS OBRAS]\n{datos_json}"},
        {"role": "assistant", "content": "Tengo los datos de las 6 obras. Pregunta lo que necesites."},
    ]

    # Historial de conversacion
    if chat_id:
        historial = _obtener_historial(chat_id)
        for msg in historial:
            messages.append(msg)

    # Pregunta actual
    messages.append({"role": "user", "content": pregunta})
    return messages


def _call_model(messages, model, token=None):
    """Llama a un modelo especifico. Retorna contenido o None."""
    token = token or GH_MODELS_TOKEN
    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.2,
        "max_tokens": 2048,
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(AI_API_URL, json=payload, headers=headers, timeout=45)
        if resp.status_code != 200:
            error_text = resp.text[:400]
            print(f"  [WARN] {model} HTTP {resp.status_code}: {error_text}")
            # Detectar errores de autenticacion
            if resp.status_code in (401, 403):
                print(f"  [AUTH] Token rechazado ({len(token)} chars)")
            return None
        result = resp.json()
        choices = result.get("choices", [])
        if choices:
            content = choices[0].get("message", {}).get("content", "")
            return content.strip() if content else None
        print(f"  [WARN] {model}: respuesta sin choices: {str(result)[:200]}")
        return None
    except requests.exceptions.Timeout:
        print(f"  [TIMEOUT] {model}: timeout de 45s")
        return None
    except Exception as e:
        print(f"  [ERROR] {model}: {type(e).__name__}: {e}")
        return None


def _ask_ai(pregunta, data, chat_id=None):
    """Envia la pregunta a la IA. Intenta gpt-4o, fallback a gpt-4o-mini.
    Si ambos fallan, intenta con GITHUB_TOKEN como token alternativo."""
    if not GH_MODELS_TOKEN:
        print("  [WARN] GH_MODELS_TOKEN vacio!")
        return None

    messages = _build_messages(pregunta, data, chat_id)

    # Intentar con gpt-4o primero
    result = _call_model(messages, AI_MODEL)
    if result:
        return result

    # Fallback a gpt-4o-mini si gpt-4o falla
    print(f"  [FALLBACK] Intentando con {AI_MODEL_FALLBACK}...")
    result = _call_model(messages, AI_MODEL_FALLBACK)
    if result:
        return result

    # Ultimo intento: usar GITHUB_TOKEN si es diferente
    alt_token = os.environ.get("GITHUB_TOKEN", "")
    if alt_token and alt_token != GH_MODELS_TOKEN:
        print(f"  [FALLBACK] Intentando con GITHUB_TOKEN...")
        result = _call_model(messages, AI_MODEL_FALLBACK, token=alt_token)
        if result:
            return result

    print(f"  [ERROR] Todos los intentos de IA fallaron.")
    return None


def _extraer_regla(respuesta):
    """Extrae [REGLA: ...] de la respuesta de la IA y la guarda."""
    match = re.search(r'\[REGLA:\s*(.+?)\]', respuesta)
    if match:
        regla = match.group(1).strip()
        if regla and len(regla) > 5:
            # Evitar duplicados
            reglas_lower = [r.lower() for r in _preferencias]
            if regla.lower() not in reglas_lower:
                _preferencias.append(regla)
                if len(_preferencias) > MAX_PREFERENCIAS:
                    _preferencias.pop(0)  # remover la mas antigua
                _guardar_preferencias()
                print(f"  [REGLA] Nueva preferencia: {regla}")
        # Limpiar el marcador de la respuesta
        respuesta = re.sub(r'\n?\[REGLA:\s*.+?\]', '', respuesta).strip()
    return respuesta


# ============================================================
# HANDLERS DE COMANDOS RAPIDOS (slash commands)
# ============================================================

def _handle_start(chat_id):
    n_prefs = len(_preferencias)
    prefs_txt = f"\n_Tengo {n_prefs} preferencias de estilo guardadas._" if n_prefs else ""
    _send(chat_id, (
        "Hola! Soy el bot de *Val. Semanal* 🏗\n\n"
        "Uso IA para responder tus consultas y *aprendo tu estilo*.\n\n"
        "*Puedes preguntarme lo que sea:*\n"
        "• _Dame el analisis de mara_\n"
        "• _Cuanto se gasto vs lo valorizado en alma mater?_\n"
        "• _Cuales obras estan en perdida?_\n"
        "• _Comparame cenepa con biomedicas_\n"
        "• _Resumen general de todas las obras_\n\n"
        "*Memoria:*\n"
        "• Recuerdo el contexto de la conversacion\n"
        "• Si me corriges el formato, lo aprendo para siempre\n"
        "• /recordar REGLA - Guardar una preferencia\n"
        "• /preferencias - Ver reglas guardadas\n"
        "• /olvidar - Borrar todas las preferencias\n\n"
        "*Comandos rapidos:*\n"
        "/resumen /montos /obra NOMBRE /costos NOMBRE\n"
        f"{prefs_txt}"
    ))


def _handle_recordar(chat_id, regla):
    """Guarda una preferencia de estilo explicitamente."""
    if not regla:
        return _send(chat_id, "Indica que quieres que recuerde.\nEj: /recordar siempre mostrar montos en miles")
    reglas_lower = [r.lower() for r in _preferencias]
    if regla.lower() in reglas_lower:
        return _send(chat_id, f"Ya tengo esa regla guardada.")
    _preferencias.append(regla)
    if len(_preferencias) > MAX_PREFERENCIAS:
        _preferencias.pop(0)
    ok = _guardar_preferencias()
    if ok:
        _send(chat_id, f"✅ Guardado: _{regla}_\n\nAhora tengo {len(_preferencias)} preferencias.")
    else:
        _send(chat_id, f"Guardado en memoria pero no pude persistir en el repo.")


def _handle_preferencias(chat_id):
    """Muestra las preferencias guardadas."""
    if not _preferencias:
        return _send(chat_id, "No tengo preferencias guardadas aun.\n\nEnseñame tu estilo corrigiendome o usa:\n/recordar REGLA")
    msg = f"📝 *Preferencias guardadas* ({len(_preferencias)}):\n\n"
    for i, p in enumerate(_preferencias, 1):
        msg += f"{i}. {p}\n"
    msg += "\n_Usa /olvidar para borrar todas._"
    _send(chat_id, msg)


def _handle_olvidar(chat_id):
    """Borra todas las preferencias."""
    _preferencias.clear()
    _guardar_preferencias()
    _send(chat_id, "🗑 Preferencias borradas. Empezamos de cero.")


def _handle_resumen(chat_id):
    data = _cargar_resumen()
    if not data or not data.get("reportes"):
        return _send(chat_id, "No hay reportes procesados aun.")

    msg = f"*Resumen Val. Semanal*\n_Actualizado: {data.get('ultima_actualizacion', '?')}_\n\n"
    for obra_key, r in sorted(data["reportes"].items()):
        a = r.get("analisis", {}).get("total", {})
        estado = a.get("estado", "")
        var = a.get("variacion", 0)
        emoji = "✅" if estado == "GANANCIA" else ("🔴" if estado == "PERDIDA" else "⚪")
        curva_txt = ""
        if r.get("tiene_curva"):
            curva_txt = f" | Curva: {r.get('curva_estado', '')}"
        msg += f"{emoji} *{obra_key}* | {r.get('fecha_reporte', '?')}\n"
        msg += f"  Val: {_fmt(r.get('total_valorizacion', 0))}"
        if estado:
            msg += f" | {'+' if var >= 0 else ''}{_fmt(var)} {estado}"
        msg += f"{curva_txt}\n\n"
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
    """Detalle rapido de una obra (slash command)."""
    data = _cargar_resumen()
    if not data:
        return _send(chat_id, "Error cargando datos.")
    nombre_lower = nombre.strip().lower()
    ALIASES = {
        "beethoven": "BEETHOVEN", "btv": "BEETHOVEN",
        "alma mater": "ALMA MATER", "mater": "ALMA MATER", "alma": "ALMA MATER",
        "mara": "MARA", "cema": "MARA",
        "cenepa": "CENEPA", "heroes": "CENEPA",
        "biomedicas": "BIOMEDICAS", "biomedica": "BIOMEDICAS", "bio": "BIOMEDICAS",
        "roosevelt": "ROOSEVELT", "rooselvet": "ROOSEVELT",
    }
    obra_key = ALIASES.get(nombre_lower, nombre.strip().upper())
    for a, k in ALIASES.items():
        if a in nombre_lower:
            obra_key = k
            break

    r = data.get("reportes", {}).get(obra_key)
    if not r:
        obras = ", ".join(data.get("reportes", {}).keys())
        return _send(chat_id, f"No encontre '{nombre}'.\nObras: {obras}")

    msg = f"📋 *{obra_key}* | {r.get('fecha_reporte', '?')}\n\n"
    msg += f"*1. CORTE DE VALORIZACION*\n"
    msg += f"  Costo Directo: {_fmt(r.get('costo_directo', 0))}\n"
    gg_p = r.get('gg_percent', 0)
    msg += f"  Gastos Generales ({gg_p:.1f}%): {_fmt(r.get('gastos_generales', 0))}\n"
    up = r.get('util_percent', 0)
    msg += f"  Utilidad ({up:.1f}%): {_fmt(r.get('utilidad', 0))}\n"
    msg += f"  Sub Total: {_fmt(r.get('sub_total', 0))}\n"
    msg += f"  IGV (18%): {_fmt(r.get('igv', 0))}\n"
    msg += f"  *TOTAL: {_fmt(r.get('total_valorizacion', 0))}*\n"

    analisis = r.get("analisis", {})
    if analisis:
        msg += f"\n*2. ANALISIS (Val vs Ejec)*\n"
        for sec, lbl in [("cd", "CD"), ("gg", "GG"), ("total", "TOTAL")]:
            s = analisis.get(sec, {})
            sv = s.get("variacion", 0)
            sp = s.get("variacion_pct", 0)
            se = s.get("estado", "")
            signo = "+" if sv >= 0 else ""
            emoji = "✅" if se == "GANANCIA" else "🔴"
            msg += f"  {emoji} {lbl}: {signo}{_fmt(sv)} ({sp:+.1f}%) {se}\n"

    if r.get("tiene_curva"):
        prog = r.get("curva_prog_acum_pct", 0) * 100
        ejec = r.get("curva_ejec_acum_pct", 0) * 100
        msg += f"\n*3. CURVA S:* Prog {prog:.1f}% | Ejec {ejec:.1f}% | {r.get('curva_estado', '')}\n"
    else:
        msg += f"\n*3. CURVA S:* _No disponible_\n"

    if r.get("drive_link"):
        msg += f"\n[Ver reporte]({r['drive_link']})"
    _send(chat_id, msg)


def _handle_costos(chat_id, nombre):
    """Gastos ejecutados (slash command)."""
    data = _cargar_resumen()
    if not data:
        return _send(chat_id, "Error cargando datos.")
    ALIASES = {
        "beethoven": "BEETHOVEN", "btv": "BEETHOVEN",
        "alma mater": "ALMA MATER", "mater": "ALMA MATER", "alma": "ALMA MATER",
        "mara": "MARA", "cema": "MARA",
        "cenepa": "CENEPA", "heroes": "CENEPA",
        "biomedicas": "BIOMEDICAS", "biomedica": "BIOMEDICAS", "bio": "BIOMEDICAS",
        "roosevelt": "ROOSEVELT", "rooselvet": "ROOSEVELT",
    }
    nombre_lower = nombre.strip().lower()
    obra_key = ALIASES.get(nombre_lower, nombre.strip().upper())
    for a, k in ALIASES.items():
        if a in nombre_lower:
            obra_key = k
            break

    r = data.get("reportes", {}).get(obra_key)
    if not r:
        return _send(chat_id, f"No encontre '{nombre}'.")

    def fp(n):
        return f"{n:.1f}%" if n else "0.0%"

    msg = f"💰 *Gastos Ejecutados - {obra_key}* | {r.get('fecha_reporte', '?')}\n\n"
    msg += f"*COSTO DIRECTO:*\n"
    msg += f"  Personal Obrero: {_fmt(r.get('personal_obrero', 0))} ({fp(r.get('personal_obrero_pct', 0))})\n"
    msg += f"  Materiales: {_fmt(r.get('materiales', 0))} ({fp(r.get('materiales_pct', 0))})\n"
    msg += f"  Alquileres: {_fmt(r.get('alquileres', 0))} ({fp(r.get('alquileres_pct', 0))})\n"
    msg += f"  Subcontratos: {_fmt(r.get('subcontratos', 0))} ({fp(r.get('subcontratos_pct', 0))})\n"
    msg += f"  Costos Varios: {_fmt(r.get('costos_varios', 0))} ({fp(r.get('costos_varios_pct', 0))})\n"
    msg += f"  *TOTAL CD: {_fmt(r.get('total_cd_ejecutado', 0))}*\n\n"
    msg += f"*GASTOS GENERALES:*\n"
    msg += f"  Planilla Staff: {_fmt(r.get('planilla_staff', 0))} ({fp(r.get('planilla_staff_pct', 0))})\n"
    msg += f"  Otros GG: {_fmt(r.get('otros_gg', 0))} ({fp(r.get('otros_gg_pct', 0))})\n"
    msg += f"  *TOTAL GG: {_fmt(r.get('total_gg_ejecutado', 0))}*\n"
    _send(chat_id, msg)


# ============================================================
# HANDLER PRINCIPAL - TODO texto libre va a IA
# ============================================================

def _handle_text(chat_id, texto):
    """Envia la pregunta a IA con historial y preferencias."""
    data = _cargar_resumen()

    if not data or not data.get("reportes"):
        return _send(chat_id, "No hay datos de reportes cargados.")

    # Enviar a IA con historial
    respuesta = _ask_ai(texto, data, chat_id=chat_id)

    if respuesta:
        # Extraer y guardar reglas de estilo si las hay
        respuesta = _extraer_regla(respuesta)

        # Guardar en historial
        _agregar_historial(chat_id, "user", texto)
        _agregar_historial(chat_id, "assistant", respuesta)

        # Truncar si es muy largo para Telegram (max 4096 chars)
        if len(respuesta) > 4000:
            respuesta = respuesta[:3950] + "\n\n_...respuesta truncada_"
        _send(chat_id, respuesta)
    else:
        # Fallback: IA no disponible
        _send(chat_id, (
            "No pude procesar tu consulta.\n\n"
            "Prueba con comandos rapidos:\n"
            "/resumen - Resumen general\n"
            "/obra MARA - Detalle de obra\n"
            "/costos BEETHOVEN - Gastos ejecutados"
        ))


# ============================================================
# PROCESAMIENTO DE UPDATES
# ============================================================

def _process_update(update):
    """Procesa un update de Telegram."""
    message = update.get("message")
    if not message:
        return

    chat_id = message["chat"]["id"]
    text = message.get("text", "").strip()
    if not text:
        return

    now = datetime.now(PERU_TZ).strftime("%H:%M:%S")
    print(f"  [{now}] Chat {chat_id}: {text[:80]}")
    sys.stdout.flush()

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
            if arg:
                _handle_obra(chat_id, arg)
            else:
                _send(chat_id, "Indica la obra. Ej: /obra MARA")
        elif cmd in ("/costos", "/desglose"):
            if arg:
                _handle_costos(chat_id, arg)
            else:
                _send(chat_id, "Indica la obra. Ej: /costos BEETHOVEN")
        elif cmd == "/recordar":
            _handle_recordar(chat_id, arg)
        elif cmd == "/preferencias":
            _handle_preferencias(chat_id)
        elif cmd == "/olvidar":
            _handle_olvidar(chat_id)
        else:
            # Slash command desconocido -> tratar como texto
            _handle_text(chat_id, text[1:])
    else:
        # Todo texto libre va a IA
        _handle_text(chat_id, text)


# ============================================================
# LOOP PRINCIPAL
# ============================================================

def main():
    global _running

    print("=" * 55)
    print("  BOT TELEGRAM WORKER - Val. Semanal")
    print(f"  Motor: GitHub Models ({AI_MODEL}) + Memoria")
    print(f"  Inicio: {datetime.now(PERU_TZ).strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Max runtime: {MAX_RUNTIME_HOURS}h")
    print(f"  Hora fin: {HORA_FIN}:00 (Peru)")
    print("=" * 55)

    if not TELEGRAM_BOT_TOKEN:
        print("[ERROR] TELEGRAM_BOT_TOKEN no configurado.")
        sys.exit(1)

    if not GH_MODELS_TOKEN:
        print("[WARN] GH_MODELS_TOKEN no configurado. IA no disponible.")

    # Verificar bot
    try:
        resp = requests.get(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getMe",
            timeout=10
        )
        bot_info = resp.json()
        if bot_info.get("ok"):
            print(f"  Bot: @{bot_info['result'].get('username', '?')}")
        else:
            print(f"[ERROR] Token invalido")
            sys.exit(1)
    except Exception as e:
        print(f"[ERROR] No se pudo conectar: {e}")
        sys.exit(1)

    # Eliminar webhook
    try:
        requests.get(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/deleteWebhook",
            timeout=10
        )
    except Exception:
        pass

    # Verificar datos
    data = _cargar_resumen()
    if data and data.get("reportes"):
        print(f"  Datos: {len(data['reportes'])} obras")
    else:
        print("  [WARN] Sin datos en resumen.json")

    # Cargar preferencias persistentes
    _cargar_preferencias()

    # Verificar IA
    if GH_MODELS_TOKEN:
        token_type = GH_MODELS_TOKEN[:4] + "..."
        print(f"  Token AI: configurado ({len(GH_MODELS_TOKEN)} chars, tipo: {token_type})")
        sys.stdout.flush()

        # Test rapido con modelo ligero
        test_msgs = [{"role": "user", "content": "responde solo: OK"}]
        test_result = _call_model(test_msgs, AI_MODEL_FALLBACK)
        if test_result:
            print(f"  IA ({AI_MODEL_FALLBACK}): OK - respuesta: {test_result[:30]}")
        else:
            print(f"  [WARN] IA: no responde (ver errores arriba)")
            # Info extra para diagnostico
            alt_token = os.environ.get("GITHUB_TOKEN", "")
            if alt_token and alt_token != GH_MODELS_TOKEN:
                print(f"  [INFO] Probando con GITHUB_TOKEN...")
                test2 = _call_model(test_msgs, AI_MODEL_FALLBACK, token=alt_token)
                if test2:
                    print(f"  [INFO] GITHUB_TOKEN SI funciona! Respuesta: {test2[:30]}")
                else:
                    print(f"  [INFO] GITHUB_TOKEN tambien falla.")
        sys.stdout.flush()
    else:
        print(f"  [WARN] GH_MODELS_TOKEN NO CONFIGURADO")
        sys.stdout.flush()

    print(f"\n  Escuchando mensajes...\n")
    sys.stdout.flush()

    start_time = time.time()
    max_runtime_seconds = MAX_RUNTIME_HOURS * 3600
    offset = None
    msg_count = 0

    while _running:
        elapsed = time.time() - start_time
        if elapsed > max_runtime_seconds:
            print(f"\n  Tiempo maximo alcanzado ({MAX_RUNTIME_HOURS}h). Terminando.")
            break

        hora_peru = datetime.now(PERU_TZ).hour
        if hora_peru >= HORA_FIN:
            print(f"\n  Hora fin ({HORA_FIN}:00 Peru). Terminando.")
            break

        try:
            params = {"timeout": 25, "allowed_updates": ["message"]}
            if offset:
                params["offset"] = offset

            resp = requests.get(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates",
                params=params,
                timeout=30,
            )

            if resp.status_code != 200:
                time.sleep(5)
                continue

            result = resp.json()
            if not result.get("ok"):
                time.sleep(5)
                continue

            updates = result.get("result", [])
            for update in updates:
                offset = update["update_id"] + 1
                try:
                    _process_update(update)
                    msg_count += 1
                except Exception as e:
                    print(f"  [ERROR] {e}")

        except requests.exceptions.Timeout:
            continue
        except Exception as e:
            print(f"  [ERROR] {e}")
            time.sleep(5)

    elapsed_min = (time.time() - start_time) / 60
    print(f"\n  === RESUMEN ===")
    print(f"  Duracion: {elapsed_min:.0f} minutos")
    print(f"  Mensajes procesados: {msg_count}")
    print(f"  Preferencias: {len(_preferencias)}")
    print(f"  Fin: {datetime.now(PERU_TZ).strftime('%d/%m/%Y %H:%M:%S')}")


if __name__ == "__main__":
    main()
