"""
AGENTE NOTIFICADOR - Notificaciones por Telegram
=================================================
Envia notificaciones cuando se completa el procesamiento
de una valorizacion semanal.
"""
import requests
from config import TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID


def enviar_notificacion(obra_name, month_name, drive_link, filename=""):
    """
    Envia notificacion de reporte generado por Telegram.

    Args:
        obra_name: nombre de la obra (ej: 'BEETHOVEN')
        month_name: mes del reporte (ej: 'Febrero')
        drive_link: link al archivo en Drive
        filename: nombre del archivo HTML generado
    """
    mensaje = (
        f"\u2705 *Reporte Valorizacion Semanal*\n"
        f"\n"
        f"\U0001f4cb *Obra:* {obra_name}\n"
        f"\U0001f4c5 *Mes:* {month_name}\n"
    )

    if filename:
        mensaje += f"\U0001f4c4 *Archivo:* `{filename}`\n"

    if drive_link:
        mensaje += f"\n\U0001f4c1 [Ver en Drive]({drive_link})\n"

    mensaje += f"\n_Listo para revision._"

    return _enviar_telegram(mensaje)


def enviar_resumen(total_procesados, total_errores, detalles=None):
    """
    Envia resumen de ejecucion completa.

    Args:
        total_procesados: cantidad de reportes generados exitosamente
        total_errores: cantidad de errores
        detalles: lista de strings con detalles de cada procesamiento
    """
    if total_procesados == 0 and total_errores == 0:
        return True  # No hay nada que reportar

    emoji = "\u2705" if total_errores == 0 else "\u26a0\ufe0f"
    mensaje = (
        f"{emoji} *Resumen Val. Semanal*\n"
        f"\n"
        f"\u2705 Procesados: {total_procesados}\n"
        f"\u274c Errores: {total_errores}\n"
    )

    if detalles:
        mensaje += "\n*Detalle:*\n"
        for det in detalles[:10]:  # Max 10 detalles
            mensaje += f"  \u2022 {det}\n"

    return _enviar_telegram(mensaje)


def enviar_error(mensaje_error):
    """
    Envia notificacion de error critico.

    Args:
        mensaje_error: descripcion del error
    """
    mensaje = (
        f"\U0001f6a8 *Error en Val. Semanal Agent*\n"
        f"\n"
        f"`{mensaje_error}`\n"
        f"\n"
        f"_Revisa los logs para mas detalle._"
    )
    return _enviar_telegram(mensaje)


# ============================================================
# FUNCION INTERNA
# ============================================================

def _enviar_telegram(mensaje):
    """
    Envia un mensaje a Telegram via Bot API.

    Returns:
        True si se envio correctamente, False si fallo.
    """
    if TELEGRAM_BOT_TOKEN == "TU_BOT_TOKEN_AQUI" or TELEGRAM_CHAT_ID == "TU_CHAT_ID_AQUI":
        print("  [TELEGRAM] Bot no configurado. Mensaje no enviado.")
        # Imprimir sin emojis para evitar UnicodeEncodeError en Windows (cp1252)
        try:
            print(f"  [TELEGRAM] Mensaje: {mensaje[:100]}...")
        except UnicodeEncodeError:
            safe_msg = mensaje[:100].encode("ascii", errors="replace").decode("ascii")
            print(f"  [TELEGRAM] Mensaje: {safe_msg}...")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": mensaje,
        "parse_mode": "Markdown",
        "disable_web_page_preview": False,
    }

    try:
        response = requests.post(url, json=payload, timeout=10)
        if response.status_code == 200:
            print("  [TELEGRAM] Notificacion enviada correctamente.")
            return True
        else:
            print(f"  [TELEGRAM] Error {response.status_code}: {response.text[:100]}")
            return False
    except Exception as e:
        print(f"  [TELEGRAM] Error enviando: {e}")
        return False
