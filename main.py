"""
ORQUESTADOR PRINCIPAL - Agente de Valorizacion Semanal
======================================================
Pipeline automatico 100% en la nube:
  1. Busca correos con Excel de valorizacion en Gmail
  2. Descarga el Excel y lo guarda en Google Drive (API)
  3. Procesa el Excel (extrae datos)
  4. Genera reporte HTML
  5. Guarda el reporte en Google Drive (API)
  6. Envia notificacion por Telegram

Disenado para ejecutarse en GitHub Actions (ubuntu-latest).
Tambien funciona localmente.

Uso:
  python main.py                  # Ejecutar pipeline completo
  python main.py --solo-buscar    # Solo buscar correos (sin procesar)
  python main.py --manual FILE    # Procesar un archivo Excel local
  python main.py --obra BEETHOVEN # Especificar obra (para modo --manual)
  python main.py --verificar      # Verificar Folder IDs de Drive
  python main.py --max 10         # Limitar busqueda a N correos
"""
import argparse
import os
import sys
import traceback
from datetime import datetime, timezone, timedelta

from email.utils import parsedate_to_datetime
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from config import REPORT_DIR, LOG_DIR, OBRA_FOLDER_IDS
from auth_gmail import autenticar_gmail, autenticar_drive, obtener_perfil
import agente_gmail
import agente_drive
import agente_excel
import agente_reporte
import agente_notificador
import resumen_data

# Zona horaria Peru (UTC-5)
PERU_TZ = timezone(timedelta(hours=-5))

console = Console()


def main():
    parser = argparse.ArgumentParser(description="Agente de Valorizacion Semanal")
    parser.add_argument("--solo-buscar", action="store_true",
                        help="Solo buscar correos sin procesar")
    parser.add_argument("--manual", type=str, default=None,
                        help="Procesar un archivo Excel local (ruta al archivo)")
    parser.add_argument("--obra", type=str, default=None,
                        help="Nombre de la obra (para modo --manual, ej: BEETHOVEN)")
    parser.add_argument("--verificar", action="store_true",
                        help="Verificar Folder IDs configurados en Drive")
    parser.add_argument("--max", type=int, default=20,
                        help="Numero maximo de correos a buscar (default: 20)")
    args = parser.parse_args()

    # Crear directorios locales si no existen (para logs y temp)
    os.makedirs(REPORT_DIR, exist_ok=True)
    os.makedirs(LOG_DIR, exist_ok=True)

    console.print(Panel.fit(
        "[bold cyan]AGENTE DE VALORIZACION SEMANAL[/bold cyan]\n"
        "[dim]100% Cloud - Google Drive API[/dim]\n"
        f"Fecha: {datetime.now(PERU_TZ).strftime('%d/%m/%Y %H:%M')}",
        border_style="cyan",
    ))

    # ========================================
    # MODO VERIFICAR: Verificar Folder IDs
    # ========================================
    if args.verificar:
        _verificar_folder_ids()
        return

    # ========================================
    # MODO MANUAL: Procesar archivo local
    # ========================================
    if args.manual:
        _procesar_manual(args.manual, args.obra)
        return

    # ========================================
    # MODO AUTOMATICO: Pipeline desde Gmail
    # ========================================
    console.print("\n[bold yellow]>>> AUTENTICACION[/bold yellow]")
    try:
        gmail_service = autenticar_gmail()
        drive_service = autenticar_drive()
        mi_email = obtener_perfil(gmail_service)
        console.print(f"  Conectado como: [green]{mi_email}[/green]")
        console.print(f"  Gmail API: [green]OK[/green]")
        console.print(f"  Drive API: [green]OK[/green]")
    except Exception as e:
        console.print(f"[bold red]Error de autenticacion: {e}[/bold red]")
        agente_notificador.enviar_error(f"Error de autenticacion: {e}")
        sys.exit(1)

    # === AGENTE 1: Buscar correos ===
    console.print("\n[bold yellow]>>> AGENTE 1: BUSQUEDA DE CORREOS[/bold yellow]")
    correos = agente_gmail.buscar_correos_valorizacion(gmail_service, max_results=args.max)

    if not correos:
        console.print("[dim]No se encontraron correos nuevos de valorizacion.[/dim]")
        console.print(Panel.fit(
            "[bold green]SIN CORREOS NUEVOS[/bold green]\n"
            "No hay valorizaciones pendientes de procesar.",
            border_style="green",
        ))
        return

    # Mostrar tabla de correos encontrados
    _mostrar_tabla_correos(correos)

    if args.solo_buscar:
        console.print("\n[green]Modo --solo-buscar: no se procesan los correos.[/green]")
        return

    # === PROCESAR CADA CORREO ===
    total_procesados = 0
    total_errores = 0
    detalles = []

    for i, correo in enumerate(correos, 1):
        console.print(f"\n[bold yellow]>>> PROCESANDO CORREO {i}/{len(correos)}[/bold yellow]")
        console.print(f"  Asunto: {correo['asunto'][:60]}")
        console.print(f"  De: {correo.get('de_email', correo['de'][:40])}")

        try:
            resultado = _procesar_correo(gmail_service, drive_service, correo)
            if resultado:
                total_procesados += 1
                detalles.append(f"{correo['obra_detectada'] or '?'}: {resultado['filename']}")
            else:
                total_errores += 1
                detalles.append(f"ERROR: {correo['asunto'][:40]}")
        except Exception as e:
            total_errores += 1
            detalles.append(f"ERROR: {correo['asunto'][:30]} - {e}")
            console.print(f"  [bold red]Error: {e}[/bold red]")
            traceback.print_exc()

    # === RESUMEN FINAL ===
    console.print(Panel.fit(
        f"[bold green]PROCESO COMPLETADO[/bold green]\n"
        f"Correos encontrados: {len(correos)}\n"
        f"Reportes generados: {total_procesados}\n"
        f"Errores: {total_errores}",
        border_style="green" if total_errores == 0 else "yellow",
    ))

    # Enviar resumen por Telegram
    if total_procesados > 0 or total_errores > 0:
        agente_notificador.enviar_resumen(total_procesados, total_errores, detalles)


def _procesar_correo(gmail_service, drive_service, correo):
    """
    Procesa un correo individual: descarga Excel, guarda en Drive,
    genera reporte HTML, guarda en Drive, notifica.

    Returns:
        dict con 'filename' y 'drive_link', o None si falla
    """
    obra_key = correo["obra_detectada"]

    # Descargar primer adjunto Excel
    if not correo["adjuntos"]:
        console.print("  [red]Sin adjuntos Excel.[/red]")
        return None

    adjunto = correo["adjuntos"][0]
    console.print(f"  [AGENTE 2] Descargando: {adjunto['filename']}")

    excel_bytes = agente_gmail.descargar_adjunto_excel(
        gmail_service, correo["id"], adjunto["attachmentId"]
    )
    if not excel_bytes:
        console.print("  [red]Error descargando adjunto.[/red]")
        return None

    # Si no se detecto la obra del remitente/asunto, intentar desde el Excel
    if not obra_key:
        console.print("  [DETECCION] Detectando obra desde contenido del Excel...")
        project_name = agente_excel.detect_obra_name(excel_bytes)
        if project_name:
            obra_key = agente_gmail.detectar_obra_de_texto(project_name)
            console.print(f"  [DETECCION] Obra detectada: {obra_key or 'NO DETECTADA'}")

    if not obra_key:
        console.print("  [yellow]No se pudo detectar la obra. Saltando...[/yellow]")
        console.print(f"  [yellow]Asunto: {correo['asunto'][:60]}[/yellow]")
        return None

    # Verificar que la obra tiene Folder ID configurado
    folder_id = OBRA_FOLDER_IDS.get(obra_key)
    if not folder_id or folder_id == "PEGAR_FOLDER_ID_AQUI":
        console.print(f"  [yellow]Obra '{obra_key}' no tiene Folder ID configurado.[/yellow]")
        return None

    # Determinar mes
    month_name = agente_drive.obtener_mes_actual()
    console.print(f"  Obra: [cyan]{obra_key}[/cyan] | Mes: [cyan]{month_name}[/cyan]")

    # === AGENTE 2: Guardar Excel en Google Drive (API) ===
    console.print(f"  [AGENTE 2] Guardando Excel en Google Drive...")
    excel_result = agente_drive.guardar_excel(
        drive_service, excel_bytes, obra_key, month_name, adjunto["filename"]
    )
    if excel_result:
        console.print(f"  [green]Excel guardado en Drive: {excel_result['file_id'][:20]}...[/green]")
    else:
        console.print(f"  [yellow]No se pudo guardar Excel en Drive (continuando...).[/yellow]")

    # === AGENTE 3: Procesar Excel ===
    console.print(f"  [AGENTE 3] Procesando Excel...")
    try:
        datos = agente_excel.procesar(excel_bytes)
        console.print(f"  [green]Excel procesado correctamente.[/green]")
        console.print(f"    CD: S/ {datos['rval']['costoDirecto']:,.2f}")
        console.print(f"    GG: S/ {datos['rval']['gastosGenerales']:,.2f}")
    except Exception as e:
        console.print(f"  [red]Error procesando Excel: {e}[/red]")
        return None

    # === FIX: Usar fecha del correo (Gmail) en vez de la fecha del Excel ===
    # La fecha del Excel puede estar desactualizada; la del correo es mas confiable
    fecha_correo = correo.get("fecha", "")
    if fecha_correo:
        try:
            fecha_dt = parsedate_to_datetime(fecha_correo)
            # Convertir a hora Peru si tiene timezone
            if fecha_dt.tzinfo:
                fecha_dt = fecha_dt.astimezone(PERU_TZ)
            datos["date"] = fecha_dt
            console.print(f"  [green]Fecha reporte (del correo): {fecha_dt.strftime('%d/%m/%Y')}[/green]")
        except Exception:
            console.print(f"  [yellow]No se pudo parsear fecha del correo, usando fecha del Excel.[/yellow]")

    # === AGENTE 4: Generar reporte HTML ===
    console.print(f"  [AGENTE 4] Generando reporte HTML...")
    try:
        html_content, filename = agente_reporte.generar(datos)
        console.print(f"  [green]Reporte generado: {filename}[/green]")
    except Exception as e:
        console.print(f"  [red]Error generando reporte: {e}[/red]")
        return None

    # === AGENTE 5: Guardar reporte HTML en Google Drive (API) ===
    console.print(f"  [AGENTE 5] Guardando reporte en Google Drive...")
    report_result = agente_drive.guardar_reporte(
        drive_service, html_content, obra_key, month_name, filename
    )

    drive_link = ""
    if report_result:
        drive_link = report_result.get("web_link", "")
        console.print(f"  [green]Reporte guardado en Drive: {report_result['file_id'][:20]}...[/green]")
        if drive_link:
            console.print(f"  [green]Link: {drive_link}[/green]")
    else:
        console.print(f"  [yellow]No se pudo guardar reporte en Drive.[/yellow]")

    # Guardar copia local (util para artefactos de GitHub Actions)
    local_path = os.path.join(REPORT_DIR, filename)
    try:
        with open(local_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        console.print(f"  [dim]Copia local: {local_path}[/dim]")
    except Exception:
        pass  # No es critico

    # === Guardar datos en resumen.json (para bot de Telegram) ===
    try:
        resumen_data.guardar_reporte(obra_key, datos, drive_link, filename)
    except Exception as e:
        console.print(f"  [yellow]Error guardando resumen (no critico): {e}[/yellow]")

    # === AGENTE 6: Notificar por Telegram ===
    console.print(f"  [AGENTE 6] Enviando notificacion...")
    try:
        agente_notificador.enviar_notificacion(obra_key, month_name, drive_link, filename)
    except Exception as e:
        console.print(f"  [yellow]Notificacion fallida (no critico): {e}[/yellow]")

    # Marcar correo como procesado
    agente_gmail.marcar_procesado(gmail_service, correo)

    return {"filename": filename, "drive_link": drive_link}


def _procesar_manual(filepath, obra_override=None):
    """
    Modo manual: procesa un archivo Excel local.
    Genera el reporte y lo guarda en Google Drive via API.
    """
    console.print(f"\n[bold yellow]>>> MODO MANUAL[/bold yellow]")
    console.print(f"  Archivo: {filepath}")

    if not os.path.exists(filepath):
        console.print(f"[bold red]Archivo no encontrado: {filepath}[/bold red]")
        sys.exit(1)

    with open(filepath, "rb") as f:
        excel_bytes = f.read()

    console.print(f"  Tamanio: {len(excel_bytes) / 1024:.0f} KB")

    # Detectar obra
    obra_key = obra_override
    if obra_key:
        obra_key = obra_key.upper()
    else:
        project_name = agente_excel.detect_obra_name(excel_bytes)
        if project_name:
            obra_key = agente_gmail.detectar_obra_de_texto(project_name)
        console.print(f"  Obra detectada: {obra_key or 'NO DETECTADA'}")

    # Procesar
    console.print(f"  Procesando Excel...")
    try:
        datos = agente_excel.procesar(excel_bytes)
        console.print(f"  [green]Excel procesado correctamente.[/green]")
        console.print(f"    Proyecto: {datos['projectName']}")
        console.print(f"    CD: S/ {datos['rval']['costoDirecto']:,.2f}")
        console.print(f"    GG: S/ {datos['rval']['gastosGenerales']:,.2f}")
    except Exception as e:
        console.print(f"  [bold red]Error: {e}[/bold red]")
        traceback.print_exc()
        sys.exit(1)

    # Generar reporte
    console.print(f"  Generando reporte HTML...")
    html_content, filename = agente_reporte.generar(datos)
    console.print(f"  [green]Reporte generado: {filename}[/green]")

    # Guardar en Google Drive (si se detecto obra y hay Folder ID)
    drive_link = ""
    if obra_key and obra_key in OBRA_FOLDER_IDS:
        folder_id = OBRA_FOLDER_IDS[obra_key]
        if folder_id and folder_id != "PEGAR_FOLDER_ID_AQUI":
            try:
                drive_service = autenticar_drive()
                month_name = agente_drive.obtener_mes_actual()

                # Guardar Excel en Drive
                console.print(f"  Guardando Excel en Drive ({obra_key}/{month_name})...")
                agente_drive.guardar_excel(
                    drive_service, excel_bytes, obra_key, month_name,
                    os.path.basename(filepath)
                )

                # Guardar reporte HTML en Drive
                console.print(f"  Guardando reporte en Drive ({obra_key}/{month_name})...")
                report_result = agente_drive.guardar_reporte(
                    drive_service, html_content, obra_key, month_name, filename
                )
                if report_result:
                    drive_link = report_result.get("web_link", "")
                    console.print(f"  [green]Guardado en Drive: {drive_link or report_result['file_id']}[/green]")
            except Exception as e:
                console.print(f"  [yellow]No se pudo guardar en Drive: {e}[/yellow]")

    # Guardar copia local
    local_path = os.path.join(REPORT_DIR, filename)
    with open(local_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    console.print(f"  [green]Copia local: {local_path}[/green]")

    console.print(Panel.fit(
        f"[bold green]PROCESO MANUAL COMPLETADO[/bold green]\n"
        f"Obra: {obra_key or 'No detectada'}\n"
        f"Archivo: {filename}\n"
        f"Drive: {drive_link or 'N/A'}\n"
        f"Local: {local_path}",
        border_style="green",
    ))


def _verificar_folder_ids():
    """Verifica cuales obras tienen Folder ID configurado."""
    console.print("\n[bold yellow]>>> VERIFICACION DE FOLDER IDS[/bold yellow]")

    resultado = agente_drive.verificar_folder_ids()

    table = Table(title="ESTADO DE FOLDER IDS EN GOOGLE DRIVE", show_lines=True)
    table.add_column("Obra", style="bold white", width=15)
    table.add_column("Estado", width=15)
    table.add_column("Folder ID", max_width=50)

    for obra_key, info in resultado.items():
        if info["configurado"]:
            estado = "[green]CONFIGURADO[/green]"
        else:
            estado = "[red]PENDIENTE[/red]"
        table.add_row(obra_key, estado, info["folder_id"][:45] + "..." if len(info["folder_id"]) > 45 else info["folder_id"])

    console.print(table)

    total_ok = sum(1 for info in resultado.values() if info["configurado"])
    total = len(resultado)

    if total_ok == total:
        console.print(f"\n[green]Todas las obras ({total}) tienen Folder ID configurado.[/green]")
    else:
        console.print(f"\n[yellow]{total_ok}/{total} obras configuradas.[/yellow]")
        console.print("[yellow]Configura los Folder IDs faltantes en config.py[/yellow]")


def _mostrar_tabla_correos(correos):
    """Muestra tabla resumen de correos encontrados."""
    table = Table(title="CORREOS DE VALORIZACION ENCONTRADOS", show_lines=True)
    table.add_column("#", style="cyan", width=4)
    table.add_column("Fecha", style="white", width=12)
    table.add_column("Asunto", style="bold white", max_width=45)
    table.add_column("De", style="yellow", max_width=30)
    table.add_column("Obra", style="green", width=15)
    table.add_column("Excel", style="cyan", max_width=25)

    for i, correo in enumerate(correos, 1):
        adj_names = correo["adjuntos"][0]["filename"][:25] if correo["adjuntos"] else "-"

        table.add_row(
            str(i),
            correo["fecha"][:16] if correo["fecha"] else "-",
            correo["asunto"][:45],
            correo.get("de_email", correo["de"][:30]),
            correo["obra_detectada"] or "[red]??[/red]",
            adj_names,
        )

    console.print(table)


if __name__ == "__main__":
    main()
