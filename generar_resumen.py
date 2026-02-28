"""
Genera resumen.json re-procesando los Excel que ya estan en Drive.
Uso unico: para crear el resumen.json inicial.
Despues, el pipeline normal lo mantiene actualizado.
"""
import io
import sys
from datetime import datetime, timezone, timedelta

from rich.console import Console

from config import OBRA_FOLDER_IDS, MONTH_ABBREVS, MONTH_NAMES
from auth_gmail import autenticar_drive
import agente_excel
import resumen_data

PERU_TZ = timezone(timedelta(hours=-5))
console = Console()


def main():
    console.print("[bold cyan]Generando resumen.json desde archivos en Drive...[/bold cyan]")

    drive_service = autenticar_drive()
    now = datetime.now(PERU_TZ)
    month_idx = now.month - 1
    month_abbrev = MONTH_ABBREVS[month_idx]
    year_short = str(now.year)[-2:]
    month_num = now.month

    # Patron de carpeta: "2.Feb-26"
    expected_folder = f"{month_num}.{month_abbrev}-{year_short}"
    console.print(f"  Buscando carpeta: {expected_folder}")

    total = 0
    errores = 0

    for obra_key, parent_folder_id in OBRA_FOLDER_IDS.items():
        if not parent_folder_id or parent_folder_id == "PEGAR_FOLDER_ID_AQUI":
            continue

        console.print(f"\n[yellow]{obra_key}[/yellow]")

        try:
            # Buscar la carpeta del mes
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
            ).execute()

            folders = results.get("files", [])
            month_folder = None
            for f in folders:
                if month_abbrev.lower() in f["name"].lower():
                    month_folder = f
                    break

            if not month_folder:
                console.print(f"  [dim]No se encontro carpeta del mes.[/dim]")
                continue

            console.print(f"  Carpeta: {month_folder['name']}")

            # Buscar archivos Excel en la carpeta
            query = (
                f"'{month_folder['id']}' in parents "
                f"and trashed = false "
                f"and (name contains '.xlsx' or name contains '.xlsm')"
            )
            results = drive_service.files().list(
                q=query,
                fields="files(id, name, webViewLink)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()

            excel_files = results.get("files", [])
            if not excel_files:
                console.print(f"  [dim]No hay archivos Excel.[/dim]")
                continue

            # Procesar el primer Excel encontrado
            excel_file = excel_files[0]
            console.print(f"  Archivo: {excel_file['name']}")

            # Descargar
            request = drive_service.files().get_media(
                fileId=excel_file["id"],
                supportsAllDrives=True,
            )
            from googleapiclient.http import MediaIoBaseDownload
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()

            excel_bytes = fh.getvalue()
            console.print(f"  Descargado: {len(excel_bytes) / 1024:.0f} KB")

            # Procesar con agente_excel
            datos = agente_excel.procesar(excel_bytes)
            console.print(f"  CD: S/ {datos['rval']['costoDirecto']:,.2f}")
            console.print(f"  Total: S/ {datos['rval']['totalValorizacion']:,.2f}")

            # Buscar link del reporte HTML
            html_query = (
                f"'{month_folder['id']}' in parents "
                f"and trashed = false "
                f"and name contains '.html'"
            )
            html_results = drive_service.files().list(
                q=html_query,
                fields="files(id, name, webViewLink)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()
            html_files = html_results.get("files", [])
            drive_link = html_files[0].get("webViewLink", "") if html_files else ""

            # Guardar en resumen
            resumen_data.guardar_reporte(
                obra_key, datos, drive_link, excel_file["name"]
            )
            total += 1
            console.print(f"  [green]OK[/green]")

        except Exception as e:
            errores += 1
            console.print(f"  [red]Error: {e}[/red]")
            import traceback
            traceback.print_exc()

    console.print(f"\n[bold green]Completado: {total} obras procesadas, {errores} errores[/bold green]")

    if total > 0:
        console.print(f"[green]resumen.json generado correctamente.[/green]")
        console.print("[dim]Recuerda hacer git add resumen.json && git commit && git push[/dim]")


if __name__ == "__main__":
    main()
