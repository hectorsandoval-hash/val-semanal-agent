"""
RESUMEN DATA - Persistencia COMPLETA de datos de reportes procesados
=====================================================================
Guarda un JSON con TODOS los datos de cada reporte procesado.
Este archivo es leido por el bot de Telegram para responder consultas.

Datos guardados por cada obra:
  - Seccion 1: Corte de Valorizacion (CD, GG, Utilidad, SubTotal, IGV, Total)
  - Seccion 2: Gastos Ejecutados - Costo Directo (desglose con %)
  - Seccion 3: Gastos Generales Ejecutados (desglose con %)
  - Seccion 4: Analisis Comparativo (Val vs Ejec, variacion, ganancia/perdida)
  - Seccion 5: Curva S completa (programado, ejecutado, planificado acumulados)
"""
import json
import os
from datetime import datetime, timezone, timedelta

PERU_TZ = timezone(timedelta(hours=-5))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RESUMEN_FILE = os.path.join(BASE_DIR, "resumen.json")


def cargar_resumen():
    """Carga el resumen existente o retorna estructura vacia."""
    if os.path.exists(RESUMEN_FILE):
        try:
            with open(RESUMEN_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass

    return {
        "ultima_actualizacion": "",
        "reportes": {},
    }


def guardar_reporte(obra_key, datos, drive_link="", filename=""):
    """
    Agrega o actualiza TODOS los datos de un reporte procesado al resumen.

    Args:
        obra_key: clave de la obra (ej: 'BEETHOVEN')
        datos: dict con los datos extraidos del Excel (de agente_excel.procesar())
        drive_link: link al reporte en Drive
        filename: nombre del archivo HTML
    """
    resumen = cargar_resumen()
    now = datetime.now(PERU_TZ)

    rval = datos.get("rval", {})
    res_costo = datos.get("resCosto", {})
    curva = datos.get("curva")

    # === SECCION 1: CORTE DE VALORIZACION (RVAL) ===
    cd = rval.get("costoDirecto", 0)
    gg = rval.get("gastosGenerales", 0)
    gg_pct = rval.get("ggPercent", 0)
    util = rval.get("utilidad", 0)
    util_pct = rval.get("utilPercent", 0)
    sub_total = cd + gg + util
    igv = sub_total * 0.18
    total_con_igv = sub_total + igv
    total_val = rval.get("totalValorizacion", 0)
    # Si el total del Excel ya incluye IGV, usar ese; si no, usar calculado
    if total_val == 0:
        total_val = total_con_igv

    # === SECCION 2: GASTOS EJECUTADOS - COSTO DIRECTO (RES-COSTO) ===
    po = res_costo.get("personalObrero", 0)
    mat = res_costo.get("materiales", 0)
    alq = res_costo.get("alquileres", 0)
    sub = res_costo.get("subcontratos", 0)
    cv = res_costo.get("costosVarios", 0)
    total_cd_ejec = res_costo.get("totalCD", 0)
    if total_cd_ejec == 0:
        total_cd_ejec = po + mat + alq + sub + cv

    # Porcentajes de cada partida sobre total CD ejecutado
    cd_pcts = {}
    if total_cd_ejec > 0:
        cd_pcts = {
            "personal_obrero_pct": (po / total_cd_ejec) * 100,
            "materiales_pct": (mat / total_cd_ejec) * 100,
            "alquileres_pct": (alq / total_cd_ejec) * 100,
            "subcontratos_pct": (sub / total_cd_ejec) * 100,
            "costos_varios_pct": (cv / total_cd_ejec) * 100,
        }

    # === SECCION 3: GASTOS GENERALES EJECUTADOS (RES-COSTO) ===
    ps = res_costo.get("planillaStaff", 0)
    ogg = res_costo.get("otrosGG", 0)
    total_gg_ejec = res_costo.get("totalGG", 0)
    if total_gg_ejec == 0:
        total_gg_ejec = ps + ogg

    gg_pcts = {}
    if total_gg_ejec > 0:
        gg_pcts = {
            "planilla_staff_pct": (ps / total_gg_ejec) * 100,
            "otros_gg_pct": (ogg / total_gg_ejec) * 100,
        }

    # === SECCION 4: ANALISIS COMPARATIVO (Val vs Ejec) ===
    # CD: Valorizacion vs Ejecutado
    cd_var = cd - total_cd_ejec
    cd_var_pct = (cd_var / cd * 100) if cd > 0 else 0
    cd_estado = "GANANCIA" if cd_var > 0 else ("PERDIDA" if cd_var < 0 else "NEUTRO")

    # GG: Valorizacion vs Ejecutado
    gg_var = gg - total_gg_ejec
    gg_var_pct = (gg_var / gg * 100) if gg > 0 else 0
    gg_estado = "GANANCIA" if gg_var > 0 else ("PERDIDA" if gg_var < 0 else "NEUTRO")

    # Total: (CD+GG) Val vs (CD+GG) Ejec
    total_val_cdgg = cd + gg
    total_ejec_cdgg = total_cd_ejec + total_gg_ejec
    total_var = total_val_cdgg - total_ejec_cdgg
    total_var_pct = (total_var / total_val_cdgg * 100) if total_val_cdgg > 0 else 0
    total_estado = "GANANCIA" if total_var > 0 else ("PERDIDA" if total_var < 0 else "NEUTRO")

    reporte_data = {
        "obra": obra_key,
        "proyecto": datos.get("projectName", ""),
        "fecha_reporte": _format_date(datos.get("date")),
        "fecha_procesado": now.strftime("%d/%m/%Y %H:%M"),
        "filename": filename,
        "drive_link": drive_link,

        # --- SECCION 1: CORTE DE VALORIZACION ---
        "costo_directo": cd,
        "gastos_generales": gg,
        "gg_percent": gg_pct,
        "utilidad": util,
        "util_percent": util_pct,
        "sub_total": sub_total,
        "igv": igv,
        "igv_percent": 18.0,
        "total_valorizacion": total_val,

        # --- SECCION 2: GASTOS EJECUTADOS - COSTO DIRECTO ---
        "personal_obrero": po,
        "materiales": mat,
        "alquileres": alq,
        "subcontratos": sub,
        "costos_varios": cv,
        "total_cd_ejecutado": total_cd_ejec,
        **cd_pcts,

        # --- SECCION 3: GASTOS GENERALES EJECUTADOS ---
        "planilla_staff": ps,
        "otros_gg": ogg,
        "total_gg_ejecutado": total_gg_ejec,
        **gg_pcts,

        # --- SECCION 4: ANALISIS COMPARATIVO ---
        "analisis": {
            "cd": {
                "valorizacion": cd,
                "ejecutado": total_cd_ejec,
                "variacion": cd_var,
                "variacion_pct": cd_var_pct,
                "estado": cd_estado,
            },
            "gg": {
                "valorizacion": gg,
                "ejecutado": total_gg_ejec,
                "variacion": gg_var,
                "variacion_pct": gg_var_pct,
                "estado": gg_estado,
            },
            "total": {
                "valorizacion": total_val_cdgg,
                "ejecutado": total_ejec_cdgg,
                "variacion": total_var,
                "variacion_pct": total_var_pct,
                "estado": total_estado,
            },
        },

        # --- SECCION 5: CURVA S ---
        "tiene_curva": curva is not None,
    }

    # Guardar datos completos de curva si existe
    if curva:
        idx = curva.get("mesActualIndex", -1)
        contractual = curva.get("contractual", [])
        valorizado = curva.get("valorizado", [])
        proyectado = curva.get("proyectado")
        total_contractual = curva.get("total", 0)

        reporte_data["curva_total_contractual"] = total_contractual

        if idx >= 0 and idx < len(contractual) and idx < len(valorizado):
            c = contractual[idx]
            v = valorizado[idx]
            reporte_data["curva_prog_acum_pct"] = c.get("acumPct", 0)
            reporte_data["curva_prog_acum_monto"] = c.get("acumulado", 0)
            reporte_data["curva_ejec_acum_pct"] = v.get("acumPct", 0)
            reporte_data["curva_ejec_acum_monto"] = v.get("acumulado", 0)
            reporte_data["curva_mes_actual"] = c.get("mes", "")

            # Planificado si existe
            if proyectado and idx < len(proyectado):
                p = proyectado[idx]
                reporte_data["curva_plan_acum_pct"] = p.get("acumPct", 0)
                reporte_data["curva_plan_acum_monto"] = p.get("acumulado", 0)

            # Desvio
            prog_pct = c.get("acumPct", 0)
            ejec_pct = v.get("acumPct", 0)
            desvio = prog_pct - ejec_pct
            reporte_data["curva_desvio_pct"] = desvio
            if desvio < 0:
                reporte_data["curva_estado"] = "ADELANTADO"
            elif desvio > 0.01:
                reporte_data["curva_estado"] = "ATRASADO"
            else:
                reporte_data["curva_estado"] = "EN TIEMPO"

    resumen["reportes"][obra_key] = reporte_data
    resumen["ultima_actualizacion"] = now.strftime("%d/%m/%Y %H:%M")

    with open(RESUMEN_FILE, "w", encoding="utf-8") as f:
        json.dump(resumen, f, ensure_ascii=False, indent=2)

    print(f"  [RESUMEN] Datos guardados para {obra_key}")


def obtener_resumen_obra(obra_key):
    """Retorna los datos del ultimo reporte de una obra, o None."""
    resumen = cargar_resumen()
    return resumen.get("reportes", {}).get(obra_key)


def obtener_todas_las_obras():
    """Retorna dict con todos los reportes guardados."""
    resumen = cargar_resumen()
    return resumen.get("reportes", {})


def _format_date(dt):
    """Formatea datetime a string legible."""
    if not dt:
        return ""
    if isinstance(dt, datetime):
        return dt.strftime("%d/%m/%Y")
    return str(dt)
