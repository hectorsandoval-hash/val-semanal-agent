"""
AGENTE REPORTE - Generacion de reporte HTML de valorizacion semanal
====================================================================
Porta la logica de generacion de reporte del index.html (JavaScript)
a Python. Genera un HTML standalone con dos paginas A4:
  - Pagina 1: Valorizacion, gastos ejecutados, analisis comparativo
  - Pagina 2: Curva S (SVG), tarjetas resumen, tabla de avance

Funcion publica principal:
    generar(data) -> (html_string, filename)

Donde `data` es el dict producido por agente_excel.procesar().
"""
import math
import re
from datetime import datetime, timedelta


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def fmt(n):
    """Formatea un numero con comas de miles y 2 decimales. Ej: 1,234,567.89"""
    if n is None:
        return "0.00"
    try:
        num = float(n)
    except (ValueError, TypeError):
        return "0.00"
    fixed = f"{abs(num):.2f}"
    parts = fixed.split(".")
    # Agregar comas de miles
    integer_part = parts[0]
    result = ""
    for i, ch in enumerate(reversed(integer_part)):
        if i > 0 and i % 3 == 0:
            result = "," + result
        result = ch + result
    return ("-" if num < 0 else "") + result + "." + parts[1]


def fmt_pct(n):
    """Formatea un decimal como porcentaje. Ej: 0.1234 -> '12.34%'"""
    if n is None:
        return "0.00%"
    try:
        num = float(n)
    except (ValueError, TypeError):
        return "0.00%"
    return f"{num * 100:.2f}%"


def format_date(d):
    """Formatea una fecha en espanol. Ej: '22 de Febrero de 2026'"""
    if not d:
        return ""
    months = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre",
    ]
    if isinstance(d, datetime):
        return f"{d.day} de {months[d.month - 1]} de {d.year}"
    # Handle Excel serial date number
    if isinstance(d, (int, float)) and 40000 < d < 60000:
        date = datetime(1899, 12, 30) + timedelta(days=int(d))
        return f"{date.day} de {months[date.month - 1]} de {date.year}"
    return str(d)


def format_date_short(d):
    """Formatea una fecha corta para el nombre de archivo. Ej: '22-Feb-2026'"""
    if not d:
        return ""
    months_short = [
        "Ene", "Feb", "Mar", "Abr", "May", "Jun",
        "Jul", "Ago", "Set", "Oct", "Nov", "Dic",
    ]
    if isinstance(d, datetime):
        return f"{d.day:02d}-{months_short[d.month - 1]}-{d.year}"
    if isinstance(d, (int, float)) and 40000 < d < 60000:
        date = datetime(1899, 12, 30) + timedelta(days=int(d))
        return f"{date.day:02d}-{months_short[date.month - 1]}-{date.year}"
    return str(d)


def get_short_name(full_name):
    """Obtiene nombre corto de la obra."""
    names = [
        "ALMA MATER", "MARA", "CENEPA", "BEETHOVEN",
        "BIOMEDICAS", "BIOMEDIC", "FRANKLIN", "ROOSEVELT",
    ]
    upper = (full_name or "").upper()
    for n in names:
        if n in upper:
            return n
    return full_name[:30] if full_name else "PROYECTO"


# ============================================================
# GENERATE CARD
# ============================================================

def generate_card(title, pct, monto):
    """Genera el HTML de una tarjeta de variacion."""
    is_positive = monto >= 0
    value_class = "positivo" if is_positive else "negativo"
    monto_class = "ganancia" if is_positive else "perdida"
    sign = "+" if is_positive else ""
    monto_sign = "+" if is_positive else "-"
    return f"""
    <div class="card">
        <div class="card-title">{title}</div>
        <div class="card-value {value_class}">{sign}{pct:.2f}%</div>
        <div class="card-monto {monto_class}">{monto_sign}S/ {fmt(abs(monto))}</div>
    </div>"""


# ============================================================
# PAGE 1: REPORTE DE VALORIZACION
# ============================================================

def generate_page1(data):
    """Genera la pagina 1 del reporte: valorizacion y analisis comparativo."""
    rc = data["resCosto"]
    rv = data["rval"]

    # Section 1: Corte de Valorizacion
    sub_total = rv["costoDirecto"] + rv["gastosGenerales"] + rv["utilidad"]
    igv = sub_total * 0.18
    total_con_igv = sub_total + igv

    # Section 4: Analisis Comparativo
    var_cd = rv["costoDirecto"] - rc["totalCD"]
    var_gg = rv["gastosGenerales"] - rc["totalGG"]
    var_total = var_cd + var_gg
    total_valorizado = rv["costoDirecto"] + rv["gastosGenerales"]
    total_ejecutado = rc["totalCD"] + rc["totalGG"]

    pct_var_cd = (var_cd / rc["totalCD"]) * 100 if rc["totalCD"] != 0 else 0
    pct_var_gg = (var_gg / rc["totalGG"]) * 100 if rc["totalGG"] != 0 else 0
    pct_var_total = (var_total / total_ejecutado) * 100 if total_ejecutado != 0 else 0

    # CD items
    cd_items = [
        {"name": "Costo de Materiales", "value": rc["materiales"]},
        {"name": "Costo de Alquileres", "value": rc["alquileres"]},
        {"name": "Costo de Subcontratos", "value": rc["subcontratos"]},
        {"name": "Costo Varios", "value": rc["costosVarios"]},
        {"name": "Costo Personal Obrero", "value": rc["personalObrero"]},
    ]

    # GG items
    gg_items = [
        {"name": "Planilla Staff", "value": rc["planillaStaff"]},
        {"name": "Otros Gastos Generales", "value": rc["otrosGG"]},
    ]

    # Build CD rows
    cd_rows_html = ""
    for it in cd_items:
        pct = f"{(it['value'] / rc['totalCD']) * 100:.2f}%" if rc["totalCD"] > 0 else "0.00%"
        cd_rows_html += f'<tr><td>{it["name"]}</td><td class="num">{fmt(it["value"])}</td><td class="num">{pct}</td></tr>'

    # Build GG rows
    gg_rows_html = ""
    for it in gg_items:
        pct = f"{(it['value'] / rc['totalGG']) * 100:.2f}%" if rc["totalGG"] > 0 else "0.00%"
        gg_rows_html += f'<tr><td>{it["name"]}</td><td class="num">{fmt(it["value"])}</td><td class="num">{pct}</td></tr>'

    # Comparative table helper
    def comp_row(concepto, val_monto, ejec_monto, var_monto, pct_var, is_total=False):
        cls_var = "valor-positivo" if var_monto >= 0 else "valor-negativo"
        sign = "+" if var_monto >= 0 else ""
        estado_cls = "estado-ganancia" if var_monto >= 0 else "estado-perdida"
        estado_txt = "GANANCIA" if var_monto >= 0 else "P&Eacute;RDIDA"
        if is_total:
            return (
                f'<tr class="total-row">'
                f'<td><strong>{concepto}</strong></td>'
                f'<td class="num"><strong>{fmt(val_monto)}</strong></td>'
                f'<td class="num"><strong>{fmt(ejec_monto)}</strong></td>'
                f'<td class="num {cls_var}"><strong>{sign}{fmt(var_monto)}</strong></td>'
                f'<td class="num {cls_var}"><strong>{sign}{pct_var:.2f}%</strong></td>'
                f'<td style="text-align:center"><span class="estado-box {estado_cls}">{estado_txt}</span></td>'
                f'</tr>'
            )
        return (
            f'<tr>'
            f'<td>{concepto}</td>'
            f'<td class="num">{fmt(val_monto)}</td>'
            f'<td class="num">{fmt(ejec_monto)}</td>'
            f'<td class="num {cls_var}">{sign}{fmt(var_monto)}</td>'
            f'<td class="num {cls_var}">{sign}{pct_var:.2f}%</td>'
            f'<td style="text-align:center"><span class="estado-box {estado_cls}">{estado_txt}</span></td>'
            f'</tr>'
        )

    cards_html = (
        generate_card("COSTO DIRECTO", pct_var_cd, var_cd)
        + generate_card("GASTOS GENERALES", pct_var_gg, var_gg)
        + generate_card("VARIACI&Oacute;N TOTAL", pct_var_total, var_total)
    )

    comp_rows = (
        comp_row("Costo Directo", rv["costoDirecto"], rc["totalCD"], var_cd, pct_var_cd)
        + comp_row("Gastos Generales", rv["gastosGenerales"], rc["totalGG"], var_gg, pct_var_gg)
        + comp_row("TOTAL", total_valorizado, total_ejecutado, var_total, pct_var_total, is_total=True)
    )

    return f"""
    <div class="page">
        <div class="header">
            <div class="header-titles">
                <h1>COS-PR02-FR02 REPORTE DE VALORIZACI&Oacute;N SEMANAL</h1>
                <h2>An&aacute;lisis Comparativo: Valorizaci&oacute;n vs Gastos Ejecutados</h2>
            </div>
            <div class="header-obra">
                <div><span class="header-obra-label">OBRA:</span> <span class="header-obra-value">{data["shortName"]}</span></div>
                <div class="header-fecha">{format_date(data["date"])}</div>
            </div>
        </div>

        <div class="section-title"><span class="numero">1</span>CORTE DE VALORIZACI&Oacute;N</div>
        <table>
            <thead><tr><th>Concepto</th><th class="num">Monto (S/)</th><th class="num">Porcentaje</th></tr></thead>
            <tbody>
                <tr><td>Costo Directo</td><td class="num">{fmt(rv["costoDirecto"])}</td><td class="num">100.00%</td></tr>
                <tr><td>Gastos Generales</td><td class="num">{fmt(rv["gastosGenerales"])}</td><td class="num">{rv["ggPercent"]:.2f}%</td></tr>
                <tr><td>Utilidad</td><td class="num">{fmt(rv["utilidad"])}</td><td class="num">{rv["utilPercent"]:.2f}%</td></tr>
                <tr><td>Sub Total</td><td class="num">{fmt(sub_total)}</td><td class="num">&mdash;</td></tr>
                <tr><td>IGV</td><td class="num">{fmt(igv)}</td><td class="num">18.00%</td></tr>
                <tr class="total-row"><td><strong>Total Valorizaci&oacute;n</strong></td><td class="num"><strong>{fmt(total_con_igv)}</strong></td><td class="num">&mdash;</td></tr>
            </tbody>
        </table>

        <div class="two-columns">
            <div>
                <div class="section-title"><span class="numero">2</span>GASTOS EJECUTADOS - COSTO DIRECTO</div>
                <table>
                    <thead><tr><th>Concepto</th><th class="num">Monto (S/)</th><th class="num">%</th></tr></thead>
                    <tbody>
                        {cd_rows_html}
                        <tr class="total-row"><td><strong>TOTAL CD EJECUTADO</strong></td><td class="num"><strong>{fmt(rc["totalCD"])}</strong></td><td class="num"><strong>100.00%</strong></td></tr>
                    </tbody>
                </table>
            </div>
            <div>
                <div class="section-title"><span class="numero">3</span>GASTOS GENERALES EJECUTADOS</div>
                <table>
                    <thead><tr><th>Concepto</th><th class="num">Monto (S/)</th><th class="num">%</th></tr></thead>
                    <tbody>
                        {gg_rows_html}
                        <tr class="total-row"><td><strong>TOTAL GG EJECUTADOS</strong></td><td class="num"><strong>{fmt(rc["totalGG"])}</strong></td><td class="num"><strong>100.00%</strong></td></tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="section-title"><span class="numero">4</span>AN&Aacute;LISIS COMPARATIVO - VALORIZACI&Oacute;N VS GASTOS EJECUTADOS</div>
        <div class="cards-container">
            {cards_html}
        </div>
        <table class="tabla-comparativa">
            <thead><tr>
                <th>Concepto</th><th class="num">Valorizaci&oacute;n (S/)</th><th class="num">Ejecutado (S/)</th>
                <th class="num">Variaci&oacute;n (S/)</th><th class="num">Var. (%)</th><th style="text-align:center">Estado</th>
            </tr></thead>
            <tbody>
                {comp_rows}
            </tbody>
        </table>
    </div>"""


# ============================================================
# SVG CHART GENERATION
# ============================================================

def generate_svg_chart(zoom_data, mes_actual_idx, has_plan):
    """
    Genera el grafico SVG de Curva S.
    zoom_data: lista de dicts con keys: index, prog, ejec, plan, isMesActual, isProyeccion
    mes_actual_idx: indice del mes actual en los datos originales
    has_plan: bool indicando si hay datos proyectados
    """
    W, H = 730, 420
    margin_left, margin_right, margin_top, margin_bottom = 50, 60, 25, 55
    chart_w = W - margin_left - margin_right
    chart_h = H - margin_top - margin_bottom

    n = len(zoom_data)
    if n < 2:
        return "<p>Datos insuficientes para grafico</p>"

    x_step = chart_w / (n - 1)

    # Y scale: tight to data with small padding
    max_pct = 0
    for d in zoom_data:
        max_pct = max(max_pct, d["prog"]["acumPct"] * 100)
        if d["ejec"]:
            max_pct = max(max_pct, (d["ejec"].get("acumPct") or 0) * 100)
        if d["plan"]:
            max_pct = max(max_pct, (d["plan"].get("acumPct") or 0) * 100)

    y_max = math.ceil(max_pct / 5) * 5 + 2  # +2% headroom
    if y_max <= 2:
        return "<p>Sin datos de porcentaje</p>"

    def x_pos(i):
        return margin_left + i * x_step

    def y_pos(pct):
        return margin_top + chart_h - (pct / y_max) * chart_h

    svg = f'<svg width="100%" viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" style="font-family:\'Segoe UI\',sans-serif">'

    # Defs (gradients)
    svg += """<defs>
        <linearGradient id="gradProg" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="#2c5aa0" stop-opacity="0.18"/>
            <stop offset="100%" stop-color="#2c5aa0" stop-opacity="0.02"/>
        </linearGradient>
        <linearGradient id="gradEjec" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="#28a745" stop-opacity="0.12"/>
            <stop offset="100%" stop-color="#28a745" stop-opacity="0.01"/>
        </linearGradient>
    </defs>"""

    # Find projection start index and current month local index
    proj_start_idx = -1
    mes_actual_local_idx = -1
    for i, d in enumerate(zoom_data):
        if d["isProyeccion"] and proj_start_idx < 0:
            proj_start_idx = i
        if d["isMesActual"]:
            mes_actual_local_idx = i

    # Background zones - current month highlight
    if mes_actual_local_idx >= 0:
        x1 = x_pos(mes_actual_local_idx) - x_step / 2
        x2 = x_pos(mes_actual_local_idx) + x_step / 2
        rect_x = max(margin_left, x1)
        rect_w = min(x2, margin_left + chart_w) - rect_x
        svg += f'<rect x="{rect_x}" y="{margin_top}" width="{rect_w}" height="{chart_h}" fill="#fff3cd" opacity="0.45"/>'

    # Projection zone
    if proj_start_idx >= 0:
        proj_x = x_pos(proj_start_idx) - x_step / 2
        svg += f'<rect x="{proj_x}" y="{margin_top}" width="{margin_left + chart_w - proj_x}" height="{chart_h}" fill="#f5f7fb"/>'
        svg += f'<line x1="{proj_x}" y1="{margin_top}" x2="{proj_x}" y2="{margin_top + chart_h}" stroke="#bbb" stroke-dasharray="6,3" stroke-width="1"/>'
        svg += f'<text x="{(proj_x + margin_left + chart_w) / 2}" y="{margin_top + 14}" text-anchor="middle" font-size="10" fill="#aaa" font-weight="700" letter-spacing="1">PROYECCI&Oacute;N</text>'

    # Grid lines (every 5%)
    grid_step = 5
    pct = 0
    while pct <= y_max:
        y = y_pos(pct)
        is_major = pct % 10 == 0
        stroke = "#ddd" if is_major else "#eee"
        sw = 1 if is_major else 0.5
        fs = "10" if is_major else "9"
        fill = "#666" if is_major else "#aaa"
        fw = "600" if is_major else "400"
        svg += f'<line x1="{margin_left}" y1="{y}" x2="{margin_left + chart_w}" y2="{y}" stroke="{stroke}" stroke-width="{sw}"/>'
        svg += f'<text x="{margin_left - 8}" y="{y + 4}" text-anchor="end" font-size="{fs}" fill="{fill}" font-weight="{fw}">{pct}%</text>'
        pct += grid_step

    # Axes
    svg += f'<line x1="{margin_left}" y1="{margin_top}" x2="{margin_left}" y2="{margin_top + chart_h}" stroke="#ddd" stroke-width="1"/>'
    svg += f'<line x1="{margin_left}" y1="{margin_top + chart_h}" x2="{margin_left + chart_w}" y2="{margin_top + chart_h}" stroke="#bbb" stroke-width="1"/>'

    # X labels
    for i in range(n):
        x = x_pos(i)
        label = re.sub(r"INICIO \d+/\d+/\d+", "INICIO", zoom_data[i]["prog"]["mes"])
        parts = label.split(" ")
        # Tick
        svg += f'<line x1="{x}" y1="{margin_top + chart_h}" x2="{x}" y2="{margin_top + chart_h + 4}" stroke="#bbb"/>'
        is_current = i == mes_actual_local_idx
        label_color = "#856404" if is_current else ("#aaa" if zoom_data[i]["isProyeccion"] else "#555")
        label_weight = "700" if is_current else "400"
        svg += f'<text x="{x}" y="{margin_top + chart_h + 16}" text-anchor="middle" font-size="9.5" fill="{label_color}" font-weight="{label_weight}">{parts[0]}</text>'
        if len(parts) > 1:
            svg += f'<text x="{x}" y="{margin_top + chart_h + 27}" text-anchor="middle" font-size="8.5" fill="{label_color}" font-weight="{label_weight}">{parts[1]}</text>'

    # Build point arrays
    prog_points = []
    ejec_points = []
    plan_points = []

    for i in range(n):
        d = zoom_data[i]
        prog_points.append({
            "x": x_pos(i),
            "y": y_pos(d["prog"]["acumPct"] * 100),
            "pct": d["prog"]["acumPct"] * 100,
            "idx": i,
        })
        if i <= mes_actual_local_idx:
            ejec_pct = (d["ejec"]["acumPct"] if d["ejec"] and d["ejec"].get("acumPct") else 0) * 100
            ejec_points.append({
                "x": x_pos(i),
                "y": y_pos(ejec_pct),
                "pct": ejec_pct,
                "idx": i,
            })
        if has_plan and d["plan"]:
            plan_pct = (d["plan"]["acumPct"] if d["plan"].get("acumPct") else 0) * 100
            if plan_pct > 0 or i == 0:
                plan_points.append({
                    "x": x_pos(i),
                    "y": y_pos(plan_pct),
                    "pct": plan_pct,
                    "idx": i,
                })

    # Area fills
    if len(prog_points) > 1:
        prog_area_end = proj_start_idx if proj_start_idx >= 0 else len(prog_points)
        area_path = f"M{prog_points[0]['x']},{margin_top + chart_h}"
        for i in range(min(prog_area_end, len(prog_points))):
            area_path += f" L{prog_points[i]['x']},{prog_points[i]['y']}"
        last_idx = min(prog_area_end - 1, len(prog_points) - 1)
        area_path += f" L{prog_points[last_idx]['x']},{margin_top + chart_h} Z"
        svg += f'<path d="{area_path}" fill="url(#gradProg)"/>'

    if len(ejec_points) > 1:
        area_path = f"M{ejec_points[0]['x']},{margin_top + chart_h}"
        for p in ejec_points:
            area_path += f" L{p['x']},{p['y']}"
        area_path += f" L{ejec_points[-1]['x']},{margin_top + chart_h} Z"
        svg += f'<path d="{area_path}" fill="url(#gradEjec)"/>'

    # Lines
    # Contractual: solid until current, dashed in projection
    if len(prog_points) > 1:
        solid_end = proj_start_idx if proj_start_idx >= 0 else len(prog_points)
        if solid_end > 0:
            path = f"M{prog_points[0]['x']},{prog_points[0]['y']}"
            for i in range(1, min(solid_end, len(prog_points))):
                path += f" L{prog_points[i]['x']},{prog_points[i]['y']}"
            svg += f'<path d="{path}" fill="none" stroke="#2c5aa0" stroke-width="3" stroke-linejoin="round"/>'
        if solid_end > 0 and solid_end < len(prog_points):
            path = f"M{prog_points[solid_end - 1]['x']},{prog_points[solid_end - 1]['y']}"
            for i in range(solid_end, len(prog_points)):
                path += f" L{prog_points[i]['x']},{prog_points[i]['y']}"
            svg += f'<path d="{path}" fill="none" stroke="#2c5aa0" stroke-width="2.5" stroke-dasharray="8,5" stroke-linejoin="round"/>'

    # Valorizado line
    if len(ejec_points) > 1:
        path = f"M{ejec_points[0]['x']},{ejec_points[0]['y']}"
        for i in range(1, len(ejec_points)):
            path += f" L{ejec_points[i]['x']},{ejec_points[i]['y']}"
        svg += f'<path d="{path}" fill="none" stroke="#28a745" stroke-width="3" stroke-linejoin="round"/>'

    # Proyectado line
    if len(plan_points) > 1:
        path = f"M{plan_points[0]['x']},{plan_points[0]['y']}"
        for i in range(1, len(plan_points)):
            path += f" L{plan_points[i]['x']},{plan_points[i]['y']}"
        svg += f'<path d="{path}" fill="none" stroke="#e6a817" stroke-width="2.5" stroke-dasharray="8,5" stroke-linejoin="round"/>'

    # Data points
    for i in range(len(prog_points)):
        p = prog_points[i]
        is_current = zoom_data[i]["isMesActual"] if i < len(zoom_data) else False
        is_proj = zoom_data[i]["isProyeccion"] if i < len(zoom_data) else False
        r = 7 if is_current else 4.5
        opacity = 0.45 if is_proj else 1
        stroke_attr = ' stroke="white" stroke-width="3"' if is_current else ""
        svg += f'<circle cx="{p["x"]}" cy="{p["y"]}" r="{r}" fill="#2c5aa0" opacity="{opacity}"{stroke_attr}/>'

    for i in range(len(ejec_points)):
        p = ejec_points[i]
        is_current = p["idx"] == mes_actual_local_idx
        r = 7 if is_current else 4.5
        stroke_attr = ' stroke="white" stroke-width="3"' if is_current else ""
        svg += f'<circle cx="{p["x"]}" cy="{p["y"]}" r="{r}" fill="#28a745"{stroke_attr}/>'

    for p in plan_points:
        svg += f'<circle cx="{p["x"]}" cy="{p["y"]}" r="4.5" fill="#e6a817"/>'

    # ================================================================
    # SMART LABEL PLACEMENT (anti-collision)
    # ================================================================
    BADGE_H = 18
    BADGE_GAP = 4
    MIN_SEP = BADGE_H + BADGE_GAP  # 22px min between badge centers

    def spread_labels(items, min_sep):
        """Spread an array of items so none overlap (min separation = min_sep)."""
        if len(items) <= 1:
            return items
        # Sort by original Y (top to bottom = smallest Y first)
        items.sort(key=lambda a: a["y"])
        # Iteratively push overlapping items apart
        for _pass in range(10):
            moved = False
            for i in range(1, len(items)):
                gap = items[i]["y"] - items[i - 1]["y"]
                if gap < min_sep:
                    push = (min_sep - gap) / 2 + 0.5
                    items[i - 1]["y"] -= push
                    items[i]["y"] += push
                    moved = True
            if not moved:
                break
        return items

    # Deviation bracket at current month
    if mes_actual_local_idx >= 0:
        mx = x_pos(mes_actual_local_idx)
        prog_pct_val = zoom_data[mes_actual_local_idx]["prog"]["acumPct"] * 100
        ejec_pct_val = zoom_data[mes_actual_local_idx]["ejec"]["acumPct"] * 100
        prog_y = y_pos(prog_pct_val)
        ejec_y = y_pos(ejec_pct_val)
        deviation = prog_pct_val - ejec_pct_val

        # Bracket
        if abs(deviation) > 0.01:
            y1 = min(prog_y, ejec_y)
            y2 = max(prog_y, ejec_y)
            bx = mx - 16
            svg += f'<line x1="{bx}" y1="{y1}" x2="{bx}" y2="{y2}" stroke="#dc3545" stroke-width="2.5"/>'
            svg += f'<line x1="{bx - 5}" y1="{y1}" x2="{bx + 5}" y2="{y1}" stroke="#dc3545" stroke-width="2.5"/>'
            svg += f'<line x1="{bx - 5}" y1="{y2}" x2="{bx + 5}" y2="{y2}" stroke="#dc3545" stroke-width="2.5"/>'
            badge_y = (y1 + y2) / 2
            dev_text = ("-" if deviation > 0 else "+") + f"{abs(deviation):.2f}%"
            dev_badge_w = len(dev_text) * 6.5 + 10
            svg += f'<rect x="{bx - dev_badge_w + 2}" y="{badge_y - 10}" width="{dev_badge_w}" height="20" rx="4" fill="#f8d7da" stroke="#dc3545" stroke-width="1"/>'
            svg += f'<text x="{bx - dev_badge_w / 2 + 2}" y="{badge_y + 4}" text-anchor="middle" font-size="10" fill="#dc3545" font-weight="700">{dev_text}</text>'

        # Value badges at current month (with anti-collision)
        badge_x = mx + 12
        badge_w = 52
        current_badges = [
            {"y": prog_y, "color": "#2c5aa0", "text": f"{prog_pct_val:.2f}%"},
            {"y": ejec_y, "color": "#28a745", "text": f"{ejec_pct_val:.2f}%"},
        ]
        if has_plan and zoom_data[mes_actual_local_idx]["plan"]:
            plan_pct_val = zoom_data[mes_actual_local_idx]["plan"]["acumPct"] * 100
            current_badges.append({"y": y_pos(plan_pct_val), "color": "#e6a817", "text": f"{plan_pct_val:.2f}%"})

        current_badges = spread_labels(current_badges, MIN_SEP)

        for b in current_badges:
            svg += f'<rect x="{badge_x}" y="{b["y"] - 9}" width="{badge_w}" height="{BADGE_H}" rx="4" fill="{b["color"]}"/>'
            svg += f'<text x="{badge_x + badge_w / 2}" y="{b["y"] + 4}" text-anchor="middle" font-size="10" fill="white" font-weight="700">{b["text"]}</text>'

        # "HOY" badge
        svg += f'<rect x="{mx - 16}" y="{margin_top + chart_h + 34}" width="32" height="18" rx="5" fill="#fff3cd" stroke="#d4a017" stroke-width="1.2"/>'
        svg += f'<text x="{mx}" y="{margin_top + chart_h + 46}" text-anchor="middle" font-size="10" fill="#856404" font-weight="700">HOY</text>'

    # Labels in projection months
    for i in range(n):
        if zoom_data[i]["isProyeccion"] and i < len(prog_points):
            p = prog_points[i]
            svg += f'<rect x="{p["x"] - 22}" y="{p["y"] - 20}" width="44" height="16" rx="3" fill="#2c5aa0" opacity="0.25"/>'
            svg += f'<text x="{p["x"]}" y="{p["y"] - 9}" text-anchor="middle" font-size="9.5" fill="#2c5aa0" opacity="0.8" font-weight="600">{p["pct"]:.1f}%</text>'

    svg += "</svg>"
    return svg


# ============================================================
# PAGE 2: CURVA S
# ============================================================

def generate_page2(data):
    """Genera la pagina 2 del reporte: Curva S y tabla de avance."""
    curva = data["curva"]
    mes_actual_idx = curva["mesActualIndex"]
    has_plan = curva["proyectado"] is not None and len(curva["proyectado"]) > 0

    # Determine zoom range: INICIO to mesActual + 2
    zoom_end = min(mes_actual_idx + 2, len(curva["contractual"]) - 1)
    zoom_data = []
    for i in range(0, zoom_end + 1):
        zoom_data.append({
            "index": i,
            "prog": curva["contractual"][i],
            "ejec": curva["valorizado"][i],
            "plan": curva["proyectado"][i] if has_plan else None,
            "isMesActual": i == mes_actual_idx,
            "isProyeccion": i > mes_actual_idx,
        })

    def label_mes(mes):
        return re.sub(r"INICIO \d+/\d+/\d+", "INICIO", mes)

    # Current month data for cards
    prog_actual = curva["contractual"][mes_actual_idx]
    ejec_actual = curva["valorizado"][mes_actual_idx]
    plan_actual = curva["proyectado"][mes_actual_idx] if has_plan else None

    # Last zoom month name
    last_zoom_mes = label_mes(curva["contractual"][zoom_end]["mes"]) if zoom_end < len(curva["contractual"]) else ""

    # SVG Chart
    svg_chart = generate_svg_chart(zoom_data, mes_actual_idx, has_plan)

    # Cards HTML
    plan_card_pct = f'{plan_actual["acumPct"] * 100:.2f}%' if has_plan and plan_actual else "N/D"
    plan_card_amt = f'S/ {fmt(plan_actual["acumulado"])}' if has_plan and plan_actual else "Sin datos"
    plan_card_class = "plan" if has_plan else "prog"
    plan_card_style = "" if has_plan else ' style="opacity:0.5"'
    plan_label_color = "#856404" if has_plan else "#666"
    plan_value_color = "#856404" if has_plan else "#999"
    plan_label_text = "Proyectado Acum." if has_plan else "Proyectado"

    cards_html = f"""
    <div class="summary-cards-curva">
        <div class="summary-card-curva prog">
            <div class="card-label" style="color:#2c5aa0">Contractual Acum.</div>
            <div class="card-pct" style="color:#2c5aa0">{prog_actual["acumPct"] * 100:.2f}%</div>
            <div class="card-amt" style="color:#2c5aa0">S/ {fmt(prog_actual["acumulado"])}</div>
        </div>
        <div class="summary-card-curva ejec">
            <div class="card-label" style="color:#155724">Valorizado Acum.</div>
            <div class="card-pct" style="color:#155724">{ejec_actual["acumPct"] * 100:.2f}%</div>
            <div class="card-amt" style="color:#155724">S/ {fmt(ejec_actual["acumulado"])}</div>
        </div>
        <div class="summary-card-curva {plan_card_class}"{plan_card_style}>
            <div class="card-label" style="color:{plan_label_color}">{plan_label_text}</div>
            <div class="card-pct" style="color:{plan_value_color}">{plan_card_pct}</div>
            <div class="card-amt" style="color:{plan_value_color}">{plan_card_amt}</div>
        </div>
    </div>"""

    # Table rows
    table_rows = ""
    for d in zoom_data:
        row_class = "mes-actual" if d["isMesActual"] else ("mes-proyeccion" if d["isProyeccion"] else "")
        marker = '<span style="color:#d4a017">&#9679;</span> ' if d["isMesActual"] else ""

        prog_parcial = fmt(d["prog"]["parcial"])
        prog_acum_pct = f'{d["prog"]["acumPct"] * 100:.2f}%'

        if d["isProyeccion"]:
            ejec_parcial = '<span class="dash">&mdash;</span>'
            ejec_acum_pct = '<span class="dash">&mdash;</span>'
            plan_parcial = '<span class="dash">&mdash;</span>'
            plan_acum_pct = '<span class="dash">&mdash;</span>'
        else:
            ejec_parcial = fmt(d["ejec"]["parcial"])
            ejec_acum_pct = f'{d["ejec"]["acumPct"] * 100:.2f}%'
            if has_plan and d["plan"]:
                plan_parcial = fmt(d["plan"]["parcial"])
                plan_acum_pct = f'{d["plan"]["acumPct"] * 100:.2f}%'
            else:
                plan_parcial = '<span class="dash">&mdash;</span>'
                plan_acum_pct = '<span class="dash">&mdash;</span>'

        # Style acum pct in current month
        prog_acum_style = 'color:#2c5aa0;font-weight:700' if d["isMesActual"] else ""
        ejec_acum_style = 'color:#28a745;font-weight:700' if d["isMesActual"] else ""
        plan_acum_style = 'color:#e6a817;font-weight:700' if d["isMesActual"] else ""

        table_rows += f"""
        <tr class="{row_class}">
            <td>{marker}{label_mes(d["prog"]["mes"])}</td>
            <td class="num">{prog_parcial}</td>
            <td class="num" style="{prog_acum_style}">{prog_acum_pct}</td>
            <td class="num">{ejec_parcial}</td>
            <td class="num" style="{ejec_acum_style}">{ejec_acum_pct}</td>
            <td class="num">{plan_parcial}</td>
            <td class="num" style="{plan_acum_style}">{plan_acum_pct}</td>
        </tr>"""

    # Legend
    legend_items = [
        {"color": "#2c5aa0", "label": "Contractual", "type": "line"},
        {"color": "#28a745", "label": "Valorizado", "type": "line"},
    ]
    if has_plan:
        legend_items.append({"color": "#e6a817", "label": "Proyectado", "type": "line-dashed"})
    legend_items.append({"color": "#fff3cd", "label": "Mes Actual", "type": "square", "border": "#d4a017"})
    legend_items.append({"color": "#f0f4fa", "label": "Proyecci&oacute;n", "type": "square", "border": "#ccc"})

    legend_html = ""
    for it in legend_items:
        if it["type"] == "line":
            legend_html += f'<div class="legend-item"><div class="legend-swatch" style="background:{it["color"]}"></div>{it["label"]}</div>'
        elif it["type"] == "line-dashed":
            legend_html += f'<div class="legend-item"><div class="legend-swatch" style="background:{it["color"]};background:repeating-linear-gradient(90deg,{it["color"]} 0,{it["color"]} 4px,transparent 4px,transparent 7px)"></div>{it["label"]}</div>'
        else:
            border = it.get("border", it["color"])
            legend_html += f'<div class="legend-item"><div class="legend-square" style="background:{it["color"]};border:1px solid {border}"></div>{it["label"]}</div>'

    header_subtitle = f'Contractual vs Valorizado{"  vs Proyectado" if has_plan else ""} (CD + GG + Utilidad) &mdash; Zoom: Inicio &rarr; {last_zoom_mes}'

    return f"""
    <div class="page">
        <div class="header">
            <div class="header-titles">
                <h1>CURVA S - AVANCE ACUMULADO DEL PROYECTO</h1>
                <h2>{header_subtitle}</h2>
            </div>
            <div class="header-obra">
                <div><span class="header-obra-label">OBRA:</span> <span class="header-obra-value">{data["shortName"]}</span></div>
                <div class="header-fecha">{format_date(data["date"])}</div>
            </div>
        </div>

        {cards_html}

        <div class="section-title" style="margin-top:8px">
            <span class="numero">S</span>
            CURVA S - AVANCE ACUMULADO (%) &mdash; ZOOM HASTA {last_zoom_mes.upper()}
        </div>
        <div class="chart-container">
            {svg_chart}
        </div>

        <div class="legend-container">{legend_html}</div>

        <div class="nota-mes-completo">
            <span class="nota-label">Nota:</span> Los montos y porcentajes <span class="nota-bold">Contractuales</span> corresponden a la valorizaci&oacute;n del <span class="nota-mes">mes completo</span>, no al corte semanal.
        </div>

        <div class="section-title" style="margin-top:8px">
            <span class="numero">T</span>
            DETALLE DE AVANCE MENSUAL (CD + GG + UTILIDAD)
        </div>
        <table class="table-curva">
            <thead>
                <tr>
                    <th rowspan="2" style="border-bottom:2px solid #333">MES</th>
                    <th colspan="2" style="background:#e8edf5;color:#2c5aa0;text-align:center;border-bottom:2px solid #2c5aa0">CONTRACTUAL</th>
                    <th colspan="2" style="background:#d4edda;color:#155724;text-align:center;border-bottom:2px solid #28a745">VALORIZADO</th>
                    <th colspan="2" style="background:#fff3cd;color:#856404;text-align:center;border-bottom:2px solid #e6a817">PROYECTADO</th>
                </tr>
                <tr>
                    <th class="num" style="background:#e8edf5;color:#2c5aa0;border-bottom:1px solid #2c5aa0">Parcial (S/)</th>
                    <th class="num" style="background:#e8edf5;color:#2c5aa0;border-bottom:1px solid #2c5aa0">Acum.(%)</th>
                    <th class="num" style="background:#d4edda;color:#155724;border-bottom:1px solid #28a745">Parcial (S/)</th>
                    <th class="num" style="background:#d4edda;color:#155724;border-bottom:1px solid #28a745">Acum.(%)</th>
                    <th class="num" style="background:#fff3cd;color:#856404;border-bottom:1px solid #e6a817">Parcial (S/)</th>
                    <th class="num" style="background:#fff3cd;color:#856404;border-bottom:1px solid #e6a817">Acum.(%)</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>

        <div class="nota-mes-completo nota-tabla">
            <span class="nota-label">Nota:</span> Los montos y porcentajes <span class="nota-bold">Contractuales</span> corresponden a la valorizaci&oacute;n del <span class="nota-mes">mes completo</span>, no al corte semanal.
        </div>
    </div>"""


# ============================================================
# GENERATE FULL REPORT (standalone HTML)
# ============================================================

def generate_report(data):
    """Genera el reporte. Page 1 siempre, Page 2 (Curva S) solo si hay datos."""
    html = generate_page1(data)
    if data.get("curva"):
        html += generate_page2(data)
    else:
        print("  [REPORTE] Sin datos de CURVA - reporte de 1 pagina.")
    return html


# ============================================================
# CSS STYLES (complete report styles from the original app)
# ============================================================

REPORT_CSS = """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f0f2f5; margin: 0;
        }

        .page {
            width: 210mm; min-height: 297mm; padding: 10mm 12mm 8mm 12mm;
            margin: 10px auto; background: white;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 12px; color: #333;
            page-break-after: always; overflow: hidden;
        }
        .page:last-child { page-break-after: auto; }

        @media print {
            body { background: white; }
            .page { box-shadow: none; margin: 0; }
            @page { size: A4 portrait; margin: 5mm; }
        }

        /* Header */
        .header {
            display: flex; justify-content: space-between; align-items: flex-start;
            margin-bottom: 8px; padding-bottom: 8px; border-bottom: 2px solid #2c5aa0;
        }
        .header-titles h1 { font-size: 16px; color: #1e4077; margin-bottom: 2px; }
        .header-titles h2 { font-size: 12px; color: #555; font-weight: 400; }
        .header-obra { text-align: right; flex-shrink: 0; }
        .header-obra-label { font-size: 12px; color: #1e4077; font-weight: 400; }
        .header-obra-value { font-size: 18px; font-weight: 700; color: #1e4077; }
        .header-fecha { font-size: 11px; color: #1e4077; margin-top: 3px; font-weight: 700; }

        /* Section titles */
        .section-title {
            background: linear-gradient(90deg, #2c5aa0 0%, #3d6db5 100%);
            color: white; padding: 6px 12px; font-size: 12px; font-weight: 600;
            display: flex; align-items: center; border-radius: 4px 4px 0 0;
            margin-top: 10px;
        }
        .section-title .numero {
            background: white; color: #2c5aa0;
            width: 22px; height: 22px; border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            margin-right: 10px; font-size: 12px; font-weight: 700;
            flex-shrink: 0;
        }

        /* Tables */
        table {
            width: 100%; border-collapse: collapse; background: white;
            border: 1px solid #ddd; border-top: none;
        }
        th {
            background: #f8f9fa; color: #1e4077; padding: 7px 10px;
            font-weight: 700; font-size: 12px; border-bottom: 2px solid #2c5aa0;
            text-align: left;
        }
        td {
            padding: 7px 10px; border-bottom: 1px solid #eee; font-size: 12px;
        }
        td.num, th.num {
            text-align: right; font-family: 'Consolas', 'Courier New', monospace;
        }
        .total-row { background: #e8f4fd !important; }
        .total-row td {
            font-weight: 700; color: #1e4077;
            border-top: 2px solid #2c5aa0; border-bottom: 2px solid #2c5aa0;
        }

        /* Two-column layout */
        .two-columns {
            display: grid; grid-template-columns: 1fr 1fr;
            gap: 12px; margin-bottom: 8px;
        }
        .two-columns .section-title { margin-top: 0; }

        /* Cards */
        .cards-container {
            display: grid; grid-template-columns: repeat(3, 1fr);
            gap: 12px; margin: 10px 0;
        }
        .card {
            background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%);
            border: 1px solid #e0e0e0; border-radius: 8px;
            padding: 12px 10px; text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .card-title { font-size: 11px; color: #666; font-weight: 600; margin-bottom: 6px; text-transform: uppercase; }
        .card-value { font-size: 26px; font-weight: 700; }
        .card-value.positivo { color: #28a745; }
        .card-value.negativo { color: #dc3545; }
        .card-monto {
            font-size: 13px; font-weight: 600; margin-top: 6px;
            padding: 4px 10px; border-radius: 12px; display: inline-block;
        }
        .card-monto.ganancia { background: #d4edda; color: #155724; }
        .card-monto.perdida { background: #f8d7da; color: #721c24; }

        /* Estado boxes */
        .estado-box {
            display: inline-block; padding: 4px 12px; border-radius: 10px;
            font-weight: 700; font-size: 10px;
        }
        .estado-ganancia { background: #d4edda; color: #155724; }
        .estado-perdida { background: #f8d7da; color: #721c24; }
        .valor-positivo { color: #28a745; font-weight: 700; }
        .valor-negativo { color: #dc3545; font-weight: 700; }

        /* Tabla comparativa */
        .tabla-comparativa th, .tabla-comparativa td { text-align: right; font-size: 12px; }
        .tabla-comparativa th:first-child, .tabla-comparativa td:first-child { text-align: left; }
        .tabla-comparativa th:last-child, .tabla-comparativa td:last-child { text-align: center; }

        /* PAGE 2: CURVA S */
        .summary-cards-curva {
            display: grid; grid-template-columns: repeat(3, 1fr);
            gap: 10px; margin: 8px 0;
        }
        .summary-card-curva {
            border-radius: 8px; padding: 10px 8px; text-align: center;
        }
        .summary-card-curva.prog { background: linear-gradient(135deg, #e8edf5, #d5dff0); border: 1px solid #b8c9e2; }
        .summary-card-curva.ejec { background: linear-gradient(135deg, #d4edda, #c3e6cb); border: 1px solid #a3d5b1; }
        .summary-card-curva.plan { background: linear-gradient(135deg, #fff3cd, #ffeaa7); border: 1px solid #e6d590; }
        .summary-card-curva .card-label { font-size: 10px; font-weight: 600; text-transform: uppercase; margin-bottom: 4px; }
        .summary-card-curva .card-pct { font-size: 22px; font-weight: 700; }
        .summary-card-curva .card-amt { font-size: 11px; font-weight: 600; margin-top: 2px; }

        /* Chart container */
        .chart-container {
            background: white; border: 1px solid #ddd;
            border-radius: 0 0 4px 4px; border-top: none;
            padding: 15px 10px 8px 10px;
        }

        /* Curva table */
        .table-curva th { font-size: 11px; padding: 7px 10px; }
        .table-curva td { font-size: 12px; padding: 7px 10px; }
        .table-curva td.num { font-size: 11px; }
        .mes-actual { background: #fff3cd !important; font-weight: 700; }
        .mes-proyeccion { background: #f0f4fa !important; font-style: italic; color: #666; }
        .mes-proyeccion .dash { color: #bbb; }

        /* Legend */
        .legend-container {
            display: flex; justify-content: center; gap: 24px; margin: 10px 0 0 0;
        }
        .legend-item {
            display: flex; align-items: center; gap: 6px;
            font-size: 11px; font-weight: 600;
        }
        .legend-swatch {
            width: 18px; height: 4px; border-radius: 2px;
        }
        .legend-square {
            width: 12px; height: 12px; border-radius: 2px;
        }

        /* Notes */
        .nota-mes-completo {
            text-align: center; font-size: 10px; font-style: italic;
            background: #f8f9fa; border: 1px solid #eee; border-radius: 4px;
            padding: 5px 12px; margin-top: 6px;
        }
        .nota-mes-completo.nota-tabla {
            font-size: 11.5px; padding: 7px 14px; border-color: #ddd;
        }
        .nota-mes-completo .nota-label { color: #2c5aa0; font-weight: 700; }
        .nota-mes-completo .nota-bold { font-weight: 700; color: #2c5aa0; }
        .nota-mes-completo .nota-mes { font-weight: 700; }
"""


def _wrap_standalone_html(report_content):
    """Envuelve el contenido del reporte en un HTML standalone con CSS completo."""
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>COS-PR02-FR02 Reporte Valorizaci&oacute;n Semanal</title>
    <style>
{REPORT_CSS}
    </style>
</head>
<body>
{report_content}
</body>
</html>"""


# ============================================================
# PUBLIC API
# ============================================================

def generar(data):
    """
    Genera el reporte HTML completo de valorizacion semanal.

    Args:
        data: dict con las claves resCosto, rval, curva, projectName,
              shortName, date, author (producido por agente_excel.procesar())

    Returns:
        Tupla (html_string, filename) donde:
            html_string: HTML standalone completo con CSS embebido
            filename: nombre sugerido para el archivo, ej:
                      COS-PR02-FR02_VAL_SEMANAL_BEETHOVEN_22-Feb-2026.html
    """
    report_body = generate_report(data)
    html_string = _wrap_standalone_html(report_body)

    # Build filename
    obra = re.sub(r"\s+", "_", (data.get("shortName") or "REPORTE").strip())
    fecha = format_date_short(data.get("date"))
    fecha_part = f"_{fecha}" if fecha else ""
    filename = f"COS-PR02-FR02_VAL_SEMANAL_{obra}{fecha_part}.html"

    return html_string, filename
