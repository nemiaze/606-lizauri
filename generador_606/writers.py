"""
Generación de salidas: formato_606.csv, agrupacion_sin_comprobante.xlsx e incidencias.txt.
"""

from collections import defaultdict
from datetime import datetime
from decimal import Decimal
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import AppConfig
from .models import Incidence, ProcessedRecord, ProviderMaster


# Encabezados exactos del Formato 606 (25 columnas)
HEADERS_606 = [
    "RNC o Cédula",
    "Tipo Id",
    "Tipo Bienes y Servicios Comprados",
    "NCF",
    "NCF ó Documento Modificado",
    "Fecha Comprobante",
    "",  # Columna 7: DD (encabezado vacío según plantilla DGII)
    "Fecha Pago",
    "",  # Columna 9: DD (encabezado vacío según plantilla DGII)
    "Monto Facturado en Servicios",
    "Monto Facturado en Bienes",
    "Total Monto Facturado",
    "ITBIS Facturado",
    "ITBIS Retenido",
    "ITBIS sujeto a Proporcionalidad (Art. 349)",
    "ITBIS llevado al Costo",
    "ITBIS por Adelantar",
    "ITBIS percibido en compras",
    "Tipo de Retención en ISR",
    "Monto Retención Renta",
    "ISR Percibido en compras",
    "Impuesto Selectivo al Consumo",
    "Otros Impuesto/Tasas",
    "Monto Propina Legal",
    "Forma de Pago",
]


def write_606_csv(
    records: List[ProcessedRecord],
    providers: Dict[str, ProviderMaster],
    output_path: Path,
    config: AppConfig
) -> int:
    """
    Escribe el archivo formato_606.csv con codificación UTF-8 con BOM (utf-8-sig).
    Retorna la cantidad de registros escritos.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    count = 0

    valid_records = [r for r in records if r.is_606]
    valid_records.sort(key=lambda r: (r.tx.proveedor_norm, r.ncf_valido or ""))

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        # Escribir encabezado
        f.write(",".join(HEADERS_606) + "\r\n")

        for rec in valid_records:
            count += 1
            tx = rec.tx
            p_master = providers.get(tx.proveedor_norm)

            # 1. RNC o Cédula
            rnc_str = rec.rnc_cedula or ""

            # 2. Tipo Id (1 = RNC, 2 = Cédula)
            tipo_id_str = rec.tipo_id or ""

            # 3. Tipo Bienes y Servicios (formato CSV requiere código de 2 dígitos, ej. '09')
            tb_code = (p_master.tipo_bienes_servicios if (p_master and p_master.tipo_bienes_servicios)
                       else config.tipo_bienes_servicios_606)
            tb_code_csv = tb_code[:2] if tb_code else "09"

            # 4. NCF
            ncf_str = rec.ncf_valido or ""

            # 5. NCF modificado
            ncf_mod = ""

            # 6 y 7. Fecha Comprobante (AAAAMM y DD)
            if tx.fecha_parsed:
                a, m, d = tx.fecha_parsed
                f_aaaamm = f"{a:04d}{m:02d}"
                f_dd = f"{d:02d}"
            else:
                f_aaaamm = ""
                f_dd = ""

            # 8 y 9. Fecha Pago (en CSV vacía por especificación DGII)
            f_pago_aaaamm = ""
            f_pago_dd = ""

            # 10 y 11. Monto Servicios vs Bienes
            subtotal = tx.importe_sujeto
            es_servicio = (config.bienes_vs_servicios_606.lower() == "servicios" or rec.tipo_gasto in ("Servicio", "Servicio Profesional", "Servicio No Profesional"))
            if es_servicio:
                monto_servicios = f"{subtotal:.2f}"
                monto_bienes = "0.00"
            else:
                monto_servicios = "0.00"
                monto_bienes = f"{subtotal:.2f}"

            # 12. Total Monto Facturado
            total_facturado = f"{subtotal:.2f}"

            # 13. ITBIS Facturado: tomar el impuesto real de la factura
            if rec.itbis_exento:
                itbis_facturado = "0.00"
            else:
                itbis_facturado = f"{tx.impuesto:.2f}"

            # 14. ITBIS Retenido
            itbis_retenido = ""

            # 15. ITBIS Proporcionalidad
            itbis_prop = "0.00"

            # 16. ITBIS llevado al Costo
            itbis_costo = "0.00"

            # 17. ITBIS por Adelantar = ITBIS Facturado
            itbis_adelantar = itbis_facturado

            # 18 a 24: vacías
            itbis_perc = ""
            tipo_ret_isr = ""
            monto_ret_renta = ""
            isr_perc = ""
            isc = ""
            otros_imp = ""
            propina = ""

            # 25. Forma de Pago (formato CSV requiere código de 2 dígitos, ej. '02')
            forma_pago_csv = config.forma_pago_pagado[:2] if config.forma_pago_pagado else "02"

            row_data = [
                rnc_str,
                tipo_id_str,
                tb_code_csv,
                ncf_str,
                ncf_mod,
                f_aaaamm,
                f_dd,
                f_pago_aaaamm,
                f_pago_dd,
                monto_servicios,
                monto_bienes,
                total_facturado,
                itbis_facturado,
                itbis_retenido,
                itbis_prop,
                itbis_costo,
                itbis_adelantar,
                itbis_perc,
                tipo_ret_isr,
                monto_ret_renta,
                isr_perc,
                isc,
                otros_imp,
                propina,
                forma_pago_csv,
            ]

            f.write(",".join(row_data) + "\r\n")

    return count


def write_606_excel(
    records: List[ProcessedRecord],
    providers: Dict[str, ProviderMaster],
    output_path: Path,
    config: AppConfig
) -> int:
    """
    Escribe el archivo formato_606.xlsx listo para usar en Excel sin necesidad
    de convertir texto a columnas. Conserva ceros a la izquierda en RNCs/cédulas
    y aplica formato numérico y bordes a los importes.
    Retorna la cantidad de registros escritos.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formato 606"

    font_family = "Arial"
    font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_regular = Font(name=font_family, size=10)
    font_bold = Font(name=font_family, size=10, bold=True)

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border_side = Side(style="thin", color="D9D9D9")
    double_bottom_side = Side(style="double", color="000000")
    top_thin_side = Side(style="thin", color="000000")

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(top=top_thin_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    num_fmt_curr = "#,##0.00;(#,##0.00);-"

    # Encabezado (Fila 1)
    ws.append(HEADERS_606)
    for col_idx in range(1, len(HEADERS_606) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    valid_records = [r for r in records if r.is_606]
    valid_records.sort(key=lambda r: (r.tx.proveedor_norm, r.ncf_valido or ""))

    count = 0
    curr_row = 2

    for rec in valid_records:
        count += 1
        tx = rec.tx
        p_master = providers.get(tx.proveedor_norm)

        # 1. RNC o Cédula
        rnc_str = rec.rnc_cedula or ""

        # 2. Tipo Id (1 = RNC, 2 = Cédula)
        tipo_id_str = rec.tipo_id or ""

        # 3. Tipo Bienes y Servicios
        tb_code = (p_master.tipo_bienes_servicios if (p_master and p_master.tipo_bienes_servicios)
                   else config.tipo_bienes_servicios_606)

        # 4. NCF
        ncf_str = rec.ncf_valido or ""

        # 5. NCF modificado
        ncf_mod = None

        # 6 y 7. Fecha Comprobante (AAAAMM y DD)
        if tx.fecha_parsed:
            a, m, d = tx.fecha_parsed
            f_aaaamm = f"{a:04d}{m:02d}"
            f_dd = f"{d:02d}"
        else:
            f_aaaamm = None
            f_dd = None

        # 8 y 9. Fecha Pago (igual a Fecha Comprobante)
        f_pago_aaaamm = f_aaaamm
        f_pago_dd = f_dd

        # 10 y 11. Monto Servicios vs Bienes
        subtotal = float(tx.importe_sujeto)
        es_servicio = (config.bienes_vs_servicios_606.lower() == "servicios" or rec.tipo_gasto in ("Servicio", "Servicio Profesional", "Servicio No Profesional"))
        if es_servicio:
            monto_servicios = subtotal
            monto_bienes = None
        else:
            monto_servicios = None
            monto_bienes = subtotal

        # 12. Total Monto Facturado
        total_facturado = subtotal

        # 13. ITBIS Facturado: tomar el impuesto real de la factura
        if rec.itbis_exento:
            itbis_facturado = 0.0
        else:
            itbis_facturado = float(tx.impuesto)

        # 14. ITBIS Retenido
        itbis_retenido = None

        # 15. ITBIS Proporcionalidad
        itbis_prop = None

        # 16. ITBIS llevado al Costo
        itbis_costo = None

        # 17. ITBIS por Adelantar = ITBIS Facturado
        itbis_adelantar = itbis_facturado

        # 18 a 24: vacías
        itbis_perc = None
        tipo_ret_isr = None
        monto_ret_renta = None
        isr_perc = None
        isc = None
        otros_imp = None
        propina = None

        # 25. Forma de Pago: forzar 02 (cheque/transferencia)
        forma_pago = config.forma_pago_pagado

        row_vals = [
            rnc_str,
            tipo_id_str,
            tb_code,
            ncf_str,
            ncf_mod,
            f_aaaamm,
            f_dd,
            f_pago_aaaamm,
            f_pago_dd,
            monto_servicios,
            monto_bienes,
            total_facturado,
            itbis_facturado,
            itbis_retenido,
            itbis_prop,
            itbis_costo,
            itbis_adelantar,
            itbis_perc,
            tipo_ret_isr,
            monto_ret_renta,
            isr_perc,
            isc,
            otros_imp,
            propina,
            forma_pago,
        ]

        ws.append(row_vals)

        # Formatear celdas de la fila
        for col_idx in range(1, 26):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_cell

            # Importes numéricos
            if col_idx in (10, 11, 12, 13, 15, 16, 17):
                cell.number_format = num_fmt_curr
                cell.alignment = align_right
            # Textos y códigos que conservan ceros
            elif col_idx in (1, 2, 3, 4, 5, 6, 7, 8, 9, 25):
                cell.number_format = "@"
                cell.alignment = align_center

        curr_row += 1

    ws.freeze_panes = "A2"

    try:
        ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = cell.value
            if val is not None:
                s_val = str(val)
                if s_val.startswith("="):
                    val_len = 12
                else:
                    val_len = len(s_val.split("\n")[0])
                if val_len > max_len:
                    max_len = val_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

    wb.save(output_path)
    return count


def write_excel_agrupacion(
    records: List[ProcessedRecord],
    output_path: Path,
    config: AppConfig
) -> None:
    """
    Genera el libro Excel agrupacion_sin_comprobante.xlsx con las 3 hojas:
    1. Agrupado por proveedor (con fórmulas dinámicas y tasa en B3)
    2. Gastos menores B13
    3. Detalle (fuente de datos para fórmulas)
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.calculation.fullCalcOnLoad = True

    # Estilos comunes (Fuente Arial en todo el libro)
    font_family = "Arial"
    font_title = Font(name=font_family, size=12, bold=True, color="1F4E78")
    font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=10, bold=True)
    font_regular = Font(name=font_family, size=10)
    font_rate = Font(name=font_family, size=11, bold=True, color="002060")

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_yellow_rate = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    fill_obs_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border_side = Side(style="thin", color="D9D9D9")
    double_bottom_side = Side(style="double", color="000000")
    top_thin_side = Side(style="thin", color="000000")

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_total = Border(top=top_thin_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)
    border_rate = Border(left=top_thin_side, right=top_thin_side, top=top_thin_side, bottom=top_thin_side)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    num_fmt_curr = "#,##0.00;(#,##0.00);-"
    num_fmt_int = "#,##0"

    # =========================================================================
    # Hoja 3: Detalle (Crear primero para tener la fuente de datos)
    # =========================================================================
    ws_detalle = wb.create_sheet(title="Detalle")
    headers_detalle = [
        "Fecha", "Proveedor", "RNC o Cédula", "Fuente del RNC", "Tipo",
        "Destino", "Nota", "Importe Total", "Monto sujeto a retención", "ITBIS",
        "Monto Retención de renta", "Importe Efectivo", "Entidad",
        "Número (no válido)", "Observaciones"
    ]
    ws_detalle.append(headers_detalle)

    # Estilizar encabezado Detalle
    for col_idx in range(1, len(headers_detalle) + 1):
        cell = ws_detalle.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    # Agregar filas de registros sin comprobante
    non_606_records = [r for r in records if not r.is_606]
    t_prof = float(config.tasa_retencion_profesional)
    t_no_prof = float(config.tasa_retencion_no_profesional)

    for r in non_606_records:
        tx = r.tx
        # Fecha formateada yyyy-mm-dd
        if tx.fecha_parsed:
            f_str = f"{tx.fecha_parsed[0]:04d}-{tx.fecha_parsed[1]:02d}-{tx.fecha_parsed[2]:02d}"
        else:
            f_str = tx.fecha_raw.split(" ")[0] if tx.fecha_raw else ""

        nota_limpia = tx.nota.replace("\r\n", " | ").replace("\n", " | ") if tx.nota else ""
        obs_str = "; ".join(r.observaciones) if r.observaciones else ""

        monto_pagado = float(tx.importe_sujeto)

        if r.tipo_gasto == "Servicio No Profesional":
            monto_sujeto = round(monto_pagado / (1.0 - t_no_prof), 2)
            ret_renta = round(monto_sujeto * t_no_prof, 2)
            itbis_calc = 0.0 if r.itbis_exento else round(monto_sujeto * 0.18, 2)
        elif r.tipo_gasto == "Servicio Profesional":
            monto_sujeto = round(monto_pagado / (1.0 - t_prof), 2)
            ret_renta = round(monto_sujeto * t_prof, 2)
            itbis_calc = 0.0 if r.itbis_exento else round(monto_sujeto * 0.18, 2)
        elif r.tipo_gasto == "Bien":
            monto_sujeto = monto_pagado
            ret_renta = 0.0
            itbis_calc = 0.0 if r.itbis_exento else round(monto_sujeto * 0.18, 2)
        else:
            monto_sujeto = monto_pagado
            ret_renta = 0.0
            itbis_calc = 0.0 if r.itbis_exento else (float(tx.impuesto) if float(tx.impuesto) > 0 else 0.0)

        importe_efectivo = round(monto_sujeto + itbis_calc, 2)

        row_vals = [
            f_str,
            tx.proveedor,
            r.rnc_cedula or "",
            r.fuente_rnc,
            r.tipo_gasto,
            r.destino,
            nota_limpia,
            monto_pagado,
            monto_sujeto,
            itbis_calc,
            ret_renta,
            importe_efectivo,
            tx.departamento,
            tx.numero,
            obs_str
        ]
        ws_detalle.append(row_vals)

    # Formatear celdas de Detalle
    for row_idx in range(2, ws_detalle.max_row + 1):
        for col_idx in range(1, len(headers_detalle) + 1):
            cell = ws_detalle.cell(row=row_idx, column=col_idx)
            cell.font = font_regular
            cell.border = border_cell
            if col_idx in (8, 9, 10, 11, 12):  # Importe Total, Monto sujeto, ITBIS, Ret. renta, Importe Efectivo
                cell.number_format = num_fmt_curr
                cell.alignment = align_right
            elif col_idx in (1, 3, 4, 5, 6, 14):  # Fecha, RNC, etc.
                cell.alignment = align_center
                if col_idx == 3:
                    cell.number_format = "@"

    ws_detalle.freeze_panes = "A2"

    # =========================================================================
    # Hoja 1: Agrupado por proveedor
    # =========================================================================
    ws_agrupado = wb.active
    ws_agrupado.title = "Agrupado por proveedor"

    # Encabezado superior
    ws_agrupado.cell(row=1, column=1, value="AGRUPACIÓN DE COMPRAS SIN COMPROBANTE (DGII)").font = font_title

    # Fila 3: Tasa Servicios Profesionales (15%)
    lbl_prof = ws_agrupado.cell(row=3, column=1, value="Tasa retención profesional (15%):")
    lbl_prof.font = font_bold
    lbl_prof.alignment = align_right

    cell_rate_prof = ws_agrupado.cell(row=3, column=2, value=float(config.tasa_retencion_profesional))
    cell_rate_prof.font = font_rate
    cell_rate_prof.fill = fill_yellow_rate
    cell_rate_prof.alignment = align_center
    cell_rate_prof.number_format = "0.0%"
    cell_rate_prof.border = border_rate
    cell_rate_prof.comment = Comment(
        "Tasa de retención ISR para Servicios Profesionales (honorarios, asesoría,\n"
        "consultoría, etc.). Por defecto: 15%. Editable para recalcular retenciones.",
        "Generador 606"
    )

    # Fila 4: Tasa Servicios No Profesionales (3%)
    lbl_no_prof = ws_agrupado.cell(row=4, column=1, value="Tasa retención no profesional (3%):")
    lbl_no_prof.font = font_bold
    lbl_no_prof.alignment = align_right

    cell_rate_no_prof = ws_agrupado.cell(row=4, column=2, value=float(config.tasa_retencion_no_profesional))
    cell_rate_no_prof.font = font_rate
    cell_rate_no_prof.fill = fill_yellow_rate
    cell_rate_no_prof.alignment = align_center
    cell_rate_no_prof.number_format = "0.0%"
    cell_rate_no_prof.border = border_rate
    cell_rate_no_prof.comment = Comment(
        "Tasa de retención ISR para Servicios No Profesionales (mano de obra,\n"
        "transporte, mantenimiento, etc.). Por defecto: 3%. Editable.",
        "Generador 606"
    )

    # Encabezados de la tabla (fila 6, encabezados de datos en fila 6)
    headers_agrupado = [
        "RNC o Cédula", "Tipo Id", "Proveedor", "Registros",
        "Serv. Profesional (15%)", "Serv. No Profesional (3%)", "Bienes", "Sin clasificar",
        "Importe Total (pagado)", "Monto sujeto a retención", "ITBIS", "Monto Retención de renta",
        "Tipo Retención ISR", "NCF (por asignar)", "Observaciones"
    ]

    header_row_agrupado = 6
    for col_idx, h_text in enumerate(headers_agrupado, start=1):
        c = ws_agrupado.cell(row=header_row_agrupado, column=col_idx, value=h_text)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center

    # Agrupar registros con destino "Agrupado"
    grouped_map: Dict[str, Dict[str, Any]] = {}
    for r in non_606_records:
        if r.destino == "Agrupado":
            norm_p = r.tx.proveedor_norm
            if norm_p not in grouped_map:
                grouped_map[norm_p] = {
                    "display_name": r.tx.proveedor,
                    "rnc": r.rnc_cedula or "",
                    "tipo_id": r.tipo_id or "",
                    "observaciones": list(r.observaciones),
                }
            else:
                for obs in r.observaciones:
                    if obs not in grouped_map[norm_p]["observaciones"]:
                        grouped_map[norm_p]["observaciones"].append(obs)

    # Insertar filas de proveedores agrupados (datos desde fila 7)
    curr_row = 7
    for norm_p, info in grouped_map.items():
        prov_name = info["display_name"]
        rnc_str = info["rnc"]
        tipo_id_str = info["tipo_id"]
        obs_text = "; ".join(info["observaciones"])

        # Col A: RNC
        cA = ws_agrupado.cell(row=curr_row, column=1, value=rnc_str)
        cA.number_format = "@"
        cA.alignment = align_center

        # Col B: Tipo Id
        cB = ws_agrupado.cell(row=curr_row, column=2, value=tipo_id_str)
        cB.alignment = align_center

        # Col C: Proveedor
        cC = ws_agrupado.cell(row=curr_row, column=3, value=prov_name)
        cC.alignment = align_left

        # Col D: Registros
        cD = ws_agrupado.cell(row=curr_row, column=4, value=f'=COUNTIFS(Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado")')
        cD.number_format = num_fmt_int
        cD.alignment = align_right

        # Col E: Servicios Profesionales (15%)
        cE = ws_agrupado.cell(row=curr_row, column=5, value=f'=SUMIFS(Detalle!$I:$I, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Servicio Profesional")')
        cE.number_format = num_fmt_curr
        cE.alignment = align_right

        # Col F: Servicios No Profesionales (3%)
        cF = ws_agrupado.cell(row=curr_row, column=6, value=f'=SUMIFS(Detalle!$I:$I, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Servicio No Profesional")')
        cF.number_format = num_fmt_curr
        cF.alignment = align_right

        # Col G: Bienes
        cG = ws_agrupado.cell(row=curr_row, column=7, value=f'=SUMIFS(Detalle!$I:$I, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Bien")')
        cG.number_format = num_fmt_curr
        cG.alignment = align_right

        # Col H: Sin clasificar
        cH = ws_agrupado.cell(row=curr_row, column=8, value=f'=SUMIFS(Detalle!$I:$I, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Sin clasificar")')
        cH.number_format = num_fmt_curr
        cH.alignment = align_right

        # Col I: Importe Total (pagado)
        cI = ws_agrupado.cell(row=curr_row, column=9, value=f'=SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado")')
        cI.number_format = num_fmt_curr
        cI.alignment = align_right

        # Col J: Monto sujeto a retención: =SUM(E:H)
        cJ = ws_agrupado.cell(row=curr_row, column=10, value=f'=SUM(E{curr_row}:H{curr_row})')
        cJ.number_format = num_fmt_curr
        cJ.alignment = align_right

        # Col K: ITBIS
        cK = ws_agrupado.cell(row=curr_row, column=11, value=f'=SUMIFS(Detalle!$J:$J, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado")')
        cK.number_format = num_fmt_curr
        cK.alignment = align_right

        # Col L: Retención renta = Prof*15% + NoProf*3%
        cL = ws_agrupado.cell(row=curr_row, column=12, value=f'=ROUND(E{curr_row}*$B$3 + F{curr_row}*$B$4, 2)')
        cL.number_format = num_fmt_curr
        cL.alignment = align_right

        # Col M: Tipo Retención ISR: "02" si hay algún servicio con retención
        cM = ws_agrupado.cell(row=curr_row, column=13, value=f'=IF(L{curr_row}>0, "{config.tipo_retencion_isr}", "")')
        cM.alignment = align_center

        # Col N: NCF (por asignar)
        cN = ws_agrupado.cell(row=curr_row, column=14, value="")
        cN.alignment = align_center

        # Col O: Observaciones
        cO = ws_agrupado.cell(row=curr_row, column=15, value=obs_text)
        cO.alignment = align_left
        if obs_text:
            cO.fill = fill_obs_warn

        # Aplicar bordes y fuentes a la fila
        for col_idx in range(1, 16):
            c_target = ws_agrupado.cell(row=curr_row, column=col_idx)
            c_target.font = font_regular
            c_target.border = border_cell

        curr_row += 1

    last_data_row_agrupado = curr_row - 1
    # Fila de TOTALES
    total_row = curr_row
    ws_agrupado.cell(row=total_row, column=3, value="TOTAL").font = font_bold
    ws_agrupado.cell(row=total_row, column=3).alignment = align_left

    if last_data_row_agrupado >= 7:
        ws_agrupado.cell(row=total_row, column=4, value=f"=SUM(D7:D{last_data_row_agrupado})").number_format = num_fmt_int
        for col_idx in range(5, 13):
            col_letter = get_column_letter(col_idx)
            ws_agrupado.cell(row=total_row, column=col_idx, value=f"=SUM({col_letter}7:{col_letter}{last_data_row_agrupado})").number_format = num_fmt_curr
    else:
        for c_idx in range(4, 13):
            ws_agrupado.cell(row=total_row, column=c_idx, value=0)

    for c_idx in range(1, 16):
        c_tot = ws_agrupado.cell(row=total_row, column=c_idx)
        c_tot.font = font_bold
        c_tot.fill = fill_total
        c_tot.border = border_total
        if 4 <= c_idx <= 12:
            c_tot.alignment = align_right

    ws_agrupado.freeze_panes = "A7"


    # =========================================================================
    # Hoja 2: Gastos menores B13
    # =========================================================================
    ws_b13 = wb.create_sheet(title="Gastos menores B13")
    headers_b13 = [
        "Fecha", "Proveedor", "Nota", "Importe Total", "Monto sujeto a retención", "ITBIS",
        "Monto Retención de renta", "Importe Efectivo", "Entidad", "Motivo", "NCF (por asignar)"
    ]
    ws_b13.append(headers_b13)

    for col_idx in range(1, len(headers_b13) + 1):
        cell = ws_b13.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    b13_records = [r for r in non_606_records if r.destino == "B13 Gastos menores"]
    b13_curr_row = 2

    for r in b13_records:
        tx = r.tx
        if tx.fecha_parsed:
            f_str = f"{tx.fecha_parsed[0]:04d}-{tx.fecha_parsed[1]:02d}-{tx.fecha_parsed[2]:02d}"
        else:
            f_str = tx.fecha_raw.split(" ")[0] if tx.fecha_raw else ""

        nota_limpia = tx.nota.replace("\r\n", " | ").replace("\n", " | ") if tx.nota else ""

        monto_pagado = float(tx.importe_sujeto)
        if r.tipo_gasto == "Servicio No Profesional":
            monto_sujeto = round(monto_pagado / (1.0 - t_no_prof), 2)
            ret_renta = round(monto_sujeto * t_no_prof, 2)
            itbis_b13 = 0.0 if r.itbis_exento else round(monto_sujeto * 0.18, 2)
        elif r.tipo_gasto == "Servicio Profesional":
            monto_sujeto = round(monto_pagado / (1.0 - t_prof), 2)
            ret_renta = round(monto_sujeto * t_prof, 2)
            itbis_b13 = 0.0 if r.itbis_exento else round(monto_sujeto * 0.18, 2)
        elif r.tipo_gasto == "Bien":
            monto_sujeto = monto_pagado
            ret_renta = 0.0
            itbis_b13 = 0.0 if r.itbis_exento else round(monto_sujeto * 0.18, 2)
        else:
            monto_sujeto = monto_pagado
            ret_renta = 0.0
            itbis_b13 = 0.0 if r.itbis_exento else float(tx.impuesto)
        importe_b13 = round(monto_sujeto + itbis_b13, 2)

        row_vals = [
            f_str,
            tx.proveedor,
            nota_limpia,
            monto_pagado,
            monto_sujeto,
            itbis_b13,
            ret_renta,
            importe_b13,
            tx.departamento,
            r.motivo_b13 or "Sin comprobante",
            ""
        ]
        ws_b13.append(row_vals)

        for col_idx in range(1, len(headers_b13) + 1):
            cell = ws_b13.cell(row=b13_curr_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_cell
            if col_idx in (4, 5, 6, 7, 8):
                cell.number_format = num_fmt_curr
                cell.alignment = align_right
            elif col_idx in (1, 9, 10, 11):
                cell.alignment = align_center

        b13_curr_row += 1

    last_data_row_b13 = b13_curr_row - 1
    total_row_b13 = b13_curr_row

    ws_b13.cell(row=total_row_b13, column=3, value="TOTAL").font = font_bold
    ws_b13.cell(row=total_row_b13, column=3).alignment = align_left

    if last_data_row_b13 >= 2:
        for c_idx in (4, 5, 6, 7, 8):
            c_letter = get_column_letter(c_idx)
            ws_b13.cell(row=total_row_b13, column=c_idx, value=f"=SUM({c_letter}2:{c_letter}{last_data_row_b13})").number_format = num_fmt_curr
    else:
        for c_idx in (4, 5, 6, 7, 8):
            ws_b13.cell(row=total_row_b13, column=c_idx, value=0).number_format = num_fmt_curr

    for c_idx in range(1, len(headers_b13) + 1):
        c_tot = ws_b13.cell(row=total_row_b13, column=c_idx)
        c_tot.font = font_bold
        c_tot.fill = fill_total
        c_tot.border = border_total
        if c_idx in (4, 5, 6, 7, 8):
            c_tot.alignment = align_right

    ws_b13.freeze_panes = "A2"

    # Reordenar hojas para que queden en el orden exacto especificado
    # 1. Agrupado por proveedor, 2. Gastos menores B13, 3. Detalle
    wb._sheets = [ws_agrupado, ws_b13, ws_detalle]

    # Ajuste automático de anchos de columnas y visualización de líneas de cuadrícula
    for ws in [ws_agrupado, ws_b13, ws_detalle]:
        try:
            ws.views.sheetView[0].showGridLines = True
        except Exception:
            pass

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                val = cell.value
                if val is not None:
                    # Omitir fórmulas largas
                    s_val = str(val)
                    if s_val.startswith("="):
                        val_len = 12
                    else:
                        val_len = len(s_val.split("\n")[0])
                    if val_len > max_len:
                        max_len = val_len
            ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

    wb.save(output_path)


def write_incidences_report(
    incidences: List[Incidence],
    total_tx: int,
    total_606: int,
    total_non_606: int,
    output_path: Path
) -> str:
    """
    Genera el reporte de incidencias en texto plano (incidencias.txt) y retorna el resumen.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    errores = [i for i in incidences if i.severidad == "ERROR"]
    advertencias = [i for i in incidences if i.severidad == "ADVERTENCIA"]
    infos = [i for i in incidences if i.severidad == "INFO"]

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines = [
        "=" * 80,
        "REPORTE DE INCIDENCIAS Y AUDITORÍA - FORMATO 606 Y COMPRAS SIN COMPROBANTE",
        f"Fecha y hora de generación: {timestamp}",
        "=" * 80,
        "",
        "RESUMEN DE EJECUCIÓN:",
        f"  Total de transacciones leídas : {total_tx}",
        f"  Comprobantes válidos (606)    : {total_606}",
        f"  Sin comprobante válido        : {total_non_606}",
        f"  Errores críticos detectados   : {len(errores)}",
        f"  Advertencias                  : {len(advertencias)}",
        f"  Notas informativas            : {len(infos)}",
        "",
        "-" * 80,
    ]

    if errores:
        lines.append("ERRORES CRÍTICOS:")
        for e in errores:
            lines.append(f"  • {e}")
        lines.append("")

    if advertencias:
        lines.append("ADVERTENCIAS:")
        for w in advertencias:
            lines.append(f"  • {w}")
        lines.append("")

    if infos:
        lines.append("INFORMACIÓN / AUDITORÍA:")
        for inf in infos:
            lines.append(f"  • {inf}")
        lines.append("")

    if not incidences:
        lines.append("No se detectaron incidencias ni advertencias. Los datos están 100% conformes.")
        lines.append("")

    lines.append("=" * 80)
    content = "\n".join(lines)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    return content
