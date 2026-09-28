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

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        # Escribir encabezado
        f.write(",".join(HEADERS_606) + "\r\n")

        for rec in records:
            if not rec.is_606:
                continue

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
            ncf_mod = ""

            # 6 y 7. Fecha Comprobante (AAAAMM y DD)
            if tx.fecha_parsed:
                a, m, d = tx.fecha_parsed
                f_aaaamm = f"{a:04d}{m:02d}"
                f_dd = f"{d:02d}"
            else:
                f_aaaamm = ""
                f_dd = ""

            # 8 y 9. Fecha Pago (vacías)
            f_pago_aaaamm = ""
            f_pago_dd = ""

            # 10 y 11. Monto Servicios vs Bienes
            subtotal = tx.importe_sujeto
            if config.bienes_vs_servicios_606.lower() == "servicios":
                monto_servicios = f"{subtotal:.2f}"
                monto_bienes = "0.00"
            else:
                monto_servicios = "0.00"
                monto_bienes = f"{subtotal:.2f}"

            # 12. Total Monto Facturado
            total_facturado = f"{subtotal:.2f}"

            # 13. ITBIS Facturado
            itbis_facturado = f"{tx.impuesto:.2f}"

            # 14. ITBIS Retenido
            itbis_retenido = ""

            # 15. ITBIS Proporcionalidad
            itbis_prop = "0.00"

            # 16. ITBIS llevado al Costo
            itbis_costo = "0.00"

            # 17. ITBIS por Adelantar = ITBIS Facturado - Prop - Costo
            itbis_adelantar = f"{tx.impuesto:.2f}"

            # 18 a 24: vacías
            itbis_perc = ""
            tipo_ret_isr = ""
            monto_ret_renta = ""
            isr_perc = ""
            isc = ""
            otros_imp = ""
            propina = ""

            # 25. Forma de Pago (02 = Pagado, 04 = Crédito u otro)
            if tx.cp_pagadas and tx.cp_pagadas.strip().lower() == "pagado":
                forma_pago = config.forma_pago_pagado
            else:
                forma_pago = config.forma_pago_credito

            row_data = [
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

            f.write(",".join(row_data) + "\r\n")

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
        "Destino", "Nota", "Subtotal", "ITBIS", "Importe", "Entidad",
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
    for r in non_606_records:
        tx = r.tx
        # Fecha formateada yyyy-mm-dd
        if tx.fecha_parsed:
            f_str = f"{tx.fecha_parsed[0]:04d}-{tx.fecha_parsed[1]:02d}-{tx.fecha_parsed[2]:02d}"
        else:
            f_str = tx.fecha_raw.split(" ")[0] if tx.fecha_raw else ""

        nota_limpia = tx.nota.replace("\r\n", " | ").replace("\n", " | ") if tx.nota else ""
        obs_str = "; ".join(r.observaciones) if r.observaciones else ""

        row_vals = [
            f_str,
            tx.proveedor,
            r.rnc_cedula or "",
            r.fuente_rnc,
            r.tipo_gasto,
            r.destino,
            nota_limpia,
            float(tx.importe_sujeto),
            float(tx.impuesto),
            float(tx.importe),
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
            if col_idx in (8, 9, 10):  # Subtotal, ITBIS, Importe
                cell.number_format = num_fmt_curr
                cell.alignment = align_right
            elif col_idx in (1, 3, 4, 5, 6, 12):  # Fecha, RNC, etc.
                cell.alignment = align_center
                if col_idx == 3:
                    cell.number_format = "@"

    ws_detalle.freeze_panes = "A2"

    # =========================================================================
    # Hoja 1: Agrupado por proveedor
    # =========================================================================
    ws_agrupado = wb.active
    ws_agrupado.title = "Agrupado por proveedor"

    # Encabezado superior / Tasa de retención
    ws_agrupado.cell(row=1, column=1, value="AGRUPACIÓN DE COMPRAS SIN COMPROBANTE (DGII)").font = font_title
    
    lbl_rate = ws_agrupado.cell(row=3, column=1, value="Tasa de retención:")
    lbl_rate.font = font_bold
    lbl_rate.alignment = align_right

    cell_rate = ws_agrupado.cell(row=3, column=2, value=float(config.tasa_retencion))
    cell_rate.font = font_rate
    cell_rate.fill = fill_yellow_rate
    cell_rate.alignment = align_center
    cell_rate.number_format = "0.0%"
    cell_rate.border = border_rate
    cell_rate.comment = Comment(
        "Tasa de retención editable aplicable a servicios.\n"
        "Modifique este porcentaje para recalcular automáticamente las retenciones.",
        "Generador 606"
    )

    headers_agrupado = [
        "RNC o Cédula", "Tipo Id", "Proveedor", "Registros",
        "Servicios (sujeto a retención)", "Bienes", "Sin clasificar",
        "Total facturado (subtotal)", "ITBIS", "Retención de renta",
        "Tipo Retención ISR", "NCF (por asignar)", "Observaciones"
    ]

    header_row_agrupado = 5
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

    # Insertar filas de proveedores agrupados
    curr_row = 6
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

        # Col D: Registros: COUNTIFS(Detalle!$B:$B, C{row}, Detalle!$F:$F, "Agrupado")
        cD = ws_agrupado.cell(row=curr_row, column=4, value=f'=COUNTIFS(Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado")')
        cD.number_format = num_fmt_int
        cD.alignment = align_right

        # Col E: Servicios: SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Servicio")
        cE = ws_agrupado.cell(row=curr_row, column=5, value=f'=SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Servicio")')
        cE.number_format = num_fmt_curr
        cE.alignment = align_right

        # Col F: Bienes: SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Bien")
        cF = ws_agrupado.cell(row=curr_row, column=6, value=f'=SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Bien")')
        cF.number_format = num_fmt_curr
        cF.alignment = align_right

        # Col G: Sin clasificar: SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Sin clasificar")
        cG = ws_agrupado.cell(row=curr_row, column=7, value=f'=SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Sin clasificar")')
        cG.number_format = num_fmt_curr
        cG.alignment = align_right

        # Col H: Total facturado: =SUM(E{row}:G{row})
        cH = ws_agrupado.cell(row=curr_row, column=8, value=f'=SUM(E{curr_row}:G{curr_row})')
        cH.number_format = num_fmt_curr
        cH.alignment = align_right

        # Col I: ITBIS: =SUMIFS(Detalle!$I:$I, Detalle!$B:$B, C{row}, Detalle!$F:$F, "Agrupado")
        cI = ws_agrupado.cell(row=curr_row, column=9, value=f'=SUMIFS(Detalle!$I:$I, Detalle!$B:$B, C{curr_row}, Detalle!$F:$F, "Agrupado")')
        cI.number_format = num_fmt_curr
        cI.alignment = align_right

        # Col J: Retención renta: =ROUND(E{row} * $B$3, 2)
        cJ = ws_agrupado.cell(row=curr_row, column=10, value=f'=ROUND(E{curr_row} * $B$3, 2)')
        cJ.number_format = num_fmt_curr
        cJ.alignment = align_right

        # Col K: Tipo Retención ISR: =IF(E{row}>0, "02", "")
        cK = ws_agrupado.cell(row=curr_row, column=11, value=f'=IF(E{curr_row}>0, "{config.tipo_retencion_isr}", "")')
        cK.alignment = align_center

        # Col L: NCF (por asignar)
        cL = ws_agrupado.cell(row=curr_row, column=12, value="")
        cL.alignment = align_center

        # Col M: Observaciones
        cM = ws_agrupado.cell(row=curr_row, column=13, value=obs_text)
        cM.alignment = align_left
        if obs_text:
            cM.fill = fill_obs_warn

        # Aplicar bordes y fuentes a la fila
        for col_idx in range(1, 14):
            c_target = ws_agrupado.cell(row=curr_row, column=col_idx)
            c_target.font = font_regular
            c_target.border = border_cell

        curr_row += 1

    last_data_row_agrupado = curr_row - 1
    # Fila de TOTALES
    total_row = curr_row
    ws_agrupado.cell(row=total_row, column=3, value="TOTAL").font = font_bold
    ws_agrupado.cell(row=total_row, column=3).alignment = align_left

    if last_data_row_agrupado >= 6:
        ws_agrupado.cell(row=total_row, column=4, value=f"=SUM(D6:D{last_data_row_agrupado})").number_format = num_fmt_int
        ws_agrupado.cell(row=total_row, column=5, value=f"=SUM(E6:E{last_data_row_agrupado})").number_format = num_fmt_curr
        ws_agrupado.cell(row=total_row, column=6, value=f"=SUM(F6:F{last_data_row_agrupado})").number_format = num_fmt_curr
        ws_agrupado.cell(row=total_row, column=7, value=f"=SUM(G6:G{last_data_row_agrupado})").number_format = num_fmt_curr
        ws_agrupado.cell(row=total_row, column=8, value=f"=SUM(H6:H{last_data_row_agrupado})").number_format = num_fmt_curr
        ws_agrupado.cell(row=total_row, column=9, value=f"=SUM(I6:I{last_data_row_agrupado})").number_format = num_fmt_curr
        ws_agrupado.cell(row=total_row, column=10, value=f"=SUM(J6:J{last_data_row_agrupado})").number_format = num_fmt_curr
    else:
        for c_idx in range(4, 11):
            ws_agrupado.cell(row=total_row, column=c_idx, value=0)

    for c_idx in range(1, 14):
        c_tot = ws_agrupado.cell(row=total_row, column=c_idx)
        c_tot.font = font_bold
        c_tot.fill = fill_total
        c_tot.border = border_total
        if c_idx >= 4 and c_idx <= 10:
            c_tot.alignment = align_right

    ws_agrupado.freeze_panes = "A6"

    # =========================================================================
    # Hoja 2: Gastos menores B13
    # =========================================================================
    ws_b13 = wb.create_sheet(title="Gastos menores B13")
    headers_b13 = [
        "Fecha", "Proveedor", "Nota", "Subtotal", "ITBIS", "Importe",
        "Entidad", "Motivo", "NCF (por asignar)"
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

        row_vals = [
            f_str,
            tx.proveedor,
            nota_limpia,
            float(tx.importe_sujeto),
            float(tx.impuesto),
            float(tx.importe),
            tx.departamento,
            r.motivo_b13 or "Sin comprobante",
            ""
        ]
        ws_b13.append(row_vals)

        for col_idx in range(1, len(headers_b13) + 1):
            cell = ws_b13.cell(row=b13_curr_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_cell
            if col_idx in (4, 5, 6):
                cell.number_format = num_fmt_curr
                cell.alignment = align_right
            elif col_idx in (1, 8, 9):
                cell.alignment = align_center

        b13_curr_row += 1

    last_data_row_b13 = b13_curr_row - 1
    total_row_b13 = b13_curr_row

    ws_b13.cell(row=total_row_b13, column=3, value="TOTAL").font = font_bold
    ws_b13.cell(row=total_row_b13, column=3).alignment = align_left

    if last_data_row_b13 >= 2:
        ws_b13.cell(row=total_row_b13, column=4, value=f"=SUM(D2:D{last_data_row_b13})").number_format = num_fmt_curr
        ws_b13.cell(row=total_row_b13, column=5, value=f"=SUM(E2:E{last_data_row_b13})").number_format = num_fmt_curr
        ws_b13.cell(row=total_row_b13, column=6, value=f"=SUM(F2:F{last_data_row_b13})").number_format = num_fmt_curr
    else:
        for c_idx in (4, 5, 6):
            ws_b13.cell(row=total_row_b13, column=c_idx, value=0).number_format = num_fmt_curr

    for c_idx in range(1, len(headers_b13) + 1):
        c_tot = ws_b13.cell(row=total_row_b13, column=c_idx)
        c_tot.font = font_bold
        c_tot.fill = fill_total
        c_tot.border = border_total
        if c_idx in (4, 5, 6):
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
