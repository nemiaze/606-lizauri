"""
Procesador central: carga de datos, cruce de proveedores, clasificación y ruteo.
"""

from collections import defaultdict
from decimal import Decimal
import difflib
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd

from .config import AppConfig
from .models import Incidence, ProcessedRecord, ProviderMaster, Transaction
from .normalizer import (
    classify_bien_servicio,
    extract_rnc_cedula_from_note,
    is_itbis_exento,
    is_valid_ncf,
    normalize_text,
    parse_date,
    to_decimal,
    validate_and_clean_rnc_cedula,
    verify_cedula_jce,
)


def find_column_name(
    df: pd.DataFrame,
    target_name: str,
    alternatives: Optional[List[str]] = None
) -> Optional[str]:
    """
    Busca una columna en un DataFrame de forma insensible a mayúsculas y acentos.
    """
    if target_name in df.columns:
        return target_name

    norm_target = normalize_text(target_name)
    for col in df.columns:
        if normalize_text(str(col)) == norm_target:
            return str(col)

    if alternatives:
        for alt in alternatives:
            if alt in df.columns:
                return alt
            norm_alt = normalize_text(alt)
            for col in df.columns:
                if normalize_text(str(col)) == norm_alt:
                    return str(col)

    return None


def load_providers(
    proveedores_path: Path,
    hoja_name: str = "Hoja1",
    columnas_cfg: Optional[Dict[str, str]] = None
) -> Dict[str, ProviderMaster]:
    """
    Carga el maestro de proveedores desde Excel en un diccionario indexado por nombre normalizado.
    """
    if not proveedores_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de proveedores: {proveedores_path}")

    columnas_cfg = columnas_cfg or {}
    col_prov_key = columnas_cfg.get("proveedor", "Proveedor")
    col_rnc_key = columnas_cfg.get("rnc_cedula", "RNC/CEDULA")
    col_rs_key = columnas_cfg.get("razon_social", "RAZON SOCIAL")
    col_tb_key = columnas_cfg.get("tipo_bienes_servicios", "Tipo Bienes y Servicios Comprados")

    df = pd.read_excel(proveedores_path, sheet_name=hoja_name, dtype=str)

    col_prov = find_column_name(df, col_prov_key, ["Proveedor", "Nombre Proveedor", "Supplier"])
    col_rnc = find_column_name(df, col_rnc_key, ["RNC/CEDULA", "RNC", "CEDULA", "RNC_CEDULA"])
    col_rs = find_column_name(df, col_rs_key, ["RAZON SOCIAL", "Razon Social", "Nombre"])
    col_tb = find_column_name(df, col_tb_key, ["Tipo Bienes y Servicios Comprados", "Tipo Bienes", "Tipo Gasto"])

    if not col_prov:
        raise ValueError(f"Falta columna de Proveedor en '{proveedores_path.name}'. Columnas encontradas: {list(df.columns)}")
    if not col_rnc:
        raise ValueError(f"Falta columna de RNC/CEDULA en '{proveedores_path.name}'. Columnas encontradas: {list(df.columns)}")

    providers: Dict[str, ProviderMaster] = {}

    for idx, row in df.iterrows():
        raw_prov = str(row[col_prov]).strip() if pd.notna(row[col_prov]) else ""
        if not raw_prov:
            continue

        norm_prov = normalize_text(raw_prov)
        raw_rnc = str(row[col_rnc]).strip() if pd.notna(row[col_rnc]) else ""
        raw_rs = str(row[col_rs]).strip() if col_rs and pd.notna(row[col_rs]) else ""
        raw_tb = str(row[col_tb]).strip() if col_tb and pd.notna(row[col_tb]) else ""

        is_valid_id, clean_id, tipo_id, err_msg = validate_and_clean_rnc_cedula(raw_rnc)

        # Si trae código de tipo de bien/servicio, verificar si es válido (01 a 11)
        valid_tb = None
        if raw_tb:
            tb_digits = "".join(filter(str.isdigit, raw_tb)).zfill(2)
            if tb_digits in [f"{i:02d}" for i in range(1, 12)]:
                valid_tb = tb_digits

        p_master = ProviderMaster(
            proveedor=raw_prov,
            proveedor_norm=norm_prov,
            rnc_cedula_raw=raw_rnc,
            rnc_cedula_clean=clean_id,
            tipo_id=tipo_id,
            razon_social=raw_rs,
            tipo_bienes_servicios=valid_tb,
            is_valid_id=is_valid_id,
            validation_error=err_msg,
        )

        # Guardar bajo el nombre normalizado
        providers[norm_prov] = p_master

    return providers


def load_transactions(
    transacciones_path: Path,
    columnas_cfg: Optional[Dict[str, str]] = None
) -> List[Transaction]:
    """
    Carga el archivo de transacciones desde Excel conservando tipos y fechas.
    """
    if not transacciones_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de transacciones: {transacciones_path}")

    columnas_cfg = columnas_cfg or {}
    df = pd.read_excel(transacciones_path, dtype=object)

    c_fecha = find_column_name(df, columnas_cfg.get("fecha", "Fecha"), ["Fecha", "Date"])
    c_tipo_tx = find_column_name(df, columnas_cfg.get("tipo_transaccion", "Tipo de transacción"), ["Tipo de transaccion", "Tipo Transaccion", "Tipo"])
    c_num = find_column_name(df, columnas_cfg.get("numero", "Número"), ["Numero", "NCF", "Comprobante", "Documento"])
    c_prov = find_column_name(df, columnas_cfg.get("proveedor", "Proveedor"), ["Proveedor", "Supplier"])
    c_prod = find_column_name(df, columnas_cfg.get("producto_servicio", "Producto/servicio"), ["Producto/servicio", "Producto", "Servicio"])
    c_nota = find_column_name(df, columnas_cfg.get("nota", "Nota"), ["Nota", "Notas", "Memo", "Descripcion"])
    c_imp = find_column_name(df, columnas_cfg.get("importe", "Importe"), ["Importe", "Total", "Monto Total"])
    c_itbis = find_column_name(df, columnas_cfg.get("impuesto", "Impuesto"), ["Impuesto", "ITBIS", "Impuestos"])
    c_dpto = find_column_name(df, columnas_cfg.get("departamento", "$$Departamento$$"), ["$$Departamento$$", "Departamento", "Entidad", "Empresa"])
    c_subt = find_column_name(df, columnas_cfg.get("importe_sujeto_impuestos", "Importe sujeto a impuestos"), ["Importe sujeto a impuestos", "Subtotal", "Base Imponible"])
    c_cp = find_column_name(df, columnas_cfg.get("cp_pagadas", "C/P pagadas"), ["C/P pagadas", "CP Pagadas", "Estado Pago", "Estado"])

    # Validar obligatorias mínimas
    if not c_prov:
        raise ValueError(f"Falta columna de Proveedor en '{transacciones_path.name}'. Columnas: {list(df.columns)}")
    if not c_imp:
        raise ValueError(f"Falta columna de Importe en '{transacciones_path.name}'. Columnas: {list(df.columns)}")
    if not c_subt:
        raise ValueError(f"Falta columna de Importe sujeto a impuestos en '{transacciones_path.name}'. Columnas: {list(df.columns)}")

    transactions: List[Transaction] = []

    for idx, row in df.iterrows():
        fila_excel = idx + 2  # 1-based, fila 1 encabezados

        raw_fecha = row[c_fecha] if c_fecha else None
        fecha_parsed = parse_date(raw_fecha) if raw_fecha is not None else None

        raw_tipo_tx = str(row[c_tipo_tx]).strip() if c_tipo_tx and pd.notna(row[c_tipo_tx]) else ""
        raw_num = str(row[c_num]).strip() if c_num and pd.notna(row[c_num]) else ""
        raw_prov = str(row[c_prov]).strip() if pd.notna(row[c_prov]) else ""
        norm_prov = normalize_text(raw_prov)

        raw_prod = str(row[c_prod]).strip() if c_prod and pd.notna(row[c_prod]) else ""
        raw_nota = str(row[c_nota]).strip() if c_nota and pd.notna(row[c_nota]) else ""
        
        importe_dec = to_decimal(row[c_imp]) if c_imp else Decimal("0.00")
        itbis_dec = to_decimal(row[c_itbis]) if c_itbis else Decimal("0.00")
        
        # Subtotal viene en negativo en el archivo: se toma valor absoluto
        raw_subt_val = row[c_subt] if c_subt else "0.00"
        subt_dec = abs(to_decimal(raw_subt_val))

        raw_dpto = str(row[c_dpto]).strip() if c_dpto and pd.notna(row[c_dpto]) else ""
        raw_cp = str(row[c_cp]).strip() if c_cp and pd.notna(row[c_cp]) else ""

        tx = Transaction(
            fila=fila_excel,
            fecha_raw=str(raw_fecha) if raw_fecha is not None else "",
            fecha_parsed=fecha_parsed,
            tipo_transaccion=raw_tipo_tx,
            numero=raw_num,
            proveedor=raw_prov,
            proveedor_norm=norm_prov,
            producto_servicio=raw_prod,
            nota=raw_nota,
            importe=importe_dec,
            impuesto=itbis_dec,
            departamento=raw_dpto,
            importe_sujeto=subt_dec,
            cp_pagadas=raw_cp,
        )
        transactions.append(tx)

    return transactions


def process_records(
    transactions: List[Transaction],
    providers: Dict[str, ProviderMaster],
    config: AppConfig
) -> Tuple[List[ProcessedRecord], List[Incidence]]:
    """
    Aplica el árbol de decisión a cada transacción, generando los registros procesados
    y el reporte completo de incidencias.
    """
    records: List[ProcessedRecord] = []
    incidences: List[Incidence] = []

    # Para verificar proveedores en múltiples entidades
    prov_departamentos: Dict[str, Set[str]] = defaultdict(set)
    prov_display_names: Dict[str, str] = {}

    for tx in transactions:
        norm_prov = tx.proveedor_norm
        if tx.proveedor:
            prov_display_names[norm_prov] = tx.proveedor
            if tx.departamento:
                prov_departamentos[norm_prov].add(tx.departamento)

        # -------------------------------------------------------------
        # 1. Validaciones generales e integridad
        # -------------------------------------------------------------
        # Descuadre de integridad: abs(Subtotal) + ITBIS == Importe
        suma_calculada = (tx.importe_sujeto + tx.impuesto).quantize(Decimal("0.01"))
        diferencia = abs(suma_calculada - tx.importe)
        if diferencia > Decimal("0.01"):
            incidences.append(Incidence(
                tipo="DESCUADRE_INTEGRIDAD",
                fila=tx.fila,
                proveedor=tx.proveedor,
                descripcion=(
                    f"Descuadre en integridad de importes: Subtotal ({tx.importe_sujeto}) + "
                    f"ITBIS ({tx.impuesto}) = {suma_calculada} != Importe ({tx.importe}) "
                    f"[Diferencia: {diferencia}]"
                ),
                severidad="ERROR"
            ))

        # Validación de tipo de transacción
        if tx.tipo_transaccion and normalize_text(tx.tipo_transaccion) != "factura de proveedor":
            incidences.append(Incidence(
                tipo="TIPO_TRANSACCION_DISTINTO",
                fila=tx.fila,
                proveedor=tx.proveedor,
                descripcion=f"Tipo de transacción '{tx.tipo_transaccion}' distinto del esperado 'Factura de proveedor'.",
                severidad="ADVERTENCIA"
            ))

        # Validación de período fiscal
        if config.periodo and tx.fecha_parsed:
            año, mes, _ = tx.fecha_parsed
            periodo_registro = f"{año}{mes:02d}"
            if periodo_registro != config.periodo:
                incidences.append(Incidence(
                    tipo="PERIODO_FUERA_DE_RANGO",
                    fila=tx.fila,
                    proveedor=tx.proveedor,
                    descripcion=f"Fecha del registro ({periodo_registro}) fuera del período declarado ({config.periodo}).",
                    severidad="ADVERTENCIA"
                ))

        # Detectar exención de ITBIS según la Nota o reglas por proveedor
        exento = is_itbis_exento(tx.nota, config.keywords_itbis_exento)
        if not exento and any(p in norm_prov for p in config.proveedores_exento_itbis):
            exento = True
        if exento and tx.impuesto > 0:
            incidences.append(Incidence(
                tipo="ITBIS_EN_PRODUCTO_EXENTO",
                fila=tx.fila,
                proveedor=tx.proveedor,
                descripcion=(
                    f"El registro tiene ITBIS ({tx.impuesto}) pero el producto/servicio "
                    f"en la Nota parece ser exento según la DGII. ITBIS corregido a 0."
                ),
                severidad="ADVERTENCIA"
            ))

        # Búsqueda en el maestro
        p_master = providers.get(norm_prov)
        if not p_master:
            # Buscar sugerencia de coincidencia cercana
            matches = difflib.get_close_matches(norm_prov, providers.keys(), n=1, cutoff=0.8)
            sug = f" (Sugerencia similar: '{providers[matches[0]].proveedor}')" if matches else ""
            incidences.append(Incidence(
                tipo="PROVEEDOR_NO_EN_MAESTRO",
                fila=tx.fila,
                proveedor=tx.proveedor,
                descripcion=f"Proveedor no existe en el maestro de proveedores{sug}.",
                severidad="ERROR"
            ))

        # -------------------------------------------------------------
        # 2. Árbol de decisión: ¿Tiene NCF válido?
        # -------------------------------------------------------------
        is_ncf, clean_ncf = is_valid_ncf(tx.numero)

        if is_ncf:
            # ---> Va al FORMATO 606 (Salida A)
            rnc_valido = None
            tipo_id_val = None

            if p_master and p_master.is_valid_id:
                rnc_valido = p_master.rnc_cedula_clean
                tipo_id_val = p_master.tipo_id
            else:
                # Si el NCF es tipo B13, no requiere RNC obligatorio
                es_ncf_b13 = clean_ncf.startswith("B13")
                if not es_ncf_b13:
                    incidences.append(Incidence(
                        tipo="NCF_SIN_RNC_EN_MAESTRO",
                        fila=tx.fila,
                        proveedor=tx.proveedor,
                        descripcion=f"Factura con NCF válido '{clean_ncf}' pero proveedor no tiene RNC/cédula válido en el maestro.",
                        severidad="ERROR"
                    ))

            rec = ProcessedRecord(
                tx=tx,
                is_606=True,
                ncf_valido=clean_ncf,
                rnc_cedula=rnc_valido,
                tipo_id=tipo_id_val,
                fuente_rnc="Maestro" if rnc_valido else "N/A",
                tipo_gasto="Bien",
                destino="606",
                observaciones=[],
                itbis_exento=exento
            )
            records.append(rec)

        else:
            # ---> Registro SIN COMPROBANTE VÁLIDO (Salida B)
            # ¿Es DJ Ludo o excepción B13?
            es_excepcion_b13 = any(exc in norm_prov for exc in config.proveedores_b13)

            if es_excepcion_b13:
                # Va directo a GASTOS MENORES (B13)
                rec = ProcessedRecord(
                    tx=tx,
                    is_606=False,
                    ncf_valido=None,
                    rnc_cedula=None,
                    tipo_id=None,
                    fuente_rnc="N/A",
                    tipo_gasto="Sin clasificar",
                    destino="B13 Gastos menores",
                    motivo_b13="Excepción DJ Ludo",
                    observaciones=["Excepción DJ Ludo"],
                    itbis_exento=exento
                )
                records.append(rec)

            else:
                # Buscar RNC / Cédula en maestro y nota
                rnc_maestro = p_master.rnc_cedula_clean if (p_master and p_master.is_valid_id) else None
                tipo_maestro = p_master.tipo_id if (p_master and p_master.is_valid_id) else None

                rnc_nota, tipo_nota = extract_rnc_cedula_from_note(tx.nota)

                observaciones_reg: List[str] = []

                # Evaluar conflicto entre maestro y nota
                if rnc_maestro and rnc_nota:
                    if rnc_maestro != rnc_nota:
                        obs_conflicto = f"Maestro ({rnc_maestro}) y nota ({rnc_nota}) no coinciden; se usó el maestro"
                        observaciones_reg.append(obs_conflicto)
                        incidences.append(Incidence(
                            tipo="CONFLICTO_IDENTIFICACION",
                            fila=tx.fila,
                            proveedor=tx.proveedor,
                            descripcion=f"Maestro ({rnc_maestro}) y nota ({rnc_nota}) no coinciden; prevalece el maestro.",
                            severidad="ADVERTENCIA"
                        ))

                # Determinar RNC final (regla: maestro primero)
                if rnc_maestro:
                    rnc_final = rnc_maestro
                    tipo_id_final = tipo_maestro
                    fuente_final = "Maestro"
                elif rnc_nota:
                    rnc_final = rnc_nota
                    tipo_id_final = tipo_nota
                    fuente_final = "Nota"
                else:
                    rnc_final = None
                    tipo_id_final = None
                    fuente_final = "N/A"

                # REGLA B13:
                # 1. Aquellos que no se encuentran en el maestro de proveedores,
                #    no tienen NCF, no tienen información en la columna Nota,
                #    y no se especifica producto/servicio -> Gasto menor B13.
                # 2. Registros sin RNC/Cédula válido ni en maestro ni en nota -> Gasto menor B13.
                sin_nota = not tx.nota or not str(tx.nota).strip()
                sin_prod = not tx.producto_servicio or not str(tx.producto_servicio).strip()
                no_en_maestro = (p_master is None or not p_master.is_valid_id)
                es_b13_por_falta_datos = (no_en_maestro and sin_nota and sin_prod)

                if (not rnc_final) or es_b13_por_falta_datos:
                    motivo = "Sin RNC/Cédula en maestro ni nota"
                    if es_b13_por_falta_datos:
                        motivo = "Proveedor no en maestro, sin NCF ni detalle en nota/producto"
                    rec = ProcessedRecord(
                        tx=tx,
                        is_606=False,
                        ncf_valido=None,
                        rnc_cedula=None,
                        tipo_id=None,
                        fuente_rnc="N/A",
                        tipo_gasto="Sin clasificar",
                        destino="B13 Gastos menores",
                        motivo_b13=motivo,
                        observaciones=[motivo],
                        itbis_exento=exento
                    )
                    records.append(rec)

                else:
                    # Encontrado -> AGRUPADO POR PROVEEDOR
                    # Determinar tipo_gasto:
                    # 1. Reglas específicas por proveedor
                    tipo_gasto = None
                    if any(p in norm_prov for p in config.proveedores_bien):
                        tipo_gasto = "Bien"
                    elif any(p in norm_prov for p in config.proveedores_servicio_profesional):
                        tipo_gasto = "Servicio Profesional"
                    elif any(p in norm_prov for p in config.proveedores_servicio_no_profesional):
                        tipo_gasto = "Servicio No Profesional"

                    # 2. Si no coincide por proveedor, clasificar por Nota y Producto/Servicio
                    if not tipo_gasto:
                        texto_eval = tx.nota
                        if not texto_eval or not str(texto_eval).strip():
                            texto_eval = tx.producto_servicio
                        tipo_gasto = classify_bien_servicio(
                            texto_eval,
                            config.keywords_servicio_profesional,
                            config.keywords_servicio_no_profesional,
                            config.keywords_bien
                        )

                    if tipo_gasto == "Sin clasificar":
                        obs_sin_clasificar = "nota no indica bien o servicio" if observaciones_reg else "Nota no indica bien o servicio"
                        observaciones_reg.append(obs_sin_clasificar)
                        incidences.append(Incidence(
                            tipo="NOTA_SIN_CLASIFICAR",
                            fila=tx.fila,
                            proveedor=tx.proveedor,
                            descripcion="La nota no indica si es bien o servicio. Asignado a 'Sin clasificar'.",
                            severidad="INFO"
                        ))

                    # Validar dígito verificador de cédula si es tipo 2
                    if tipo_id_final == "2":
                        is_valid_jce, exp_v, act_v = verify_cedula_jce(rnc_final)
                        if not is_valid_jce:
                            incidences.append(Incidence(
                                tipo="CEDULA_DIGITO_INVALIDO",
                                fila=tx.fila,
                                proveedor=tx.proveedor,
                                descripcion=f"Cédula '{rnc_final}' tiene dígito verificador JCE inválido ({act_v} != esperado {exp_v}).",
                                severidad="ADVERTENCIA"
                            ))

                    rec = ProcessedRecord(
                        tx=tx,
                        is_606=False,
                        ncf_valido=None,
                        rnc_cedula=rnc_final,
                        tipo_id=tipo_id_final,
                        fuente_rnc=fuente_final,
                        tipo_gasto=tipo_gasto,
                        destino="Agrupado",
                        observaciones=observaciones_reg,
                        itbis_exento=exento
                    )
                    records.append(rec)

    # -----------------------------------------------------------------
    # 3. Validar proveedores que aparecen en más de una entidad
    # -----------------------------------------------------------------
    for norm_p, dptos in prov_departamentos.items():
        if len(dptos) > 1:
            disp_name = prov_display_names.get(norm_p, norm_p)
            lista_dptos = ", ".join(sorted(dptos))
            incidences.append(Incidence(
                tipo="PROVEEDOR_MULTI_ENTIDAD",
                fila=None,
                proveedor=disp_name,
                descripcion=f"Proveedor aparece en múltiples entidades / departamentos ({len(dptos)}): {lista_dptos}.",
                severidad="INFO"
            ))

    return records, incidences
