#!/usr/bin/env python3
"""
Generador del Formato 606 (DGII) y Agrupación de Compras sin Comprobante.

Uso básico:
    python generar_606.py --transacciones muestra.xlsx --proveedores Proveedores.xlsx --salida ./salida

Modo interactivo / gráfico:
    python generar_606.py --gui
"""

import argparse
import os
from pathlib import Path
import sys

from generador_606.config import AppConfig
from generador_606.processor import load_providers, load_transactions, process_records
from generador_606.writers import write_606_csv, write_excel_agrupacion, write_incidences_report


def run_pipeline(config: AppConfig) -> int:
    """
    Ejecuta el flujo completo de generación.
    """
    print("\n" + "=" * 75)
    print("  GENERADOR DEL FORMATO 606 (DGII) Y COMPRAS SIN COMPROBANTE")
    print("=" * 75)
    print(f"Transacciones : {config.transacciones_path}")
    print(f"Proveedores   : {config.proveedores_path}")
    print(f"Directorio    : {config.salida_dir}")
    print(f"Tasa Retención: {float(config.tasa_retencion) * 100:.1f}%")
    if config.periodo:
        print(f"Período Fiscal: {config.periodo}")
    print("-" * 75)

    # 1. Cargar proveedores
    print("Cargando maestro de proveedores...")
    try:
        providers = load_providers(
            config.proveedores_path,
            hoja_name=config.hoja_proveedores,
            columnas_cfg=config.columnas_proveedores
        )
        print(f"  -> {len(providers):,} proveedores cargados e indexados.")
    except Exception as e:
        print(f"[ERROR CRÍTICO] No se pudo cargar el archivo de proveedores: {e}")
        return 1

    # 2. Cargar transacciones
    print("Cargando archivo de transacciones...")
    try:
        transactions = load_transactions(
            config.transacciones_path,
            columnas_cfg=config.columnas_transacciones
        )
        print(f"  -> {len(transactions):,} transacciones cargadas.")
    except Exception as e:
        print(f"[ERROR CRÍTICO] No se pudo cargar el archivo de transacciones: {e}")
        return 1

    # 3. Procesar y clasificar
    print("Procesando transacciones y evaluando reglas fiscales...")
    records, incidences = process_records(transactions, providers, config)

    rec_606 = [r for r in records if r.is_606]
    rec_agrupados = [r for r in records if not r.is_606 and r.destino == "Agrupado"]
    rec_b13 = [r for r in records if not r.is_606 and r.destino == "B13 Gastos menores"]

    print(f"  -> Registros con NCF válido (606) : {len(rec_606)}")
    print(f"  -> Registros sin NCF (Agrupados)  : {len(rec_agrupados)}")
    print(f"  -> Registros sin NCF (Gastos B13) : {len(rec_b13)}")

    # 4. Generar salidas
    config.salida_dir.mkdir(parents=True, exist_ok=True)

    path_606 = config.salida_dir / config.archivo_606
    path_agrupacion = config.salida_dir / config.archivo_agrupacion
    path_incidencias = config.salida_dir / config.archivo_incidencias

    print(f"\nEscribiendo archivo Formato 606: {path_606.name}...")
    write_606_csv(records, providers, path_606, config)

    print(f"Escribiendo libro Excel: {path_agrupacion.name}...")
    write_excel_agrupacion(records, path_agrupacion, config)

    print(f"Escribiendo reporte de incidencias: {path_incidencias.name}...")
    report_text = write_incidences_report(
        incidences=incidences,
        total_tx=len(transactions),
        total_606=len(rec_606),
        total_non_606=len(rec_agrupados) + len(rec_b13),
        output_path=path_incidencias
    )

    # 5. Resumen final en consola
    print("\n" + "=" * 75)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 75)
    print(f"  1. Archivo 606 generado     : {path_606.resolve()}")
    print(f"  2. Archivo Agrupación Excel : {path_agrupacion.resolve()}")
    print(f"  3. Reporte de Incidencias   : {path_incidencias.resolve()}")
    print("-" * 75)

    errores = [i for i in incidences if i.severidad == "ERROR"]
    advertencias = [i for i in incidences if i.severidad == "ADVERTENCIA"]

    if errores:
        print(f"[!] Se encontraron {len(errores)} ERRORES que requieren revisión:")
        for e in errores[:5]:
            print(f"    • {e}")
        if len(errores) > 5:
            print(f"    ... y {len(errores) - 5} más (ver {path_incidencias.name})")

    if advertencias:
        print(f"[i] Se registraron {len(advertencias)} advertencias:")
        for w in advertencias[:5]:
            print(f"    • {w}")
        if len(advertencias) > 5:
            print(f"    ... y {len(advertencias) - 5} más (ver {path_incidencias.name})")

    if not errores and not advertencias:
        print("[OK] Todos los registros cumplen con las validaciones de integridad.")

    print("=" * 75 + "\n")
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Generador del formato DGII 606 y agrupación de compras sin comprobante."
    )
    parser.add_argument(
        "-t", "--transacciones",
        help="Ruta al archivo de transacciones (Excel). Por defecto busca muestra.xlsx o grande.xlsx."
    )
    parser.add_argument(
        "-p", "--proveedores",
        help="Ruta al maestro de proveedores (Excel). Por defecto Proveedores.xlsx."
    )
    parser.add_argument(
        "-s", "--salida",
        help="Directorio de salida para los archivos generados. Por defecto ./salida."
    )
    parser.add_argument(
        "-r", "--tasa-retencion",
        type=float,
        help="Tasa de retención sobre servicios (ej. 0.10 para 10%%)."
    )
    parser.add_argument(
        "--periodo",
        help="Período fiscal en formato AAAAMM (ej. 202608) para validar fechas."
    )
    parser.add_argument(
        "-c", "--config",
        default="config.yaml",
        help="Ruta al archivo de configuración YAML (por defecto config.yaml)."
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Inicia la interfaz gráfica de usuario (GUI)."
    )

    args = parser.parse_args()

    if args.gui:
        try:
            from app_gui import launch_gui
            launch_gui()
            return
        except ImportError as e:
            print(f"Error al iniciar GUI: {e}. Ejecutando en modo consola...")

    # Cargar configuración con flags de línea de comandos
    config = AppConfig.load(
        config_path=args.config,
        transacciones_path=args.transacciones,
        proveedores_path=args.proveedores,
        salida_dir=args.salida,
        tasa_retencion=args.tasa_retencion,
        periodo=args.periodo
    )

    exit_code = run_pipeline(config)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
