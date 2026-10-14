"""
Pruebas del procesador con el caso de prueba oficial de la Sección 7.
"""

from decimal import Decimal
from pathlib import Path
import pytest

from generador_606.config import AppConfig
from generador_606.processor import load_providers, load_transactions, process_records


@pytest.fixture
def test_setup():
    cfg = AppConfig.load(
        config_path="config.yaml",
        transacciones_path="muestra.xlsx",
        proveedores_path="Proveedores.xlsx",
        salida_dir="./salida",
        tasa_retencion_profesional=0.15,
        tasa_retencion_no_profesional=0.03
    )
    providers = load_providers(cfg.proveedores_path, hoja_name=cfg.hoja_proveedores)
    transactions = load_transactions(cfg.transacciones_path)
    return cfg, providers, transactions


def test_seccion_7_processing(test_setup):
    cfg, providers, transactions = test_setup
    assert len(transactions) == 9

    records, incidences = process_records(transactions, providers, cfg)

    # 1. Separación de registros
    rec_606 = [r for r in records if r.is_606]
    rec_agrupados = [r for r in records if not r.is_606 and r.destino == "Agrupado"]
    rec_b13 = [r for r in records if not r.is_606 and r.destino == "B13 Gastos menores"]

    assert len(rec_606) == 3
    assert len(rec_agrupados) == 5
    assert len(rec_b13) == 1

    # 2. Validar registros 606
    # Fila 1: Agropin B0100000188
    r1 = rec_606[0]
    assert r1.ncf_valido == "B0100000188"
    assert r1.rnc_cedula == "06600056540"
    assert r1.tipo_id == "2"
    assert r1.tx.importe_sujeto == Decimal("2033.90")
    assert r1.tx.impuesto == Decimal("366.10")

    # Fila 2: Agropin B0100000204
    r2 = rec_606[1]
    assert r2.ncf_valido == "B0100000204"
    assert r2.rnc_cedula == "06600056540"
    assert r2.tipo_id == "2"
    assert r2.tx.importe_sujeto == Decimal("3389.83")
    assert r2.tx.impuesto == Decimal("610.17")

    # Fila 3: Alta Gama B0100000383
    r3 = rec_606[2]
    assert r3.ncf_valido == "B0100000383"
    assert r3.rnc_cedula == "132465555"
    assert r3.tipo_id == "1"
    assert r3.tx.importe_sujeto == Decimal("26957.64")
    assert r3.tx.impuesto == Decimal("3642.77")

    # 3. Validar registros Agrupados
    # Adalfi Bido (2 registros, Servicios)
    adalfi_recs = [r for r in rec_agrupados if "adalfi" in r.tx.proveedor_norm]
    assert len(adalfi_recs) == 2
    # Adalfi tiene "Servicio" en la nota, que es No Profesional según las keywords
    assert all(r.tipo_gasto in ("Servicio Profesional", "Servicio No Profesional") for r in adalfi_recs)
    assert sum(r.tx.importe_sujeto for r in adalfi_recs) == Decimal("2500.00")

    # Adelaida Sarante (1 registro, Sin clasificar, RNC de maestro)
    adelaida_recs = [r for r in rec_agrupados if "adelaida" in r.tx.proveedor_norm]
    assert len(adelaida_recs) == 1
    assert adelaida_recs[0].rnc_cedula == "13400049261"
    assert adelaida_recs[0].tipo_gasto == "Sin clasificar"
    assert adelaida_recs[0].tx.importe_sujeto == Decimal("15000.00")

    # Domingo Ebanista (2 registros, Sin clasificar, prevalece maestro sobre nota)
    domingo_recs = [r for r in rec_agrupados if "domingo" in r.tx.proveedor_norm]
    assert len(domingo_recs) == 2
    assert all(r.rnc_cedula == "06600054594" for r in domingo_recs)
    assert all(r.tipo_gasto == "Sin clasificar" for r in domingo_recs)
    assert sum(r.tx.importe_sujeto for r in domingo_recs) == Decimal("32570.00")

    # 4. Validar Gastos menores B13
    assert len(rec_b13) == 1
    r_dj = rec_b13[0]
    assert "dj ludo" in r_dj.tx.proveedor_norm
    assert r_dj.tx.importe_sujeto == Decimal("8000.00")
    assert "DJ Ludo" in r_dj.motivo_b13

    # 5. Validar incidencias esperadas
    tipos_incidencias = [i.tipo for i in incidences]
    assert "CONFLICTO_IDENTIFICACION" in tipos_incidencias
    assert "NOTA_SIN_CLASIFICAR" in tipos_incidencias
    assert "PROVEEDOR_MULTI_ENTIDAD" in tipos_incidencias
