"""
Pruebas unitarias de normalización, validación y extracción de datos.
"""

from datetime import datetime
from decimal import Decimal
import pytest

from generador_606.normalizer import (
    classify_bien_servicio,
    extract_rnc_cedula_from_note,
    is_valid_ncf,
    normalize_text,
    parse_date,
    to_decimal,
    validate_and_clean_rnc_cedula,
    verify_cedula_jce,
)


def test_normalize_text():
    assert normalize_text("  Agropin (Pablo Beato Espinal)  ") == "agropin (pablo beato espinal)"
    assert normalize_text("DJ LUDO\nYhouberth") == "dj ludo yhouberth"
    assert normalize_text("Reparación y Mantenimiento") == "reparacion y mantenimiento"
    assert normalize_text(None) == ""


def test_parse_date():
    # Formato MM/DD/YYYY
    assert parse_date("08/13/2026") == (2026, 8, 13)
    assert parse_date("01/30/2026") == (2026, 1, 30)

    # Objeto datetime real
    dt = datetime(2026, 5, 8, 0, 0, 0)
    assert parse_date(dt) == (2026, 5, 8)

    # Formato ISO YYYY-MM-DD
    assert parse_date("2026-08-08") == (2026, 8, 8)
    assert parse_date("2026-08-08 14:30:00") == (2026, 8, 8)

    # Inválidos
    assert parse_date("") is None
    assert parse_date("fecha_invalida") is None


def test_to_decimal():
    assert to_decimal("2400") == Decimal("2400.00")
    assert to_decimal(366.1) == Decimal("366.10")
    assert to_decimal("-2033.90") == Decimal("-2033.90")
    assert to_decimal("RD$ 1,500.50") == Decimal("1500.50")
    assert to_decimal(None) == Decimal("0.00")


def test_validate_and_clean_rnc_cedula():
    # RNC válido (9 dígitos)
    ok, cid, tid, err = validate_and_clean_rnc_cedula("132465555")
    assert ok is True and cid == "132465555" and tid == "1"

    # Cédula válida (11 dígitos con y sin guiones)
    ok, cid, tid, err = validate_and_clean_rnc_cedula("402-3581367-8")
    assert ok is True and cid == "40235813678" and tid == "2"

    ok, cid, tid, err = validate_and_clean_rnc_cedula("06600054594")
    assert ok is True and cid == "06600054594" and tid == "2"

    # Pasaporte con texto (NO debe extraer dígitos)
    ok, cid, tid, err = validate_and_clean_rnc_cedula("Pasaporte 179054679")
    assert ok is False
    assert cid is None

    # Longitud inválida
    ok, cid, tid, err = validate_and_clean_rnc_cedula("1234567")
    assert ok is False

    # Vacío
    ok, cid, tid, err = validate_and_clean_rnc_cedula("")
    assert ok is False


def test_verify_cedula_jce():
    # Cédulas dominicanas reales del caso de prueba
    valid_cedulas = [
        "40235813678",
        "13400049261",
        "06600056540",
        "06600054594",
        "06600050030",
    ]
    for c in valid_cedulas:
        valid, exp, act = verify_cedula_jce(c)
        assert valid is True, f"Cédula {c} falló: esperado={exp}, actual={act}"

    # Cédula con dígito inválido
    valid, exp, act = verify_cedula_jce("40235813679")
    assert valid is False


def test_is_valid_ncf():
    # NCF tradicionales válidos (11 caracteres: B + 2 tipo + 8 secuencia)
    assert is_valid_ncf("B0100000188") == (True, "B0100000188")
    assert is_valid_ncf("B0200000001") == (True, "B0200000001")
    assert is_valid_ncf("b1300000055") == (True, "B1300000055")

    # e-CF válidos (13 caracteres: E + 2 tipo + 10 secuencia)
    assert is_valid_ncf("E310000000001") == (True, "E310000000001")
    assert is_valid_ncf("E450000000099") == (True, "E450000000099")

    # Comprobantes inválidos o sin comprobante
    assert is_valid_ncf("Ene30")[0] is False
    assert is_valid_ncf("00003")[0] is False
    assert is_valid_ncf("")[0] is False
    assert is_valid_ncf(None)[0] is False
    assert is_valid_ncf("B9900000001")[0] is False  # Tipo inválido


def test_extract_rnc_cedula_from_note():
    # Cédula con guiones
    cid, tid = extract_rnc_cedula_from_note("Adalfi Bido Garcia\n402-3581367-8\nServicio")
    assert cid == "40235813678" and tid == "2"

    cid, tid = extract_rnc_cedula_from_note("Diomedes de los Santos 066-0005003-0 (concho Yiyo)")
    assert cid == "06600050030" and tid == "2"

    # Pasaporte alfanumérico (NO debe extraerse)
    cid, tid = extract_rnc_cedula_from_note("Yhouberth Roselv Ludovic Rodriguez\n93LY5366192\nDJ LUDO")
    assert cid is None and tid is None


def test_classify_bien_servicio():
    kw_prof = ["honorario", "consultoria", "profesional"]
    kw_no_prof = ["servicio", "servicios", "mantenimiento"]
    kw_bien = ["bien", "bienes", "material", "compra"]

    assert classify_bien_servicio("Honorarios de abogado", kw_prof, kw_no_prof, kw_bien) == "Servicio Profesional"
    assert classify_bien_servicio("Servicio de cocina", kw_prof, kw_no_prof, kw_bien) == "Servicio No Profesional"
    assert classify_bien_servicio("Servicio de consultoría", kw_prof, kw_no_prof, kw_bien) == "Servicio Profesional"
    assert classify_bien_servicio("Compra de bienes de oficina", kw_prof, kw_no_prof, kw_bien) == "Bien"
    assert classify_bien_servicio("Diomedes de los Santos 066-0005003-0", kw_prof, kw_no_prof, kw_bien) == "Sin clasificar"
    assert classify_bien_servicio(None, kw_prof, kw_no_prof, kw_bien) == "Sin clasificar"
