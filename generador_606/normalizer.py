"""
Funciones de normalización, validación y extracción de datos.
"""

from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
import re
from typing import Any, List, Optional, Tuple
import unicodedata
import pandas as pd


# Regex para NCF tradicionales y electrónicos (e-CF)
REGEX_NCF_TRADICIONAL = re.compile(r"^B(01|02|03|04|11|12|13|14|15|16|17)\d{8}$")
REGEX_NCF_ELECTRONICO = re.compile(r"^E(31|32|33|34|41|43|44|45|46|47)\d{10}$")

# Regex para extraer cédulas y RNCs de la nota
REGEX_NOTA_CEDULA = re.compile(r"(?<!\d)(\d{3}-?\d{7}-?\d)(?!\d)")
REGEX_NOTA_RNC_GUIONES = re.compile(r"(?<!\d)(\d-\d{2}-\d{5}-\d)(?!\d)")
REGEX_NOTA_RNC_PLANO = re.compile(r"(?<!\d)(\d{9})(?!\d)")


def normalize_text(text: Any) -> str:
    """
    Normaliza texto: quita espacios en extremos, colapsa espacios múltiples,
    elimina acentos y convierte a minúsculas (casefold).
    """
    if text is None or pd.isna(text):
        return ""
    s = str(text)
    # Quitar acentos
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")
    # Colapsar espacios y recortar
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()


def parse_date(val: Any) -> Optional[Tuple[int, int, int]]:
    """
    Parsea una fecha proveniente de Excel (sea datetime real o texto).
    Retorna una tupla (año, mes, día) o None si no es válida.
    """
    if val is None or pd.isna(val):
        return None
    
    # Si ya es un objeto datetime o pd.Timestamp
    if isinstance(val, (datetime, date, pd.Timestamp)):
        return val.year, val.month, val.day

    s = str(val).strip()
    if not s:
        return None

    # Formatos esperados: MM/DD/YYYY, YYYY-MM-DD, etc.
    formatos = [
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m-%d-%y",
        "%m-%d-%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
    ]
    for fmt in formatos:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.year, dt.month, dt.day
        except ValueError:
            pass

    # Fallback con pandas
    try:
        dt = pd.to_datetime(s)
        if not pd.isna(dt):
            return dt.year, dt.month, dt.day
    except Exception:
        pass

    return None


def to_decimal(val: Any, default: str = "0.00") -> Decimal:
    """
    Convierte cualquier valor a Decimal redondeado a 2 decimales (ROUND_HALF_UP).
    Limpia símbolos de moneda, comas, espacios y maneja números negativos en paréntesis.
    """
    if val is None or pd.isna(val):
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    s = str(val).strip()
    if not s:
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()

    # Quitar todo excepto dígitos y punto decimal
    cleaned = re.sub(r"[^\d.]", "", s)
    if not cleaned:
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    try:
        d = Decimal(cleaned)
        if is_negative:
            d = -d
        return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal(default).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def validate_and_clean_rnc_cedula(raw_val: Any) -> Tuple[bool, Optional[str], Optional[str], Optional[str]]:
    """
    Valida y limpia un RNC o Cédula según la sección 2.2 de la especificación:
    - Solo válido si al quitar espacios y guiones queda SOLO dígitos y longitud 9 u 11.
    - No se extraen dígitos de textos como 'Pasaporte 179054679'.
    
    Retorna: (is_valid, clean_id, tipo_id, error_message)
    tipo_id: '1' para RNC (9 dígitos), '2' para Cédula (11 dígitos).
    """
    if raw_val is None or pd.isna(raw_val):
        return False, None, None, "Valor vacío o no proporcionado"

    s = str(raw_val).strip()
    if not s:
        return False, None, None, "Valor vacío"

    # Si contiene letras u otros caracteres no permitidos (ej. 'Pasaporte 179054679')
    if not re.fullmatch(r"[\d\s\-]+", s):
        return False, None, None, f"Contiene caracteres alfabéticos o no numéricos ('{s}')"

    digits = re.sub(r"\D", "", s)
    length = len(digits)

    if length == 9:
        return True, digits, "1", None
    elif length == 11:
        return True, digits, "2", None
    else:
        return False, None, None, f"Longitud inválida ({length} dígitos, se esperan 9 u 11)"


def verify_cedula_jce(cedula_digits: str) -> Tuple[bool, int, int]:
    """
    Verifica el dígito verificador de una cédula dominicana de 11 dígitos
    usando el algoritmo de la JCE (pesos 1, 2, 1, 2...):
    - Pesos: 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 sobre los primeros 10 dígitos.
    - Si el producto >= 10, se suman sus dígitos (ej. 14 -> 1 + 4 = 5).
    - Verificador = (10 - (suma % 10)) % 10.
    
    Retorna: (is_valid, expected_verifier, actual_verifier)
    """
    if len(cedula_digits) != 11 or not cedula_digits.isdigit():
        return False, -1, -1

    weights = [1, 2, 1, 2, 1, 2, 1, 2, 1, 2]
    total = 0
    for d_char, w in zip(cedula_digits[:10], weights):
        prod = int(d_char) * w
        if prod >= 10:
            prod = (prod // 10) + (prod % 10)
        total += prod

    expected_verifier = (10 - (total % 10)) % 10
    actual_verifier = int(cedula_digits[10])
    return (expected_verifier == actual_verifier), expected_verifier, actual_verifier


def is_valid_ncf(numero_val: Any) -> Tuple[bool, Optional[str]]:
    """
    Verifica si un número de comprobante es un NCF válido (B... de 11 caracteres o E... de 13).
    Retorna (is_valid, ncf_limpio).
    """
    if numero_val is None or pd.isna(numero_val):
        return False, None

    s = str(numero_val).strip().upper()
    if not s:
        return False, None

    if REGEX_NCF_TRADICIONAL.match(s) or REGEX_NCF_ELECTRONICO.match(s):
        return True, s

    return False, None


def extract_rnc_cedula_from_note(note_val: Any) -> Tuple[Optional[str], Optional[str]]:
    r"""
    Extrae RNC o Cédula desde la columna Nota según la sección 5.2.
    - Cédula: (?<!\d)(\d{3}-?\d{7}-?\d)(?!\d)
    - RNC: con guiones (\d-\d{2}-\d{5}-\d) o plano de 9 dígitos.
    - Pasaporte alfanumérico no cuenta.
    
    Retorna (clean_id, tipo_id) o (None, None).
    """
    if note_val is None or pd.isna(note_val):
        return None, None

    text = str(note_val)

    # 1. Buscar Cédula
    match_ced = REGEX_NOTA_CEDULA.search(text)
    if match_ced:
        clean_ced = re.sub(r"\D", "", match_ced.group(1))
        if len(clean_ced) == 11:
            return clean_ced, "2"

    # 2. Buscar RNC con formato
    match_rnc_fmt = REGEX_NOTA_RNC_GUIONES.search(text)
    if match_rnc_fmt:
        clean_rnc = re.sub(r"\D", "", match_rnc_fmt.group(1))
        if len(clean_rnc) == 9:
            return clean_rnc, "1"

    # 3. Buscar RNC plano
    match_rnc_plano = REGEX_NOTA_RNC_PLANO.search(text)
    if match_rnc_plano:
        clean_rnc = match_rnc_plano.group(1)
        if len(clean_rnc) == 9:
            return clean_rnc, "1"

    return None, None


def classify_bien_servicio(
    note_val: Any,
    keywords_servicio_profesional: List[str],
    keywords_servicio_no_profesional: List[str],
    keywords_bien: List[str],
    keywords_servicio: Optional[List[str]] = None,
) -> str:
    """
    Clasifica un registro en 'Servicio Profesional', 'Servicio No Profesional',
    'Bien' o 'Sin clasificar' según la columna Nota.

    Tasas de retención ISR aplicables (DGII):
      - Servicio Profesional     → 15%
      - Servicio No Profesional  →  3%
      - Bien / Sin clasificar    →  0%

    Prioridad de búsqueda: Profesional > No Profesional > Bien.
    """
    if note_val is None or pd.isna(note_val):
        return "Sin clasificar"

    norm_note = normalize_text(note_val)
    if not norm_note:
        return "Sin clasificar"

    # 1. Servicio Profesional (mayor retención → prioridad alta)
    for kw in keywords_servicio_profesional:
        norm_kw = normalize_text(kw)
        if norm_kw and norm_kw in norm_note:
            return "Servicio Profesional"

    # 2. Servicio No Profesional
    for kw in keywords_servicio_no_profesional:
        norm_kw = normalize_text(kw)
        if norm_kw and norm_kw in norm_note:
            return "Servicio No Profesional"

    # 3. Bien
    for kw in keywords_bien:
        norm_kw = normalize_text(kw)
        if norm_kw and norm_kw in norm_note:
            return "Bien"

    return "Sin clasificar"


def is_itbis_exento(note_val: Any, keywords_exento: List[str]) -> bool:
    """
    Determina si un producto/servicio está exento de ITBIS según la DGII,
    buscando palabras clave de exención en la columna Nota.

    Bienes exentos (Art. 343 CNDR): animales vivos, carnes frescas/refrigeradas/
    congeladas, pescado de consumo popular, lácteos (excl. yogurt y mantequilla),
    leche, miel, plantas para siembra, legumbres/hortalizas/tubérculos sin
    procesar, frutas sin procesar, café sin tostar, cereales/harinas/granos,
    productos de molinería, semillas oleaginosas, embutidos, cacao, leña, etc.

    Servicios exentos: financieros/seguros, pensiones/jubilaciones, transporte
    terrestre, electricidad/agua/basura, alquiler de viviendas, salud,
    educativos/culturales, funerarios, salones de belleza/peluquerías.

    Retorna True si el producto/servicio es exento de ITBIS, False si no.
    """
    if note_val is None or pd.isna(note_val):
        return False

    norm_note = normalize_text(note_val)
    if not norm_note:
        return False

    for kw in keywords_exento:
        norm_kw = normalize_text(kw)
        if norm_kw and norm_kw in norm_note:
            return True

    return False
