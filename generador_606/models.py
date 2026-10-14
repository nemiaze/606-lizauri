"""
Modelos de datos para el generador de formato 606 e incidencias.
"""

from dataclasses import dataclass, field
from decimal import Decimal
from typing import List, Optional, Tuple


@dataclass
class Transaction:
    fila: int
    fecha_raw: str
    fecha_parsed: Optional[Tuple[int, int, int]]  # (YYYY, MM, DD)
    tipo_transaccion: str
    numero: str
    proveedor: str
    proveedor_norm: str
    producto_servicio: str
    nota: str
    importe: Decimal
    impuesto: Decimal
    departamento: str
    importe_sujeto: Decimal
    cp_pagadas: str


@dataclass
class ProviderMaster:
    proveedor: str
    proveedor_norm: str
    rnc_cedula_raw: str
    rnc_cedula_clean: Optional[str]  # 9 u 11 dígitos
    tipo_id: Optional[str]           # '1' (RNC, 9 dígitos) o '2' (Cédula, 11 dígitos)
    razon_social: str
    tipo_bienes_servicios: Optional[str]
    is_valid_id: bool
    validation_error: Optional[str] = None


@dataclass
class Incidence:
    tipo: str
    fila: Optional[int]
    proveedor: Optional[str]
    descripcion: str
    severidad: str = "ADVERTENCIA"  # "ERROR", "ADVERTENCIA", "INFO"

    def __str__(self) -> str:
        parts = []
        if self.fila is not None:
            parts.append(f"Fila {self.fila}")
        if self.proveedor:
            parts.append(f"[{self.proveedor}]")
        parts.append(f"({self.severidad}) {self.descripcion}")
        return " - ".join(parts)


@dataclass
class ProcessedRecord:
    tx: Transaction
    is_606: bool
    ncf_valido: Optional[str] = None
    rnc_cedula: Optional[str] = None
    tipo_id: Optional[str] = None
    fuente_rnc: str = "N/A"          # "Maestro", "Nota", "N/A"
    tipo_gasto: str = "Sin clasificar" # "Servicio", "Bien", "Sin clasificar"
    destino: str = "Agrupado"        # "606", "Agrupado", "B13 Gastos menores"
    motivo_b13: Optional[str] = None
    observaciones: List[str] = field(default_factory=list)
    itbis_exento: bool = False         # True si el producto/servicio está exento de ITBIS (DGII)
