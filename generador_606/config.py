"""
Módulo de configuración para el generador 606 y agrupación de comprobantes.
"""

from dataclasses import dataclass, field
from decimal import Decimal
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import yaml


DEFAULT_CONFIG: Dict[str, Any] = {
    "archivos": {
        "transacciones_defecto": "muestra.xlsx",
        "proveedores_defecto": "Proveedores.xlsx",
        "hoja_proveedores": "Hoja1",
        "salida_defecto": "./salida",
        "archivo_606": "formato_606.csv",
        "archivo_agrupacion": "agrupacion_sin_comprobante.xlsx",
        "archivo_incidencias": "incidencias.txt",
    },
    "fiscal": {
        "tasa_retencion_renta": 0.10,
        "tipo_bienes_servicios_606": "09",
        "tipo_retencion_isr": "02",
        "bienes_vs_servicios_606": "Bienes",
        "forma_pago_defecto_credito": "04",
        "forma_pago_defecto_pagado": "02",
    },
    "clasificacion_nota": {
        "keywords_servicio": [
            "servicio", "servicios", "honorario", "honorarios",
            "mano de obra", "reparacion", "reparación", "mantenimiento",
            "alquiler", "arrendamiento", "transporte", "flete",
            "asesoria", "asesoría", "consultoria", "consultoría",
            "instalacion", "instalación", "confeccion", "confección"
        ],
        "keywords_bien": [
            "bien", "bienes", "mercancia", "mercancía", "material",
            "materiales", "compra", "suministro", "suministros",
            "pieza", "piezas", "repuesto", "repuestos", "equipo",
            "equipos", "herramienta", "herramientas"
        ],
    },
    "excepciones_gastos_menores": {
        "proveedores_b13": ["dj ludo"],
    },
    "columnas_transacciones": {
        "fecha": "Fecha",
        "tipo_transaccion": "Tipo de transacción",
        "numero": "Número",
        "proveedor": "Proveedor",
        "producto_servicio": "Producto/servicio",
        "nota": "Nota",
        "importe": "Importe",
        "impuesto": "Impuesto",
        "departamento": "$$Departamento$$",
        "importe_sujeto_impuestos": "Importe sujeto a impuestos",
        "cp_pagadas": "C/P pagadas",
    },
    "columnas_proveedores": {
        "proveedor": "Proveedor",
        "rnc_cedula": "RNC/CEDULA",
        "razon_social": "RAZON SOCIAL",
        "tipo_bienes_servicios": "Tipo Bienes y Servicios Comprados",
    }
}


@dataclass
class AppConfig:
    transacciones_path: Path
    proveedores_path: Path
    salida_dir: Path
    hoja_proveedores: str = "Hoja1"
    archivo_606: str = "formato_606.csv"
    archivo_agrupacion: str = "agrupacion_sin_comprobante.xlsx"
    archivo_incidencias: str = "incidencias.txt"
    tasa_retencion: Decimal = Decimal("0.10")
    tipo_bienes_servicios_606: str = "09"
    tipo_retencion_isr: str = "02"
    bienes_vs_servicios_606: str = "Bienes"
    forma_pago_credito: str = "04"
    forma_pago_pagado: str = "02"
    keywords_servicio: List[str] = field(default_factory=list)
    keywords_bien: List[str] = field(default_factory=list)
    proveedores_b13: List[str] = field(default_factory=list)
    periodo: Optional[str] = None
    columnas_transacciones: Dict[str, str] = field(default_factory=dict)
    columnas_proveedores: Dict[str, str] = field(default_factory=dict)

    @classmethod
    def load(
        cls,
        config_path: Optional[str] = "config.yaml",
        transacciones_path: Optional[str] = None,
        proveedores_path: Optional[str] = None,
        salida_dir: Optional[str] = None,
        tasa_retencion: Optional[float] = None,
        periodo: Optional[str] = None,
    ) -> "AppConfig":
        raw_cfg = dict(DEFAULT_CONFIG)

        # Load YAML if present
        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    user_cfg = yaml.safe_load(f)
                    if user_cfg and isinstance(user_cfg, dict):
                        for section, val in user_cfg.items():
                            if isinstance(val, dict) and section in raw_cfg and isinstance(raw_cfg[section], dict):
                                raw_cfg[section].update(val)
                            else:
                                raw_cfg[section] = val
            except Exception as e:
                print(f"[ADVERTENCIA] No se pudo leer {config_path}: {e}. Se usarán valores predeterminados.")

        archivos = raw_cfg.get("archivos", {})
        fiscal = raw_cfg.get("fiscal", {})
        clasificacion = raw_cfg.get("clasificacion_nota", {})
        excepciones = raw_cfg.get("excepciones_gastos_menores", {})

        # Resolve paths
        t_path = Path(transacciones_path or archivos.get("transacciones_defecto", "muestra.xlsx"))
        p_path = Path(proveedores_path or archivos.get("proveedores_defecto", "Proveedores.xlsx"))
        out_dir = Path(salida_dir or archivos.get("salida_defecto", "./salida"))

        # Resolve tasa de retención
        if tasa_retencion is not None:
            tasa_dec = Decimal(str(tasa_retencion))
        else:
            tasa_dec = Decimal(str(fiscal.get("tasa_retencion_renta", 0.10)))

        return cls(
            transacciones_path=t_path,
            proveedores_path=p_path,
            salida_dir=out_dir,
            hoja_proveedores=archivos.get("hoja_proveedores", "Hoja1"),
            archivo_606=archivos.get("archivo_606", "formato_606.csv"),
            archivo_agrupacion=archivos.get("archivo_agrupacion", "agrupacion_sin_comprobante.xlsx"),
            archivo_incidencias=archivos.get("archivo_incidencias", "incidencias.txt"),
            tasa_retencion=tasa_dec,
            tipo_bienes_servicios_606=str(fiscal.get("tipo_bienes_servicios_606", "09")),
            tipo_retencion_isr=str(fiscal.get("tipo_retencion_isr", "02")),
            bienes_vs_servicios_606=str(fiscal.get("bienes_vs_servicios_606", "Bienes")),
            forma_pago_credito=str(fiscal.get("forma_pago_defecto_credito", "04")),
            forma_pago_pagado=str(fiscal.get("forma_pago_defecto_pagado", "02")),
            keywords_servicio=clasificacion.get("keywords_servicio", []),
            keywords_bien=clasificacion.get("keywords_bien", []),
            proveedores_b13=excepciones.get("proveedores_b13", ["dj ludo"]),
            periodo=str(periodo).strip() if periodo else None,
            columnas_transacciones=raw_cfg.get("columnas_transacciones", {}),
            columnas_proveedores=raw_cfg.get("columnas_proveedores", {}),
        )
