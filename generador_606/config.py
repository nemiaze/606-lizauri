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
        "archivo_606": "formato_606.xlsx",
        "archivo_606_csv": "formato_606.csv",
        "archivo_agrupacion": "agrupacion_sin_comprobante.xlsx",
        "archivo_incidencias": "incidencias.txt",
    },
    "fiscal": {
        "tasa_retencion_profesional": 0.15,
        "tasa_retencion_no_profesional": 0.03,
        "tipo_bienes_servicios_606": "09 - COMPRAS Y GASTOS QUE FORMARAN PARTE DEL COSTO DE VENTA",
        "tipo_retencion_isr": "02",
        "bienes_vs_servicios_606": "Bienes",
        "forma_pago_defecto_credito": "04 - COMPRA A CRÉDITO",
        "forma_pago_defecto_pagado": "02 - CHEQUES/TRANSFERENCIAS/DEPÓSITO",
    },
    "clasificacion_nota": {
        "keywords_servicio_profesional": [
            "honorario", "honorarios", "asesoria", "asesoría",
            "consultoria", "consultoría", "profesional", "abogado",
            "contador", "ingeniero", "arquitecto", "medico", "médico",
            "notario", "auditoria", "auditoría"
        ],
        "keywords_servicio_no_profesional": [
            "servicio", "servicios", "mano de obra", "reparacion",
            "reparación", "mantenimiento", "alquiler", "arrendamiento",
            "transporte", "flete", "instalacion", "instalación",
            "confeccion", "confección", "limpieza", "seguridad"
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
    "reglas_proveedor": {
        "proveedores_bien": ["pescaderia chepe", "javier valerio"],
        "proveedores_servicio_no_profesional": [
            "faris polanco", "antonio polanco", "jose miguel brito", "brito batista"
        ],
        "proveedores_servicio_profesional": [],
        "proveedores_exento_itbis": ["pescaderia chepe"],
    },
    "itbis_exento": {
        "keywords_exento": [
            # Bienes exentos
            "animal vivo", "animales vivos",
            "carne fresca", "carne refrigerada", "carne congelada", "carne",
            "pescado", "pesca",
            "lacteo", "lácteo", "leche", "queso", "yogurt",
            "miel",
            "huevo", "huevos",
            "planta para siembra", "plantula", "plántula",
            "legumbre", "hortaliza", "tuberculo", "tubérculo",
            "fruta", "frutas",
            "cafe", "café",
            "cereal", "harina", "grano",
            "molineria", "molinería",
            "semilla oleaginosa", "semilla",
            "embutido", "embutidos",
            "cacao",
            # Servicios exentos
            "servicio financiero", "seguro", "seguros",
            "pension", "pensión", "jubilacion", "jubilación",
            "transporte terrestre", "transporte de carga",
            "electricidad", "energia electrica", "energía eléctrica",
            "agua potable", "acueducto",
            "alquiler vivienda", "renta vivienda",
            "salud", "medico", "médico", "hospital", "clinica", "clínica",
            "educacion", "educación", "escuela", "colegio", "universidad",
            "cultural", "cultura",
            "funerario", "funeraria", "funeral",
            "salon de belleza", "salón de belleza", "peluqueria", "peluquería",
        ]
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
    archivo_606: str = "formato_606.xlsx"
    archivo_606_csv: str = "formato_606.csv"
    archivo_agrupacion: str = "agrupacion_sin_comprobante.xlsx"
    archivo_incidencias: str = "incidencias.txt"
    tasa_retencion_profesional: Decimal = Decimal("0.15")
    tasa_retencion_no_profesional: Decimal = Decimal("0.03")
    tipo_bienes_servicios_606: str = "09 - COMPRAS Y GASTOS QUE FORMARAN PARTE DEL COSTO DE VENTA"
    tipo_retencion_isr: str = "02"
    bienes_vs_servicios_606: str = "Bienes"
    forma_pago_credito: str = "04 - COMPRA A CRÉDITO"
    forma_pago_pagado: str = "02 - CHEQUES/TRANSFERENCIAS/DEPÓSITO"
    keywords_servicio: List[str] = field(default_factory=list)
    keywords_servicio_profesional: List[str] = field(default_factory=list)
    keywords_servicio_no_profesional: List[str] = field(default_factory=list)
    keywords_bien: List[str] = field(default_factory=list)
    keywords_itbis_exento: List[str] = field(default_factory=list)
    proveedores_b13: List[str] = field(default_factory=list)
    proveedores_bien: List[str] = field(default_factory=list)
    proveedores_servicio_no_profesional: List[str] = field(default_factory=list)
    proveedores_servicio_profesional: List[str] = field(default_factory=list)
    proveedores_exento_itbis: List[str] = field(default_factory=list)
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
        tasa_retencion_profesional: Optional[float] = None,
        tasa_retencion_no_profesional: Optional[float] = None,
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

        # Resolve tasas de retención diferenciadas
        if tasa_retencion_profesional is not None:
            tasa_prof_dec = Decimal(str(tasa_retencion_profesional))
        else:
            tasa_prof_dec = Decimal(str(fiscal.get("tasa_retencion_profesional", 0.15)))

        if tasa_retencion_no_profesional is not None:
            tasa_no_prof_dec = Decimal(str(tasa_retencion_no_profesional))
        else:
            tasa_no_prof_dec = Decimal(str(fiscal.get("tasa_retencion_no_profesional", 0.03)))

        return cls(
            transacciones_path=t_path,
            proveedores_path=p_path,
            salida_dir=out_dir,
            hoja_proveedores=archivos.get("hoja_proveedores", "Hoja1"),
            archivo_606=archivos.get("archivo_606", "formato_606.xlsx"),
            archivo_606_csv=archivos.get("archivo_606_csv", "formato_606.csv"),
            archivo_agrupacion=archivos.get("archivo_agrupacion", "agrupacion_sin_comprobante.xlsx"),
            archivo_incidencias=archivos.get("archivo_incidencias", "incidencias.txt"),
            tasa_retencion_profesional=tasa_prof_dec,
            tasa_retencion_no_profesional=tasa_no_prof_dec,
            tipo_bienes_servicios_606=str(fiscal.get("tipo_bienes_servicios_606", "09 - COMPRAS Y GASTOS QUE FORMARAN PARTE DEL COSTO DE VENTA")),
            tipo_retencion_isr=str(fiscal.get("tipo_retencion_isr", "02")),
            bienes_vs_servicios_606=str(fiscal.get("bienes_vs_servicios_606", "Bienes")),
            forma_pago_credito=str(fiscal.get("forma_pago_defecto_credito", "04 - COMPRA A CRÉDITO")),
            forma_pago_pagado=str(fiscal.get("forma_pago_defecto_pagado", "02 - CHEQUES/TRANSFERENCIAS/DEPÓSITO")),
            keywords_servicio=clasificacion.get("keywords_servicio_profesional", []) + clasificacion.get("keywords_servicio_no_profesional", []),
            keywords_servicio_profesional=clasificacion.get("keywords_servicio_profesional", []),
            keywords_servicio_no_profesional=clasificacion.get("keywords_servicio_no_profesional", []),
            keywords_bien=clasificacion.get("keywords_bien", []),
            keywords_itbis_exento=raw_cfg.get("itbis_exento", {}).get("keywords_exento", []),
            proveedores_b13=excepciones.get("proveedores_b13", ["dj ludo"]),
            proveedores_bien=[str(x).strip().lower() for x in raw_cfg.get("reglas_proveedor", {}).get("proveedores_bien", ["pescaderia chepe", "javier valerio"])],
            proveedores_servicio_no_profesional=[str(x).strip().lower() for x in raw_cfg.get("reglas_proveedor", {}).get("proveedores_servicio_no_profesional", ["faris polanco", "antonio polanco", "jose miguel brito", "brito batista"])],
            proveedores_servicio_profesional=[str(x).strip().lower() for x in raw_cfg.get("reglas_proveedor", {}).get("proveedores_servicio_profesional", [])],
            proveedores_exento_itbis=[str(x).strip().lower() for x in raw_cfg.get("reglas_proveedor", {}).get("proveedores_exento_itbis", ["pescaderia chepe"])],
            periodo=str(periodo).strip() if periodo else None,
            columnas_transacciones=raw_cfg.get("columnas_transacciones", {}),
            columnas_proveedores=raw_cfg.get("columnas_proveedores", {}),
        )
