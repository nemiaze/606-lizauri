# Generador del Formato 606 (DGII) y Agrupación de Compras sin Comprobante

Programa profesional en Python para la lectura, validación fiscal, separación y generación de reportes tributarios ante la Dirección General de Impuestos Internos (DGII) de la República Dominicana.

---

## 🚀 Formas de Ejecución (Sencillo de Usar)

El sistema ha sido diseñado para que cualquier usuario, técnico o contable, pueda utilizarlo de forma intuitiva:

### Opción 1: Doble Clic en Windows (Recomendada para usuarios)
Simplemente haz doble clic sobre el archivo:
```text
ejecutar_606.bat
```
Se abrirá un menú interactivo que por defecto inicia la **Interfaz Gráfica (Ventana)**.

---

### Opción 2: Interfaz Gráfica de Usuario (GUI)
Puedes abrir directamente la ventana ejecutando:
```bash
python app_gui.py
```
o también:
```bash
python generar_606.py --gui
```

**Ventajas de la Interfaz:**
- Botones **Examinar...** para seleccionar el archivo de transacciones, maestro de proveedores y carpeta destino.
- Campo para modificar la tasa de retención de ISR (por defecto 10%).
- Campo opcional para auditar período fiscal (`AAAAMM`, ej. `202608`).
- Botón **PROCESAR Y GENERAR ARCHIVOS**.
- Visor en tiempo real con código de colores para éxitos, errores y advertencias.
- Botón directo **Abrir Carpeta de Salida** para abrir los archivos generados en el Explorador de Windows.

---

### Opción 3: Línea de Comandos (CLI / Automatización)
Puedes ejecutarlo directamente desde CMD o PowerShell:

```bash
# Ejecución estándar con valores por defecto
python generar_606.py

# Especificando archivos y parámetros personalizados
python generar_606.py --transacciones muestra.xlsx --proveedores Proveedores.xlsx --salida ./salida --tasa-retencion 0.10

# Filtrando o auditando un período específico
python generar_606.py --transacciones grande.xlsx --periodo 202608
```

#### Parámetros disponibles en CLI:
| Parámetro | Abreviatura | Descripción | Valor por defecto |
|---|---|---|---|
| `--transacciones` | `-t` | Ruta al archivo Excel de transacciones | `muestra.xlsx` |
| `--proveedores` | `-p` | Ruta al maestro de proveedores | `Proveedores.xlsx` |
| `--salida` | `-s` | Carpeta de salida para los archivos generados | `./salida` |
| `--tasa-retencion`| `-r` | Tasa de retención sobre servicios (ej. `0.10` para 10%) | `0.10` |
| `--periodo` | | Período fiscal a auditar (`AAAAMM`) | Opcional |
| `--config` | `-c` | Archivo de configuración YAML personalizado | `config.yaml` |
| `--gui` | | Inicia la interfaz gráfica de usuario | Desactivado |

---

## 📁 Archivos de Salida Generados

El programa deposita en la carpeta de salida (por defecto `./salida/`) tres archivos:

### 1. `formato_606.xlsx` (y formato alternativo `formato_606.csv`)
- **`formato_606.xlsx` (Excel nativo):** Viene directamente dividido en sus **25 columnas oficiales**, estilizado con encabezado azul oscuro, importes numéricos con formato de moneda y texto para RNCs/fechas. **No requiere que el usuario haga conversión de texto a columnas ni configure delimitadores en Excel**. Incluye fila de totales de control.
- **`formato_606.csv` (CSV oficial DGII):** Generado simultáneamente en codificación **UTF-8 con BOM** (`utf-8-sig`) con las 25 columnas separadas por comas, listo para subir directamente a las plataformas web o herramientas de validación de la DGII.

### 2. `agrupacion_sin_comprobante.xlsx`
Libro de Excel estilizado con tipografía **Arial**, encabezados oscuros, paneles inmovilizados y **fórmulas vivas de Excel** (`fullCalcOnLoad = True`):

- **Hoja 1 (`Agrupado por proveedor`):**
  - Una fila por proveedor con RNC o cédula identificada.
  - **Celda B3:** Tasa de retención editable (fondo amarillo, texto azul y comentario explicativo).
  - Columnas de totales y retención calculadas mediante fórmulas:
    - Registros: `=COUNTIFS(Detalle!$B:$B, C6, Detalle!$F:$F, "Agrupado")`
    - Servicios: `=SUMIFS(Detalle!$H:$H, Detalle!$B:$B, C6, Detalle!$F:$F, "Agrupado", Detalle!$E:$E, "Servicio")`
    - Bienes: `=SUMIFS(...)`
    - Sin clasificar: `=SUMIFS(...)`
    - Total facturado: `=SUM(E6:G6)`
    - Retención renta: `=ROUND(E6 * $B$3, 2)` (recalcula al cambiar B3 o al reclasificar en Detalle).
    - Tipo Retención ISR: `=IF(E6>0, "02", "")`
  - Fila final de `TOTAL` con sumas dinámicas.

- **Hoja 2 (`Gastos menores B13`):**
  - Registros sin NCF que no tienen RNC/cédula ni en maestro ni en la nota.
  - Registros de la excepción **DJ Ludo** (clasificados con motivo `Excepción DJ Ludo`).
  - Fila de `TOTAL` con fórmulas de suma.

- **Hoja 3 (`Detalle`):**
  - Base de datos granular con todos los registros sin comprobante.
  - Saltos de línea en notas reemplazados por ` | `.
  - Columna `Destino` (`Agrupado` vs `B13 Gastos menores`) y columna `Tipo` (`Servicio`, `Bien`, `Sin clasificar`).

### 3. `incidencias.txt`
Informe de auditoría detallado con fecha, hora, conteos estadísticos e incidencias clasificadas por severidad:
- **ERRORES CRÍTICOS:** Proveedores no encontrados en maestro, descuadres de integridad (`Subtotal + ITBIS != Importe`), facturas con NCF pero sin RNC en maestro.
- **ADVERTENCIAS:** Conflictos entre RNC del maestro y cédula en la nota (prevalece el maestro), cédulas con dígito verificador JCE inválido, fechas fuera del período fiscal.
- **NOTAS INFORMATIVAS:** Registros sin clasificar (la nota no especifica si es bien o servicio), proveedores que facturan en múltiples entidades/departamentos.

---

## ⚙️ Configuración Personalizada (`config.yaml`)

El archivo [config.yaml](file:///c:/Users/Esdras/Desktop/Proyectos%20Neheli/606%20Lizauri/config.yaml) permite modificar sin tocar el código fuente:
- Tasa de retención por defecto.
- Códigos de tipo de gasto (por defecto `09` - costo de venta).
- Mapeo de columnas de transacciones y maestro.
- Palabras clave para clasificar servicios (`servicio`, `honorarios`, `mano de obra`, `alquiler`, etc.) y bienes (`bien`, `compra`, `suministro`, etc.).
- Lista de excepciones para Gastos Menores B13 (ej. `dj ludo`).

---

## 🧪 Pruebas Automatizadas

El proyecto cuenta con una batería de pruebas unitarias y de integración desarrolladas con `pytest` basadas en el caso de prueba oficial:

```bash
pytest -v
```

Las pruebas validan:
- Algoritmo JCE de dígito verificador de cédulas dominicanas.
- Regex y validación de NCF tradicionales y electrónicos.
- Extracción de identificadores desde notas y manejo de pasaportes.
- Manejo de fechas mixtas (`datetime` de Excel y strings `MM/DD/YYYY`).
- Conversión precisa de montos con `Decimal` (sin imprecisiones de coma flotante).
- Cuadre matemático exacto contra la Sección 7 de la especificación.

---

## 📦 Estructura del Proyecto

```text
606 Lizauri/
├── config.yaml                    <- Configuración externa (tasas, palabras clave, columnas)
├── generar_606.py                 <- Punto de entrada CLI
├── app_gui.py                     <- Interfaz gráfica de usuario (Tkinter)
├── ejecutar_606.bat               <- Lanzador doble clic para Windows
├── generador_606/                 <- Paquete modular principal
│   ├── __init__.py
│   ├── config.py                  <- Cargador y validador de configuración
│   ├── models.py                  <- Clases de datos y modelos
│   ├── normalizer.py              <- Algoritmos de limpieza, fechas, JCE, NCF
│   ├── processor.py               <- Árbol de decisión y cruce con maestro
│   └── writers.py                 <- Generador CSV 606, Excel con fórmulas e incidencias
├── tests/                         <- Suite de pruebas automáticas
│   ├── __init__.py
│   ├── test_normalizer.py
│   ├── test_processor.py
│   └── test_integration.py
├── Proveedores.xlsx               <- Maestro de proveedores
├── muestra.xlsx                   <- Datos de prueba
└── README.md                      <- Este manual
```
