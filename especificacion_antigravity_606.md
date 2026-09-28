# Especificación: generador del formato 606 (DGII) y agrupación de compras sin comprobante

## 1. Objetivo

Construir un script/programa (preferiblemente **Python 3.10+**, con `pandas` y `openpyxl`) que:

1. Lea un **archivo grande de transacciones** (Excel, exportado del sistema contable) y el **maestro de proveedores** (Excel).
2. Separe los registros **con comprobante fiscal válido** de los que **no lo tienen**.
3. Genere **dos archivos de salida**:
   - **A) `formato_606.csv`**: solo los registros con comprobante válido, en el formato 606 de la DGII (compras de bienes y servicios).
   - **B) `agrupacion_sin_comprobante.xlsx`**: los registros sin comprobante, agrupados por proveedor (con retención de renta calculada) y los que pasan a gastos menores (B13).
4. Imprima (y opcionalmente guarde en `incidencias.txt`) un reporte de incidencias y advertencias.

Nada debe quedar fijo en el código (rutas, tasas, códigos por defecto, nombres de columnas): todo va en argumentos de línea de comandos o en un archivo de configuración (YAML/JSON).

Uso sugerido:

```
python generar_606.py --transacciones grande.xlsx --proveedores Proveedores.xlsx --salida ./salida --tasa-retencion 0.10
```

---

## 2. Archivos de entrada

### 2.1 Transacciones (ejemplo: `muestra.xlsx`, hoja única)

| Columna | Contenido |
|---|---|
| `Fecha` | Fecha del registro. **Formato mixto**: texto `MM/DD/YYYY` (ej. `08/13/2026`) y celdas de fecha reales de Excel (formato `mm-dd-yy`). Hay que soportar ambos tipos. |
| `Tipo de transacción` | En la muestra solo aparece `Factura de proveedor`. |
| `Número` | NCF/comprobante. Puede venir vacío o con basura (ej. `Ene30`, `00003`). |
| `Proveedor` | Nombre del proveedor, tal como está en el maestro (llave de cruce). |
| `Producto/servicio` | Vacío en la muestra. |
| `Nota` | Texto libre, puede tener saltos de línea (`\n`). Ej.: `Adalfi Bido Garcia\n402-3581367-8\nServicio`. |
| `Importe` | **Total de la factura con impuestos incluidos.** |
| `Impuesto` | ITBIS. |
| `$$Departamento$$` | Entidad (`ENTIDAD 1`, `ENTIDAD 2` o el nombre de una empresa). |
| `Importe sujeto a impuestos` | **Subtotal de la factura** (sin impuestos). Viene **en negativo**: usar el valor absoluto. |
| `C/P pagadas` | Estado de pago (`Pagado`). |

Regla de integridad: `abs(Importe sujeto a impuestos) + Impuesto == Importe` (a 2 decimales). Si no cuadra, registrar una incidencia.

### 2.2 Maestro de proveedores (`Proveedores.xlsx`, hoja `Hoja1`, ~1,476 filas)

| Columna | Contenido |
|---|---|
| `Proveedor` | Llave de cruce con la columna `Proveedor` de transacciones. |
| `RNC/CEDULA` | RNC (9 dígitos) o cédula (11 dígitos). **Datos sucios**, ver abajo. |
| `RAZON SOCIAL` | Razón social. |
| `Tipo Bienes y Servicios Comprados` | Prácticamente vacía hoy (contiene basura en 5 filas). No confiar en ella, pero si en el futuro se llena con un código válido (01 a 11), usarla como fuente prioritaria. |

Problemas conocidos en `RNC/CEDULA` (validar siempre):
- Longitudes distintas de 9 u 11 (6, 7, 10, 12 y hasta 18 dígitos).
- Valores vacíos.
- Texto como `Pasaporte 179054679`. **No extraer dígitos de un texto así**: un valor solo es RNC/cédula válido si, quitando espacios y guiones, queda **solo dígitos** y longitud **9 u 11**.

### 2.3 Reglas de lectura

- Leer **todo como texto** al inicio (`dtype=str` u `object`) para no perder ceros a la izquierda (ej. `06600054594`).
- Los importes se convierten con `Decimal` y se redondean a 2 decimales (`ROUND_HALF_UP`). No usar `float` para calcular.
- Parsear `Fecha`: si es `datetime`, usarlo tal cual; si es texto, interpretarlo como **`MM/DD/YYYY`**.

---

## 3. Flujo de decisión por registro

Para cada fila de transacciones:

```
1. ¿"Número" es un NCF válido?
   └─ SÍ → va al 606 (sección 4).
   └─ NO → registro SIN comprobante:
        2. ¿El proveedor es DJ Ludo? (normalizado contiene "dj ludo")
           └─ SÍ → GASTOS MENORES (B13), aunque tenga pasaporte, RNC o cédula.
        3. Buscar RNC/Cédula:
           a) en el maestro de proveedores (si es válido, sección 2.2);
           b) si no, en la columna Nota (regex de la sección 5.2).
           └─ No se encuentra en ninguno → GASTOS MENORES (B13).
           └─ Se encuentra → AGRUPADO POR PROVEEDOR (sección 5).
```

### 3.1 NCF válido

Regex sobre `Número` (recortado y en mayúsculas):

```
^B(01|02|03|04|11|12|13|14|15|16|17)\d{8}$      # NCF tradicional: B + 2 de tipo + 8 de secuencia = 11 caracteres
^E(31|32|33|34|41|43|44|45|46|47)\d{10}$        # e-CF: E + 2 de tipo + 10 de secuencia = 13 caracteres
```

Cualquier otra cosa (vacío, `Ene30`, `00003`, etc.) = **sin comprobante**.

### 3.2 Normalización de nombres (para cruzar proveedores)

`strip`, colapsar espacios múltiples, quitar acentos y pasar a minúsculas (`casefold`). El cruce transacciones ↔ maestro es por coincidencia exacta del nombre normalizado. Si un proveedor no aparece en el maestro, registrar incidencia (no inventar coincidencias; opcionalmente sugerir candidatos por similitud, sin aplicarlos automáticamente).

---

## 4. Salida A: `formato_606.csv`

- Delimitador coma, codificación **UTF-8 con BOM** (`utf-8-sig`), una fila por factura con NCF válido.
- Importes con punto decimal, 2 decimales, sin separador de miles.
- Los RNC/cédulas son **texto** (conservar ceros a la izquierda).

### 4.1 Encabezados exactos (25 columnas, en este orden)

```
RNC o Cédula, Tipo Id, Tipo Bienes y Servicios Comprados, NCF, NCF ó Documento Modificado,
Fecha Comprobante, (vacío), Fecha Pago, (vacío),
Monto Facturado en Servicios, Monto Facturado en Bienes, Total Monto Facturado, ITBIS Facturado,
ITBIS Retenido, ITBIS sujeto a Proporcionalidad (Art. 349), ITBIS llevado al Costo, ITBIS por Adelantar,
ITBIS percibido en compras, Tipo de Retención en ISR, Monto Retención Renta, ISR Percibido en compras,
Impuesto Selectivo al Consumo, Otros Impuesto/Tasas, Monto Propina Legal, Forma de Pago
```

Las fechas replican la plantilla de Excel de la DGII: cada fecha ocupa **dos columnas** (la segunda con el encabezado en blanco): **`AAAAMM`** y **`DD`**. Ej.: 17 de agosto de 2026 → `202608` y `17`.

### 4.2 Mapeo columna por columna

| Columna 606 | Origen / regla |
|---|---|
| RNC o Cédula | Del maestro (solo dígitos, 9 u 11). Si falta o es inválido: incidencia. Excepción: un NCF tipo B13 no requiere RNC, no reportar incidencia. |
| Tipo Id | `1` si tiene 9 dígitos (RNC); `2` si tiene 11 (cédula). |
| Tipo Bienes y Servicios Comprados | Código del maestro si es válido (01 a 11); si no, el **valor por defecto configurable** (hoy `09`, ver sección 8). |
| NCF | El valor de `Número`, en mayúsculas. |
| NCF ó Documento Modificado | Vacío (solo aplica a notas de crédito/débito). |
| Fecha Comprobante | `AAAAMM` y `DD` de `Fecha`. |
| Fecha Pago | Vacía (el dato no está en el archivo; la DGII la exige cuando hay retención de ITBIS o de renta). |
| Monto Facturado en Servicios | `0.00` por defecto (ver sección 8). |
| Monto Facturado en Bienes | Subtotal (`abs(Importe sujeto a impuestos)`) por defecto. |
| Total Monto Facturado | Servicios + Bienes. |
| ITBIS Facturado | Columna `Impuesto`. |
| ITBIS sujeto a Proporcionalidad | `0.00`. |
| ITBIS llevado al Costo | `0.00`. |
| ITBIS por Adelantar | `ITBIS Facturado − ITBIS sujeto a Proporcionalidad − ITBIS llevado al Costo`. |
| Forma de Pago | `02` si `C/P pagadas` = `Pagado` (sin importar mayúsculas); `04` (compra a crédito) en cualquier otro caso. |
| ITBIS Retenido, ITBIS percibido, Tipo de Retención ISR, Monto Retención Renta, ISR Percibido, ISC, Otros Impuestos/Tasas, Propina Legal | Vacías (no hay datos en el archivo). |

### 4.3 Tablas de referencia (verificar contra la norma vigente de la DGII)

- **Tipo Bienes y Servicios Comprados:** 01 Gastos de personal · 02 Trabajos, suministros y servicios · 03 Arrendamientos · 04 Activos fijos · 05 Representación · 06 Otras deducciones admitidas · 07 Gastos financieros · 08 Gastos extraordinarios · 09 Compras y gastos que formarán parte del costo de venta · 10 Adquisiciones de activos · 11 Gastos de seguros.
- **Forma de Pago:** 01 Efectivo · 02 Cheques/Transferencias/Depósito · 03 Tarjeta crédito/débito · 04 Compra a crédito · 05 Permuta · 06 Notas de crédito · 07 Mixto.
- **Tipo de Retención en ISR:** 01 Alquileres · 02 Honorarios por servicios · 03 Otras rentas · 04 Otras rentas (rentas presuntas) · 05 Intereses a personas jurídicas residentes · 06 Intereses a personas físicas residentes · 07 Retención por proveedores del Estado · 08 Juegos telefónicos.

---

## 5. Salida B: `agrupacion_sin_comprobante.xlsx`

Fuente tipográfica Arial en todo el libro. Tres hojas.

### 5.1 Clasificación bien / servicio (según la columna `Nota`)

Sobre la nota normalizada (sin acentos, minúsculas):
- Contiene `servicio` → **Servicio** (aplica retención).
- Contiene la palabra `bien` o `bienes` → **Bien** (no lleva retención).
- Ninguna de las dos → **Sin clasificar** (no lleva retención, pero se marca con observación para que el usuario lo defina).

La lista de palabras clave debe ser configurable. Si en el archivo grande la columna `Producto/servicio` viene llena, evaluar usarla como señal secundaria (confirmar con el usuario).

### 5.2 Extracción de RNC/cédula desde la nota

Se usa solo si el proveedor **no** tiene RNC/cédula válido en el maestro.

- Cédula: `(?<!\d)(\d{3}-?\d{7}-?\d)(?!\d)` → quitar guiones. Ej.: `402-3581367-8` → `40235813678`.
- RNC: 9 dígitos, con o sin guiones (`\d-\d{2}-\d{5}-\d` o `\d{9}`).
- Un pasaporte alfanumérico (ej. `93LY5366192`) **no** cuenta como RNC/cédula.
- Si el maestro y la nota traen identificadores distintos, **prevalece el maestro** y se agrega una observación con ambos valores.
- Advertencia (no bloqueante): validar el dígito verificador de la cédula (algoritmo de la JCE: pesos 1,2,1,2… sobre los primeros 10 dígitos, sumando los dígitos de los productos mayores que 9; verificador = `(10 − suma % 10) % 10`).

### 5.3 Hoja 1: `Agrupado por proveedor`

Encabezado en la fila 5; parámetros arriba. **Celda B3 = tasa de retención** (por defecto `10%`, editable, en azul con fondo amarillo y con nota explicativa). Una fila por proveedor (clave: nombre normalizado):

| Col | Encabezado | Contenido |
|---|---|---|
| A | RNC o Cédula | Texto. |
| B | Tipo Id | `1` (9 dígitos) o `2` (11 dígitos). |
| C | Proveedor | Nombre. |
| D | Registros | `COUNTIFS` sobre la hoja `Detalle`. |
| E | Servicios (sujeto a retención) | `SUMIFS` de subtotal, tipo = Servicio. |
| F | Bienes | `SUMIFS`, tipo = Bien. |
| G | Sin clasificar | `SUMIFS`, tipo = Sin clasificar. |
| H | Total facturado (subtotal) | `=SUM(E:G)`. |
| I | ITBIS | `SUMIFS` de ITBIS (en estos casos siempre 0). |
| J | Retención de renta | `=ROUND(E * $B$3, 2)`. Solo sobre servicios. |
| K | Tipo Retención ISR | `=IF(E>0,"02","")`. |
| L | NCF (por asignar) | **Vacío**: el usuario lo asigna después según la última asignación. |
| M | Observaciones | Advertencias (fondo amarillo claro si hay). |

Fila `TOTAL` con `SUM` de D a J.

### 5.4 Hoja 2: `Gastos menores B13`

Una fila por registro (no se agrupan): `Fecha, Proveedor, Nota, Subtotal, ITBIS, Importe, Entidad, Motivo, NCF (por asignar)`, con fila `TOTAL`. Entran aquí:
- Registros sin comprobante y **sin** RNC/cédula (ni en maestro ni en nota).
- Todos los del proveedor **DJ Ludo**, aunque tengan pasaporte, RNC o cédula (motivo: `Excepción DJ Ludo`).

Ninguno de estos lleva ITBIS.

### 5.5 Hoja 3: `Detalle`

Todos los registros sin comprobante, una fila por registro: `Fecha (yyyy-mm-dd), Proveedor, RNC o Cédula, Fuente del RNC (Maestro/Nota), Tipo (según nota), Destino (Agrupado / B13 Gastos menores), Nota (saltos de línea reemplazados por " | "), Subtotal, ITBIS, Importe, Entidad, Número (no válido), Observaciones`. Es la fuente de las fórmulas de la hoja 1.

### 5.6 Reglas técnicas del Excel

- Usar **fórmulas** (no valores calculados en Python) para los totales, sumas por proveedor y retención, de modo que el libro recalcule al cambiar la tasa o corregir un tipo en `Detalle`.
- Activar `wb.calculation.fullCalcOnLoad = True` para que Excel recalcule al abrir.
- Formato numérico `#,##0.00;(#,##0.00);-`, encabezados con fondo oscuro y letra blanca, paneles inmovilizados, anchos de columna razonables, RNC con formato de texto (`@`).

---

## 6. Incidencias que el programa debe reportar

- Proveedor de la transacción que no existe en el maestro.
- Factura con NCF válido cuyo proveedor no tiene RNC/cédula válido en el maestro.
- Subtotal + ITBIS ≠ Importe.
- Maestro y nota con identificadores distintos.
- Dígito verificador de cédula inválido.
- Registro agrupado con tipo **Sin clasificar** (la nota no dice si es bien o servicio).
- Proveedor que aparece en más de una entidad (`$$Departamento$$`) dentro de un mismo grupo.
- Fechas fuera del período que se está declarando (ver sección 8).
- Valores de `Tipo de transacción` distintos de `Factura de proveedor` (ver sección 8).

---

## 7. Caso de prueba con la muestra (resultado esperado)

Entrada (9 filas):

| # | Fecha | Número | Proveedor | Nota | Importe | ITBIS | Subtotal | Entidad |
|---|---|---|---|---|---|---|---|---|
| 1 | 08/01/2026 | B0100000188 | Agropin (Pablo Beato Espinal) | | 2400 | 366.10 | 2033.90 | ENTIDAD 1 |
| 2 | 08/12/2026 | B0100000204 | Agropin (Pablo Beato Espinal) | | 4000 | 610.17 | 3389.83 | ENTIDAD 2 |
| 3 | 08/17/2026 | B0100000383 | Alta Gama Liquor Store | | 30600.41 | 3642.77 | 26957.64 | ENTIDAD 1 |
| 4 | 08/13/2026 | (vacío) | Adalfi Bido Garcia | `…402-3581367-8 / Servicio` | 1500 | 0 | 1500 | ENTIDAD 1 |
| 5 | 2026-05-08 (fecha real) | (vacío) | Adalfi Bido Garcia | `…402-3581367-8 / Servicio Cocina` | 1000 | 0 | 1000 | ENTIDAD 1 |
| 6 | 2026-08-08 (fecha real) | (vacío) | Adelaida Sarante | | 15000 | 0 | 15000 | ENTIDAD 1 |
| 7 | 01/30/2026 | Ene30 | Domingo Ebanista | `Diomedes de los Santos 066-0005003-0 (concho Yiyo)` | 13000 | 0 | 13000 | ENTIDAD 1 |
| 8 | 07/21/2026 | 00003 | Domingo Ebanista | (igual que la 7) | 19570 | 0 | 19570 | ENTIDAD 2 |
| 9 | 02/07/2026 | (vacío) | DJ Ludo Yhouberth Roselv Ludovic Rodriguez | `…93LY5366192 / DJ LUDO` | 8000 | 0 | 8000 | Inversiones Las Mosquiteras Unlimited SRL |

**`formato_606.csv` esperado** (filas 1 a 3; el encabezado va antes, ver sección 4.1):

```
06600056540,2,09,B0100000188,,202608,01,,,0.00,2033.90,2033.90,366.10,,0.00,0.00,366.10,,,,,,,,02
06600056540,2,09,B0100000204,,202608,12,,,0.00,3389.83,3389.83,610.17,,0.00,0.00,610.17,,,,,,,,02
132465555,1,09,B0100000383,,202608,17,,,0.00,26957.64,26957.64,3642.77,,0.00,0.00,3642.77,,,,,,,,02
```

**Hoja `Agrupado por proveedor` esperada** (tasa 10%):

| RNC/Cédula | Proveedor | Registros | Servicios | Sin clasificar | Total | Retención | Observación |
|---|---|---|---|---|---|---|---|
| 40235813678 | Adalfi Bido Garcia | 2 | 2,500.00 | 0 | 2,500.00 | 250.00 | (ninguna) |
| 13400049261 | Adelaida Sarante | 1 | 0 | 15,000.00 | 15,000.00 | 0 | Nota no indica bien o servicio |
| 06600054594 | Domingo Ebanista | 2 | 0 | 32,570.00 | 32,570.00 | 0 | Maestro (06600054594) y nota (06600050030) no coinciden; se usó el maestro; nota no indica bien o servicio |

Totales: registros 5, subtotal 50,070.00, retención 250.00.

**Hoja `Gastos menores B13` esperada:** una fila, DJ Ludo, 8,000.00, motivo `Excepción DJ Ludo`.

Notas del caso: el RNC de Adalfi está en el maestro (`40235813678`) y coincide con la nota. Domingo Ebanista está en el maestro con `06600054594`, pero su nota trae otra cédula (`06600050030`, Diomedes de los Santos); por la regla "maestro primero" se usa la del maestro y se advierte. El maestro de DJ Ludo trae `Pasaporte 179054679`, que no es un RNC/cédula válido.

---

## 8. Supuestos y decisiones pendientes (deben ser configurables)

| Tema | Valor actual | Pendiente de confirmar |
|---|---|---|
| Tipo Bienes y Servicios (606) | `09` para todos | Regla por proveedor o categoría; el maestro trae la columna pero está vacía. |
| Bienes vs Servicios en el 606 | Todo como **Bienes** | Regla para separar el monto en servicios y bienes. |
| Forma de pago | `02` si Pagado, `04` si no | Método real de pago (efectivo, tarjeta, etc.). |
| Fecha de pago | Vacía | De dónde sacarla; es obligatoria cuando hay retenciones. |
| Tasa de retención de renta | 10% sobre servicios | Confirmar la tasa aplicable. Tipo de retención `02`. |
| Líneas "Sin clasificar" | Sin retención, con observación | Regla por defecto (¿servicio o bien?). |
| Proveedor en el maestro pero sin nota (caso Adelaida) | Se agrupa como proveedor | Confirmar que no debe ir a B13. |
| Conflicto maestro vs nota (caso Domingo Ebanista) | Prevalece el maestro | Confirmar si debe prevalecer la nota. |
| Proveedor en varias entidades | Se agrupa sin separar | ¿Separar por entidad? Cada empresa declara su propio 606. |
| Período | No se filtra por fecha | Parámetro `--periodo AAAAMM` para filtrar el archivo grande. |
| NCF de los agrupados y de B13 | Vacío | El usuario los asigna después según la última asignación (posible mejora: cargar la última secuencia y asignarlos automáticamente). |
| Otros tipos de transacción (notas de crédito, etc.) | No contemplados | Si aparecen, definir NCF B04, columna "NCF ó Documento Modificado" y manejo de signos. |

---

## 9. Requisitos de calidad

- Código modular: lectura, normalización, clasificación, cálculo, escritura de salidas.
- Configuración externa (YAML/JSON) para: nombres de columnas, palabras clave de bien/servicio, lista de proveedores con excepción (hoy `DJ Ludo`), tasa de retención, códigos por defecto y mapeo de forma de pago.
- Pruebas unitarias con `pytest` usando el caso de la sección 7 como fixture (incluye NCF inválidos, fechas mixtas, pasaporte en el maestro, nota con saltos de línea y conflicto maestro/nota).
- Rendimiento: usar diccionarios para el cruce con el maestro (no búsquedas fila por fila en el DataFrame).
- Mensajes de error claros si faltan columnas obligatorias en cualquiera de los dos archivos.
- Documentar en un `README` cómo ejecutarlo y qué significa cada advertencia.
