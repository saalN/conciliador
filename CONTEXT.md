# Conciliación Bancaria — Contexto del Proyecto

## Descripción
Aplicación de escritorio Python (Tkinter) que cruza un **listado de facturas pendientes** (Excel) contra un **extracto bancario** (Excel), generando un Excel de resultado coloreado con el estado de cobro de cada factura.

## Stack
- Python 3
- pandas, openpyxl, xlrd (ver `requirements.txt`)
- Tkinter (interfaz gráfica nativa)

## Archivo principal
`conciliacion.py` — único archivo de código fuente.

---

## Estructura del código

### Constantes de color (líneas ~20-23)
```python
BG = "#F8F8F6"
BLUE_DARK = "1F4E79"   # cabecera Excel
GREEN_FILL = "C6EFCE"  # fila cobrada
RED_FILL   = "FFC7CE"  # fila pendiente
AMBER_FILL = "FFEB9C"  # observaciones/avisos
```

### Funciones auxiliares (privadas, prefijo `_`)
| Función | Propósito |
|---|---|
| `_limpiar(valor)` | Elimina caracteres ilegales en XML 1.0 de strings |
| `_normalizar(texto)` | Primeros 15 chars en minúsculas sin espacios |
| `_añadir_observaciones(df, col_importe, col_factura)` | Marca duplicados de importe y/o nº factura con ⚠ |
| `_formatear_fecha(valor)` | Convierte a `DD/MM/YYYY` con formatos explícitos |
| `_aplicar_formato_texto_fechas(ws, df)` | Marca celdas de fecha como texto (@) en openpyxl |
| `_detectar_cols_pancho(facturas)` | Auto-detecta columnas nombre/importe en listados Pancho |

### Funciones de conciliación (10 funciones)
Todas devuelven el mismo `dict` de estadísticas:
```python
{
    "total": int,
    "cobradas": int,
    "pendientes": int,
    "importe_cobradas": float,
    "importe_pendientes": float,
    "avisos_importe": int,
    "avisos_factura": int,
    "ruta_resultado": str,
}
```

#### Listado tipo Pancho
| Función | Banco | skiprows banco | Columnas banco clave |
|---|---|---|---|
| `conciliar_bankinter` | Bankinter | 5 | CATEGORÍA, HABER, FECHA CONTABLE, REFERENCIA |
| `conciliar_abanca` | Abanca | 5 | TIPO OPERACIÓN, IMPORTE, F. OPERACIÓN, REFERENCIA |
| `conciliar_lacaixa` | La Caixa | 3, sin cabecera (header=None) | col 1=fecha, col 3=desc, col 4=importe |
| `conciliar_bbva` | BBVA | 15 | F. CONTABLE, CONCEPTO, BENEFICIARIO/ORDENANTE, IMPORTE |
| `conciliar_unicaja` | Unicaja | 10 | Fecha de operación, Importe, Nº mov |
| `conciliar` | Original/legacy | 3 | Importe, Tipo movimiento, Fecha de la operación, Nro. Apunte |

#### Listado tipo SIDI (columnas: `Cliente/SubCliente`, `Total`, `Código`)
| Función | Banco |
|---|---|
| `conciliar_bankinter_sidi` | Bankinter |
| `conciliar_abanca_sidi` | Abanca |
| `conciliar_lacaixa_sidi` | La Caixa |
| `conciliar_bbva_sidi` | BBVA |
| `conciliar_unicaja_sidi` | Unicaja |

### Clase `App(tk.Tk)`
Interfaz gráfica principal. Métodos relevantes:
- `_build_ui()` — construye todos los widgets
- `_actualizar_labels()` — actualiza textos al cambiar banco/listado
- `_lanzar()` — valida entradas y lanza hilo
- `_ejecutar_hilo(banco, facturas, modo, tipo_listado)` — selecciona y llama la función correcta
- `_mostrar_resultado(stats)` — actualiza la UI con resultados
- `_mostrar_error(mensaje)` — muestra error en UI
- `_abrir_fichero(ruta)` — abre el Excel resultado (multiplataforma)

---

## Formatos de entrada esperados

### Extractos bancarios
Cada banco tiene su propio formato de cabecera. El programa usa `skiprows` para saltar filas de metadatos:
- **Bankinter**: 5 filas de cabecera → columnas: `CATEGORÍA`, `DESCRIPCIÓN`, `HABER`, `FECHA CONTABLE`, `REFERENCIA`
- **Abanca**: 5 filas → `TIPO OPERACIÓN`, `IMPORTE`, `F. OPERACIÓN`, `REFERENCIA`
- **BBVA**: 15 filas → `F. CONTABLE`, `CONCEPTO`, `BENEFICIARIO/ORDENANTE`, `IMPORTE`
- **La Caixa**: 3 filas, sin cabecera (columnas por posición: 0=tipo, 1=fecha, 3=desc, 4=importe)
- **Unicaja**: 10 filas → `Fecha de operación`, `Importe`, `Nº mov`
- **Original**: 3 filas → `Importe`, `Tipo movimiento`, `Fecha de la operación`, `Nro. Apunte`

### Listados de facturas
- **Pancho**: 2 filas de cabecera (`skiprows=2`). Columnas: `TIPO EFECTO`, `IMPORTE`, `RAZON SOCIAL`. Solo se procesan filas con `TIPO EFECTO == "TRANSFERENCIA"` e importe > 0. Auto-detecta también `EMPRESA`/`Nombre` y `Total`/`IMPORTE`/`Importe`.
- **SIDI**: Sin `skiprows`. Columnas fijas: `Cliente/SubCliente`, `Total`, `Código` (para detección de duplicados de factura).

---

## Lógica de cruce
1. Se filtran facturas con importe > 0
2. Se filtran movimientos bancarios de tipo transferencia con importe positivo
3. Por cada factura se busca en `banco_disponible` el primer movimiento con **importe exactamente igual**
4. En `conciliar` (original) también se verifica que el nombre de la razón social esté contenido en el campo `Tipo movimiento` del banco
5. El movimiento encontrado se elimina de `banco_disponible` (evita doble asignación)
6. Se añaden columnas: `COBRADA`, `FECHA COBRO`, `NRO. APUNTE BANCO`

---

## Salida generada
Archivo Excel guardado en la **misma carpeta que el fichero de facturas**:
- Nombre: `conciliacion_<Banco>[_SIDI]_DD-MM-YYYY_HH-MM-SS.xlsx`
- Filas verdes = cobradas, rojas = pendientes, ámbar en OBSERVACIONES = duplicados
- Columnas de fecha formateadas como texto para evitar auto-conversión en LibreOffice

---

## Convenciones de código
- Funciones privadas con prefijo `_`
- Separadores visuales con `# ──` para bloques dentro de funciones largas
- Todas las funciones de conciliación siguen el mismo patrón (leer banco → leer facturas → filtrar → cruzar → exportar → estadísticas)
- Los importes se leen como `str` y se limpian manualmente (`,` → `.`, quitar símbolos) antes de convertir a numérico
- Se usa `dtype=str` al leer Excel para evitar conversiones automáticas

---

## Posibles mejoras / puntos de atención
- La comparación de importe es **exacta** (`==` entre floats); puede fallar con diferencias de redondeo
- `conciliar` (original) es la única función con match por nombre; el resto solo cruza por importe
- No hay tests automatizados
- La función `conciliar_lacaixa` no filtra por tipo de operación (acepta cualquier movimiento positivo)
- El código de exportación Excel está duplicado en cada función (candidato a refactorizar en helper)
