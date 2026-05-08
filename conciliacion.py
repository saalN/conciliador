"""
Conciliación de facturas pendientes contra movimientos bancarios.
Dependencias: pandas, openpyxl, xlrd
"""

import threading
import subprocess
import sys
import os
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ── Colores ──────────────────────────────────────────────────────────────────
BG = "#F8F8F6"
BLUE_DARK = "1F4E79"
GREEN_FILL = "C6EFCE"
RED_FILL   = "FFC7CE"
AMBER_FILL = "FFEB9C"


# ── Lógica de conciliación ───────────────────────────────────────────────────

import re as _re

# Caracteres no permitidos en XML 1.0 (los xlsx son XML internamente)
_ILEGAL_XML = _re.compile(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFFFE\uFFFF]"
)


def _limpiar(valor):
    """Elimina caracteres ilegales en XML de cadenas de texto."""
    if isinstance(valor, str):
        return _ILEGAL_XML.sub("", valor)
    return valor


def _normalizar(texto: str) -> str:
    """Devuelve los primeros 15 caracteres en minúsculas sin espacios extremos."""
    return str(texto).strip().lower()[:15]


def _añadir_observaciones(df: pd.DataFrame, col_importe: str, col_factura: str = "Factura") -> pd.DataFrame:
    """Marca con aviso las filas con importe duplicado y/o número de factura duplicado."""
    dup_importe = df[col_importe].duplicated(keep=False)
    dup_factura = (
        df[col_factura].duplicated(keep=False)
        if col_factura in df.columns
        else pd.Series(False, index=df.index)
    )

    def _aviso(imp_dup: bool, fac_dup: bool) -> str:
        msgs = []
        if imp_dup:
            msgs.append("⚠ Importe duplicado - revisar")
        if fac_dup:
            msgs.append("⚠ Nº factura duplicado - revisar")
        return " | ".join(msgs)

    df["OBSERVACIONES"] = [_aviso(i, f) for i, f in zip(dup_importe, dup_factura)]
    return df


def _formatear_fecha(valor) -> str:
    """Convierte una fecha al formato DD/MM/YYYY usando formatos explícitos para evitar
    la ambigüedad de dayfirst en pandas (que no es estricto en versiones recientes)."""
    if not valor or str(valor).strip() == "":
        return ""
    # Si ya es un objeto date/datetime, formatear directamente sin re-parsear
    if isinstance(valor, (datetime.date, datetime.datetime)):
        return valor.strftime("%d/%m/%Y")
    s = str(valor).strip()
    # Probar formatos explícitos en orden: DD/MM/YYYY primero
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s


def _aplicar_formato_texto_fechas(ws, df: pd.DataFrame) -> None:
    """Marca como texto (@) las columnas de fecha para evitar que LibreOffice
    las auto-convierta al abrir el archivo y muestre el formato MM/DD en edición."""
    cols_fecha = [c for c in df.columns if "fecha" in c.lower() or c == "Fecha"]
    for col_name in cols_fecha:
        col_idx = df.columns.get_loc(col_name) + 1
        for row_idx in range(1, ws.max_row + 1):
            ws.cell(row=row_idx, column=col_idx).number_format = "@"


def conciliar(ruta_banco: str, ruta_facturas: str) -> dict:
    """
    Cruza el extracto bancario con las facturas y genera resultado_conciliacion.xlsx.
    Devuelve un dict con estadísticas.
    """
    # ── Leer banco ────────────────────────────────────────────────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=3)
    banco.columns = [c.strip() for c in banco.columns]

    col_importe_banco = "Importe"
    col_tipo_mov = "Tipo movimiento"
    col_fecha_banco = "Fecha de la operación"
    col_apunte = "Nro. Apunte"

    for col in (col_importe_banco, col_tipo_mov, col_fecha_banco, col_apunte):
        if col not in banco.columns:
            raise ValueError(f"Columna no encontrada en el banco: '{col}'\nColumnas disponibles: {list(banco.columns)}")

    banco[col_importe_banco] = (
        banco[col_importe_banco]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_banco] = pd.to_numeric(banco[col_importe_banco], errors="coerce")

    # ── Leer facturas ─────────────────────────────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd", skiprows=2)
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl", skiprows=2)
    facturas.columns = [c.strip() for c in facturas.columns]

    col_tipo_ef = "TIPO EFECTO"
    col_importe_fac = "IMPORTE"
    col_razon = "RAZON SOCIAL"

    for col in (col_tipo_ef, col_importe_fac, col_razon):
        if col not in facturas.columns:
            raise ValueError(f"Columna no encontrada en facturas: '{col}'\nColumnas disponibles: {list(facturas.columns)}")

    facturas[col_importe_fac] = (
        facturas[col_importe_fac]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe_fac] = pd.to_numeric(facturas[col_importe_fac], errors="coerce")

    # ── Filtrar transferencias positivas ──────────────────────────────────────
    mask = (
        (facturas[col_tipo_ef].str.strip().str.upper() == "TRANSFERENCIA") &
        (facturas[col_importe_fac] > 0)
    )
    df = facturas[mask].copy().reset_index(drop=True)

    # ── Cruce ─────────────────────────────────────────────────────────────────
    cobrada_flags = []
    fechas_cobro = []
    apuntes_banco = []

    banco_disponible = banco.copy()

    for _, fac in df.iterrows():
        importe_fac = fac[col_importe_fac]
        razon_norm = _normalizar(fac[col_razon])

        coincidencia = banco_disponible[
            (banco_disponible[col_importe_banco] == importe_fac) &
            (banco_disponible[col_tipo_mov].apply(lambda x: str(x).strip().lower()).str.contains(razon_norm, regex=False))
        ]

        if not coincidencia.empty:
            idx = coincidencia.index[0]
            row_banco = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_banco[col_fecha_banco])
            apuntes_banco.append(row_banco[col_apunte])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            apuntes_banco.append("")

    df["COBRADA"] = cobrada_flags
    df["FECHA COBRO"] = fechas_cobro
    df["NRO. APUNTE BANCO"] = apuntes_banco

    df = _añadir_observaciones(df, col_importe_fac)
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    # ── Exportar Excel con formato ────────────────────────────────────────────
    carpeta = os.path.dirname(ruta_facturas)
    fecha_hoy = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Original_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)

    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill = PatternFill(fill_type="solid", fgColor=RED_FILL)

    # Cabecera
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Filas de datos
    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1  # 1-based

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cobrada_val = ws.cell(row=row[0].row, column=cobrada_col_idx).value
        fill = green_fill if cobrada_val == "SÍ" else red_fill
        for cell in row:
            cell.fill = fill

    # Celda OBSERVACIONES en ámbar si tiene aviso
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value:
                cell.fill = amber_fill
                cell.font = amber_font

    # Autoajuste de columnas
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    wb.save(ruta_resultado)

    # ── Estadísticas ──────────────────────────────────────────────────────────
    n_total = len(df)
    cobradas = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]

    return {
        "total": n_total,
        "cobradas": len(cobradas),
        "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe_fac].sum(),
        "importe_pendientes": pendientes[col_importe_fac].sum(),
        "avisos_importe": len(dup_importe),
        "avisos_factura": len(dup_factura),
        "ruta_resultado": ruta_resultado,
    }


def conciliar_bankinter(ruta_banco: str, ruta_facturas: str) -> dict:
    """
    Cruza facturas (Nombre/Importe) contra extracto Bankinter.
    Genera resultado_conciliacion.xlsx.
    """
    # ── Leer banco Bankinter (fila 6 = cabecera) ──────────────────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=5)
    banco.columns = [c.strip() for c in banco.columns]

    col_categoria   = "CATEGORÍA"
    col_descripcion = "DESCRIPCIÓN"
    col_haber       = "HABER"
    col_fecha       = "FECHA CONTABLE"
    col_referencia  = "REFERENCIA"

    for col in (col_categoria, col_descripcion, col_haber, col_fecha, col_referencia):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco Bankinter: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_haber] = (
        banco[col_haber]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_haber] = pd.to_numeric(banco[col_haber], errors="coerce")

    # ── Leer facturas (fila 1 = cabecera) ────────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Nombre"
    col_importe = "Importe"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    # ── Solo facturas con importe positivo ────────────────────────────────────
    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    # ── Banco: filtrar solo Transferencias con HABER positivo ────────────────
    categorias_validas = {"transferencias"}
    banco_disponible = banco[
        banco[col_categoria].fillna("").str.strip().str.lower().isin(categorias_validas) &
        (banco[col_haber] > 0)
    ].copy()

    # ── Cruce ─────────────────────────────────────────────────────────────────
    cobrada_flags = []
    fechas_cobro  = []
    referencias   = []

    for _, fac in df.iterrows():
        importe_fac = fac[col_importe]

        coincidencia = banco_disponible[
            banco_disponible[col_haber] == importe_fac
        ]

        if not coincidencia.empty:
            idx   = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_referencia])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe)
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    # ── Exportar Excel con formato ────────────────────────────────────────────
    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Bankinter_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)

    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)

    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cobrada_val = ws.cell(row=row[0].row, column=cobrada_col_idx).value
        fill = green_fill if cobrada_val == "SÍ" else red_fill
        for cell in row:
            cell.fill = fill

    # Celda OBSERVACIONES en ámbar si tiene aviso
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value:
                cell.fill = amber_fill
                cell.font = amber_font

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    wb.save(ruta_resultado)

    # ── Estadísticas ──────────────────────────────────────────────────────────
    n_total    = len(df)
    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]

    return {
        "total":              n_total,
        "cobradas":           len(cobradas),
        "pendientes":         len(pendientes),
        "importe_cobradas":   cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe":     len(dup_importe),
        "avisos_factura":     len(dup_factura),
        "ruta_resultado":     ruta_resultado,
    }


def conciliar_abanca(ruta_banco: str, ruta_facturas: str) -> dict:
    """
    Cruza facturas (Nombre/Importe) contra extracto Abanca.
    Genera resultado_conciliacion.xlsx.
    """
    # ── Leer banco Abanca (fila 6 = cabecera) ─────────────────────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=5)
    banco.columns = [c.strip() for c in banco.columns]

    col_tipo_op   = "TIPO OPERACIÓN"
    col_importe_b = "IMPORTE"
    col_fecha     = "F. OPERACIÓN"
    col_referencia = "REFERENCIA"

    for col in (col_tipo_op, col_importe_b, col_fecha, col_referencia):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco Abanca: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_importe_b] = (
        banco[col_importe_b]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_b] = pd.to_numeric(banco[col_importe_b], errors="coerce")

    # ── Leer facturas (fila 1 = cabecera) ─────────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Nombre"
    col_importe = "Importe"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    # ── Solo facturas con importe positivo ────────────────────────────────────
    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    # ── Banco: filtrar transferencias con IMPORTE positivo ────────────────────
    categorias_validas = {
        "transferencias de otras entidades",
        "transferencias propia entidad",
    }
    banco_disponible = banco[
        banco[col_tipo_op].fillna("").str.strip().str.lower().isin(categorias_validas) &
        (banco[col_importe_b] > 0)
    ].copy()

    # ── Cruce por importe exacto ───────────────────────────────────────────────
    cobrada_flags = []
    fechas_cobro  = []
    referencias   = []

    for _, fac in df.iterrows():
        importe_fac = fac[col_importe]

        coincidencia = banco_disponible[
            banco_disponible[col_importe_b] == importe_fac
        ]

        if not coincidencia.empty:
            idx   = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_referencia])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe)
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    # ── Exportar Excel con formato ─────────────────────────────────────────────
    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Abanca_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)

    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)

    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cobrada_val = ws.cell(row=row[0].row, column=cobrada_col_idx).value
        fill = green_fill if cobrada_val == "SÍ" else red_fill
        for cell in row:
            cell.fill = fill

    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value:
                cell.fill = amber_fill
                cell.font = amber_font

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    wb.save(ruta_resultado)

    # ── Estadísticas ───────────────────────────────────────────────────────────
    n_total    = len(df)
    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]

    return {
        "total":              n_total,
        "cobradas":           len(cobradas),
        "pendientes":         len(pendientes),
        "importe_cobradas":   cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe":     len(dup_importe),
        "avisos_factura":     len(dup_factura),
        "ruta_resultado":     ruta_resultado,
    }


def conciliar_lacaixa(ruta_banco: str, ruta_facturas: str) -> dict:
    """
    Cruza facturas (Nombre/Importe) contra extracto La Caixa.
    El extracto no tiene cabecera; los datos empiezan en la fila 4 (skiprows=3).
    Columnas por posición: 0=tipo, 1=fecha op., 2=fecha valor, 3=descripción, 4=importe, 5=saldo.
    Genera conciliacion_LaCaixa_DD-MM-YYYY.xlsx.
    """
    # ── Leer banco La Caixa (sin cabecera, datos desde fila 4) ────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=3, header=None)

    if banco.shape[1] < 5:
        raise ValueError(
            f"El extracto de La Caixa debe tener al menos 5 columnas. "
            f"Se encontraron {banco.shape[1]}."
        )

    COL_FECHA   = 1
    COL_DESC    = 3
    COL_IMPORTE = 4

    banco[COL_IMPORTE] = (
        banco[COL_IMPORTE]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[COL_IMPORTE] = pd.to_numeric(banco[COL_IMPORTE], errors="coerce")

    # ── Leer facturas (fila 1 = cabecera) ─────────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Nombre"
    col_importe = "Importe"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    # ── Solo facturas con importe positivo ────────────────────────────────────
    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    # ── Banco: solo movimientos con importe positivo ───────────────────────────
    banco_disponible = banco[banco[COL_IMPORTE] > 0].copy()

    # ── Cruce por importe exacto ───────────────────────────────────────────────
    cobrada_flags = []
    fechas_cobro  = []
    referencias   = []

    for _, fac in df.iterrows():
        importe_fac = fac[col_importe]

        coincidencia = banco_disponible[
            banco_disponible[COL_IMPORTE] == importe_fac
        ]

        if not coincidencia.empty:
            idx   = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[COL_FECHA])
            referencias.append(row_b[COL_DESC])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe)
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    # ── Exportar Excel con formato ─────────────────────────────────────────────
    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_LaCaixa_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)

    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)

    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cobrada_val = ws.cell(row=row[0].row, column=cobrada_col_idx).value
        fill = green_fill if cobrada_val == "SÍ" else red_fill
        for cell in row:
            cell.fill = fill

    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill  = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font  = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value:
                cell.fill = amber_fill
                cell.font = amber_font

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    wb.save(ruta_resultado)

    # ── Estadísticas ───────────────────────────────────────────────────────────
    n_total    = len(df)
    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]

    return {
        "total":              n_total,
        "cobradas":           len(cobradas),
        "pendientes":         len(pendientes),
        "importe_cobradas":   cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe":     len(dup_importe),
        "avisos_factura":     len(dup_factura),
        "ruta_resultado":     ruta_resultado,
    }


def conciliar_bbva(ruta_banco: str, ruta_facturas: str) -> dict:
    """
    Cruza facturas (Nombre/Importe) contra extracto BBVA.
    El extracto tiene 15 filas de cabecera; los datos empiezan en la fila 16.
    Genera conciliacion_BBVA_DD-MM-YYYY.xlsx.
    """
    # ── Leer banco BBVA (cabecera en fila 16) ─────────────────────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=15)
    banco.columns = [c.strip() for c in banco.columns]

    col_fecha       = "F. CONTABLE"
    col_concepto    = "CONCEPTO"
    col_beneficiario = "BENEFICIARIO/ORDENANTE"
    col_importe_b   = "IMPORTE"

    for col in (col_fecha, col_concepto, col_beneficiario, col_importe_b):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco BBVA: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_importe_b] = (
        banco[col_importe_b]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_b] = pd.to_numeric(banco[col_importe_b], errors="coerce")

    # ── Leer facturas (fila 1 = cabecera) ─────────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Nombre"
    col_importe = "Importe"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    # ── Solo facturas con importe positivo ────────────────────────────────────
    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    # ── Banco: filtrar TRANSFERENCIAS con IMPORTE positivo ────────────────────
    banco_disponible = banco[
        banco[col_concepto].fillna("").str.strip().str.lower().str.contains("transferencia", regex=False) &
        (banco[col_importe_b] > 0)
    ].copy()

    # ── Cruce por importe exacto ───────────────────────────────────────────────
    cobrada_flags = []
    fechas_cobro  = []
    referencias   = []

    for _, fac in df.iterrows():
        importe_fac = fac[col_importe]

        coincidencia = banco_disponible[
            banco_disponible[col_importe_b] == importe_fac
        ]

        if not coincidencia.empty:
            idx   = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_beneficiario])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe)
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    # ── Exportar Excel con formato ─────────────────────────────────────────────
    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_BBVA_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)

    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)

    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cobrada_val = ws.cell(row=row[0].row, column=cobrada_col_idx).value
        fill = green_fill if cobrada_val == "SÍ" else red_fill
        for cell in row:
            cell.fill = fill

    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill  = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font  = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value:
                cell.fill = amber_fill
                cell.font = amber_font

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    wb.save(ruta_resultado)

    # ── Estadísticas ───────────────────────────────────────────────────────────
    n_total    = len(df)
    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]

    return {
        "total":              n_total,
        "cobradas":           len(cobradas),
        "pendientes":         len(pendientes),
        "importe_cobradas":   cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe":     len(dup_importe),
        "avisos_factura":     len(dup_factura),
        "ruta_resultado":     ruta_resultado,
    }


# ── Funciones SIDI ───────────────────────────────────────────────────────────

def conciliar_bankinter_sidi(ruta_banco: str, ruta_facturas: str) -> dict:
    """Cruza facturas SIDI (Cliente/SubCliente + Total) contra extracto Bankinter."""
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=5)
    banco.columns = [c.strip() for c in banco.columns]

    col_categoria   = "CATEGORÍA"
    col_haber       = "HABER"
    col_fecha       = "FECHA CONTABLE"
    col_referencia  = "REFERENCIA"

    for col in (col_categoria, col_haber, col_fecha, col_referencia):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco Bankinter: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_haber] = (
        banco[col_haber]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_haber] = pd.to_numeric(banco[col_haber], errors="coerce")

    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Cliente/SubCliente"
    col_importe = "Total"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas SIDI: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    categorias_validas = {"transferencias"}
    banco_disponible = banco[
        banco[col_categoria].fillna("").str.strip().str.lower().isin(categorias_validas) &
        (banco[col_haber] > 0)
    ].copy()

    cobrada_flags, fechas_cobro, referencias = [], [], []

    for _, fac in df.iterrows():
        coincidencia = banco_disponible[banco_disponible[col_haber] == fac[col_importe]]
        if not coincidencia.empty:
            idx = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_referencia])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe, col_factura="Código")
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("Fecha", "F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Bankinter_SIDI_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)
    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fill = green_fill if ws.cell(row=row[0].row, column=cobrada_col_idx).value == "SÍ" else red_fill
        for cell in row: cell.fill = fill
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value: cell.fill = amber_fill; cell.font = amber_font
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)
    wb.save(ruta_resultado)

    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    return {
        "total": len(df), "cobradas": len(cobradas), "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe": len(dup_importe), "avisos_factura": len(dup_factura),
        "ruta_resultado": ruta_resultado,
    }


def conciliar_abanca_sidi(ruta_banco: str, ruta_facturas: str) -> dict:
    """Cruza facturas SIDI (Cliente/SubCliente + Total) contra extracto Abanca."""
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=5)
    banco.columns = [c.strip() for c in banco.columns]

    col_tipo_op    = "TIPO OPERACIÓN"
    col_importe_b  = "IMPORTE"
    col_fecha      = "F. OPERACIÓN"
    col_referencia = "REFERENCIA"

    for col in (col_tipo_op, col_importe_b, col_fecha, col_referencia):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco Abanca: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_importe_b] = (
        banco[col_importe_b]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_b] = pd.to_numeric(banco[col_importe_b], errors="coerce")

    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Cliente/SubCliente"
    col_importe = "Total"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas SIDI: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    categorias_validas = {
        "transferencias de otras entidades",
        "transferencias propia entidad",
    }
    banco_disponible = banco[
        banco[col_tipo_op].fillna("").str.strip().str.lower().isin(categorias_validas) &
        (banco[col_importe_b] > 0)
    ].copy()

    cobrada_flags, fechas_cobro, referencias = [], [], []

    for _, fac in df.iterrows():
        coincidencia = banco_disponible[banco_disponible[col_importe_b] == fac[col_importe]]
        if not coincidencia.empty:
            idx = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_referencia])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe, col_factura="Código")
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("Fecha", "F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Abanca_SIDI_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)
    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fill = green_fill if ws.cell(row=row[0].row, column=cobrada_col_idx).value == "SÍ" else red_fill
        for cell in row: cell.fill = fill
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value: cell.fill = amber_fill; cell.font = amber_font
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)
    wb.save(ruta_resultado)

    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    return {
        "total": len(df), "cobradas": len(cobradas), "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe": len(dup_importe), "avisos_factura": len(dup_factura),
        "ruta_resultado": ruta_resultado,
    }


def conciliar_lacaixa_sidi(ruta_banco: str, ruta_facturas: str) -> dict:
    """Cruza facturas SIDI (Cliente/SubCliente + Total) contra extracto La Caixa."""
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=3, header=None)

    if banco.shape[1] < 5:
        raise ValueError(
            f"El extracto de La Caixa debe tener al menos 5 columnas. "
            f"Se encontraron {banco.shape[1]}."
        )

    COL_FECHA   = 1
    COL_DESC    = 3
    COL_IMPORTE = 4

    banco[COL_IMPORTE] = (
        banco[COL_IMPORTE]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[COL_IMPORTE] = pd.to_numeric(banco[COL_IMPORTE], errors="coerce")

    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Cliente/SubCliente"
    col_importe = "Total"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas SIDI: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)
    banco_disponible = banco[banco[COL_IMPORTE] > 0].copy()

    cobrada_flags, fechas_cobro, referencias = [], [], []

    for _, fac in df.iterrows():
        coincidencia = banco_disponible[banco_disponible[COL_IMPORTE] == fac[col_importe]]
        if not coincidencia.empty:
            idx = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[COL_FECHA])
            referencias.append(row_b[COL_DESC])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe, col_factura="Código")
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("Fecha", "F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_LaCaixa_SIDI_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)
    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fill = green_fill if ws.cell(row=row[0].row, column=cobrada_col_idx).value == "SÍ" else red_fill
        for cell in row: cell.fill = fill
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value: cell.fill = amber_fill; cell.font = amber_font
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)
    wb.save(ruta_resultado)

    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    return {
        "total": len(df), "cobradas": len(cobradas), "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe": len(dup_importe), "avisos_factura": len(dup_factura),
        "ruta_resultado": ruta_resultado,
    }


def conciliar_unicaja(ruta_banco: str, ruta_facturas: str) -> dict:
    """
    Cruza facturas (Nombre/Importe) contra extracto Unicaja.
    El extracto tiene 10 filas de cabecera; los datos empiezan en la fila 11.
    Columnas: Fecha de operación, Fecha valor, Concepto, Importe, Divisa, Saldo, Divisa, Nº mov, Oficina.
    Genera conciliacion_Unicaja_DD-MM-YYYY_HH-MM-SS.xlsx.
    """
    # ── Leer banco Unicaja (cabecera en fila 11) ──────────────────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=10)
    banco.columns = [c.strip() for c in banco.columns]

    col_fecha    = "Fecha de operación"
    col_importe_b = "Importe"
    col_nmov     = "Nº mov"

    for col in (col_fecha, col_importe_b, col_nmov):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco Unicaja: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_importe_b] = (
        banco[col_importe_b]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_b] = pd.to_numeric(banco[col_importe_b], errors="coerce")

    # ── Leer facturas (fila 1 = cabecera) ─────────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Nombre"
    col_importe = "Importe"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    # ── Solo facturas con importe positivo ────────────────────────────────────
    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    # ── Banco: solo movimientos con IMPORTE positivo ──────────────────────────
    banco_disponible = banco[banco[col_importe_b] > 0].copy()

    # ── Cruce por importe exacto ───────────────────────────────────────────────
    cobrada_flags = []
    fechas_cobro  = []
    referencias   = []

    for _, fac in df.iterrows():
        importe_fac = fac[col_importe]

        coincidencia = banco_disponible[
            banco_disponible[col_importe_b] == importe_fac
        ]

        if not coincidencia.empty:
            idx   = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_nmov])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe)
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    # ── Exportar Excel con formato ─────────────────────────────────────────────
    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Unicaja_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)

    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)

    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cobrada_val = ws.cell(row=row[0].row, column=cobrada_col_idx).value
        fill = green_fill if cobrada_val == "SÍ" else red_fill
        for cell in row:
            cell.fill = fill

    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill  = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font  = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value:
                cell.fill = amber_fill
                cell.font = amber_font

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    wb.save(ruta_resultado)

    # ── Estadísticas ───────────────────────────────────────────────────────────
    n_total    = len(df)
    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]

    return {
        "total":              n_total,
        "cobradas":           len(cobradas),
        "pendientes":         len(pendientes),
        "importe_cobradas":   cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe":     len(dup_importe),
        "avisos_factura":     len(dup_factura),
        "ruta_resultado":     ruta_resultado,
    }


def conciliar_bbva_sidi(ruta_banco: str, ruta_facturas: str) -> dict:
    """Cruza facturas SIDI (Cliente/SubCliente + Total) contra extracto BBVA."""
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=15)
    banco.columns = [c.strip() for c in banco.columns]

    col_fecha        = "F. CONTABLE"
    col_concepto     = "CONCEPTO"
    col_beneficiario = "BENEFICIARIO/ORDENANTE"
    col_importe_b    = "IMPORTE"

    for col in (col_fecha, col_concepto, col_beneficiario, col_importe_b):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco BBVA: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_importe_b] = (
        banco[col_importe_b]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_b] = pd.to_numeric(banco[col_importe_b], errors="coerce")

    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Cliente/SubCliente"
    col_importe = "Total"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas SIDI: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)

    banco_disponible = banco[
        banco[col_concepto].fillna("").str.strip().str.lower().str.contains("transferencia", regex=False) &
        (banco[col_importe_b] > 0)
    ].copy()

    cobrada_flags, fechas_cobro, referencias = [], [], []

    for _, fac in df.iterrows():
        coincidencia = banco_disponible[banco_disponible[col_importe_b] == fac[col_importe]]
        if not coincidencia.empty:
            idx = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_beneficiario])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe, col_factura="Código")
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("Fecha", "F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_BBVA_SIDI_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)
    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fill = green_fill if ws.cell(row=row[0].row, column=cobrada_col_idx).value == "SÍ" else red_fill
        for cell in row: cell.fill = fill
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value: cell.fill = amber_fill; cell.font = amber_font
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)
    wb.save(ruta_resultado)

    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    return {
        "total": len(df), "cobradas": len(cobradas), "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe": len(dup_importe), "avisos_factura": len(dup_factura),
        "ruta_resultado": ruta_resultado,
    }


def conciliar_unicaja_sidi(ruta_banco: str, ruta_facturas: str) -> dict:
    """Cruza facturas SIDI (Cliente/SubCliente + Total) contra extracto Unicaja."""
    # ── Leer banco Unicaja (cabecera en fila 11) ──────────────────────────────
    banco = pd.read_excel(ruta_banco, dtype=str, skiprows=10)
    banco.columns = [c.strip() for c in banco.columns]

    col_fecha     = "Fecha de operación"
    col_importe_b = "Importe"
    col_nmov      = "Nº mov"

    for col in (col_fecha, col_importe_b, col_nmov):
        if col not in banco.columns:
            raise ValueError(
                f"Columna no encontrada en el banco Unicaja: '{col}'\n"
                f"Columnas disponibles: {list(banco.columns)}"
            )

    banco[col_importe_b] = (
        banco[col_importe_b]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    banco[col_importe_b] = pd.to_numeric(banco[col_importe_b], errors="coerce")

    # ── Leer facturas SIDI (fila 1 = cabecera) ────────────────────────────────
    ext = os.path.splitext(ruta_facturas)[1].lower()
    if ext == ".xls":
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="xlrd")
    else:
        facturas = pd.read_excel(ruta_facturas, dtype=str, engine="openpyxl")
    facturas.columns = [c.strip() for c in facturas.columns]

    col_nombre  = "Cliente/SubCliente"
    col_importe = "Total"

    for col in (col_nombre, col_importe):
        if col not in facturas.columns:
            raise ValueError(
                f"Columna no encontrada en facturas SIDI: '{col}'\n"
                f"Columnas disponibles: {list(facturas.columns)}"
            )

    facturas[col_importe] = (
        facturas[col_importe]
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    facturas[col_importe] = pd.to_numeric(facturas[col_importe], errors="coerce")

    df = facturas[facturas[col_importe] > 0].copy().reset_index(drop=True)
    banco_disponible = banco[banco[col_importe_b] > 0].copy()

    cobrada_flags, fechas_cobro, referencias = [], [], []

    for _, fac in df.iterrows():
        coincidencia = banco_disponible[banco_disponible[col_importe_b] == fac[col_importe]]
        if not coincidencia.empty:
            idx = coincidencia.index[0]
            row_b = banco_disponible.loc[idx]
            cobrada_flags.append("SÍ")
            fechas_cobro.append(row_b[col_fecha])
            referencias.append(row_b[col_nmov])
            banco_disponible = banco_disponible.drop(index=idx)
        else:
            cobrada_flags.append("NO")
            fechas_cobro.append("")
            referencias.append("")

    df["COBRADA"]           = cobrada_flags
    df["FECHA COBRO"]       = fechas_cobro
    df["NRO. APUNTE BANCO"] = referencias

    df = _añadir_observaciones(df, col_importe, col_factura="Código")
    df["FECHA COBRO"] = df["FECHA COBRO"].apply(_formatear_fecha)
    for _col_fecha in ("Fecha", "F. Factura", "Fecha vto."):
        if _col_fecha in df.columns:
            df[_col_fecha] = df[_col_fecha].apply(_formatear_fecha)

    carpeta        = os.path.dirname(ruta_facturas)
    fecha_hoy      = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    ruta_resultado = os.path.join(carpeta, f"conciliacion_Unicaja_SIDI_{fecha_hoy}.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active
    _aplicar_formato_texto_fechas(ws, df)
    header_fill = PatternFill(fill_type="solid", fgColor=BLUE_DARK)
    header_font = Font(bold=True, color="FFFFFF")
    green_fill  = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
    red_fill    = PatternFill(fill_type="solid", fgColor=RED_FILL)
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    cobrada_col_idx = df.columns.get_loc("COBRADA") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fill = green_fill if ws.cell(row=row[0].row, column=cobrada_col_idx).value == "SÍ" else red_fill
        for cell in row: cell.fill = fill
    if "OBSERVACIONES" in df.columns:
        obs_col_idx = df.columns.get_loc("OBSERVACIONES") + 1
        amber_fill = PatternFill(fill_type="solid", fgColor=AMBER_FILL)
        amber_font = Font(bold=True, color="7D4800")
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=obs_col_idx)
            if cell.value: cell.fill = amber_fill; cell.font = amber_font
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)
    wb.save(ruta_resultado)

    cobradas   = df[df["COBRADA"] == "SÍ"]
    pendientes = df[df["COBRADA"] == "NO"]
    dup_importe = df[df["OBSERVACIONES"].str.contains("Importe", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    dup_factura = df[df["OBSERVACIONES"].str.contains("factura", na=False)] if "OBSERVACIONES" in df.columns else df.iloc[0:0]
    return {
        "total": len(df), "cobradas": len(cobradas), "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "avisos_importe": len(dup_importe), "avisos_factura": len(dup_factura),
        "ruta_resultado": ruta_resultado,
    }


# ── Interfaz gráfica ─────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conciliación Bancaria")
        self.geometry("580x680")
        self.resizable(True, True)
        self.configure(bg=BG)

        self._ruta_banco      = tk.StringVar()
        self._ruta_facturas   = tk.StringVar()
        self._modo            = tk.StringVar(value="bankinter")
        self._tipo_listado    = tk.StringVar(value="pancho")

        self._build_ui()

    # ── Construcción de widgets ───────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 20, "pady": 0}

        # ── Título ────────────────────────────────────────────────────────────
        tk.Label(
            self, text="Conciliación de Facturas", font=("Segoe UI", 14, "bold"),
            bg=BG, fg=f"#{BLUE_DARK}"
        ).pack(pady=(22, 4))

        tk.Label(
            self, text="Cruza facturas pendientes con el extracto bancario",
            font=("Segoe UI", 9), bg=BG, fg="#666"
        ).pack(pady=(0, 12))

        # ── Selector de formato ───────────────────────────────────────────────
        frame_modo = tk.Frame(self, bg=BG)
        frame_modo.pack(fill="x", padx=20, pady=(0, 14))
        tk.Label(frame_modo, text="Banco:", font=("Segoe UI", 9, "bold"),
                 bg=BG).pack(side="left", padx=(0, 10))
        for texto, valor in [
            ("Bankinter", "bankinter"),
            ("Abanca",    "abanca"),
            ("BBVA",      "bbva"),
            ("La Caixa",  "lacaixa"),
            ("Unicaja",   "unicaja"),
        ]:
            tk.Radiobutton(
                frame_modo, text=texto, variable=self._modo, value=valor,
                font=("Segoe UI", 9), bg=BG, activebackground=BG,
                command=self._actualizar_labels,
            ).pack(side="left", padx=(0, 12))

        # ── Selector tipo listado ──────────────────────────────────────────────
        frame_listado = tk.Frame(self, bg=BG)
        frame_listado.pack(fill="x", padx=20, pady=(0, 14))
        tk.Label(frame_listado, text="Listado:", font=("Segoe UI", 9, "bold"),
                 bg=BG).pack(side="left", padx=(0, 10))
        for texto, valor in [
            ("Pancho", "pancho"),
            ("SIDI",   "sidi"),
        ]:
            tk.Radiobutton(
                frame_listado, text=texto, variable=self._tipo_listado, value=valor,
                font=("Segoe UI", 9), bg=BG, activebackground=BG,
                command=self._actualizar_labels,
            ).pack(side="left", padx=(0, 12))

        # ── Selector banco ────────────────────────────────────────────────────
        self._lbl_banco = tk.StringVar(value="Extracto Bankinter  (*.xlsx)")
        self._selector_frame_var(
            labelvar=self._lbl_banco,
            var=self._ruta_banco,
        )

        # ── Selector facturas ─────────────────────────────────────────────────
        self._lbl_facturas = tk.StringVar(value="Facturas pendientes  (*.xls / .xlsx)")
        self._selector_frame_var(
            labelvar=self._lbl_facturas,
            var=self._ruta_facturas,
        )

        # ── Botón ejecutar ────────────────────────────────────────────────────
        self._btn_ejecutar = tk.Button(
            self,
            text="Ejecutar conciliación Bankinter · Pancho  →",
            font=("Segoe UI", 11, "bold"),
            bg=f"#{BLUE_DARK}", fg="white",
            activebackground="#16375a", activeforeground="white",
            relief="flat", cursor="hand2",
            padx=20, pady=10,
            command=self._lanzar,
        )
        self._btn_ejecutar.pack(pady=(24, 0))

        # ── Panel de resultados ───────────────────────────────────────────────
        frame_res = tk.LabelFrame(
            self, text="Resultados", font=("Segoe UI", 9, "bold"),
            bg=BG, fg="#444", padx=14, pady=10
        )
        frame_res.pack(fill="x", padx=20, pady=(20, 0))

        self._lbl_total = self._res_label(frame_res, "Total procesadas:", "—")
        self._lbl_cobradas = self._res_label(frame_res, "Cobradas:", "—", fg=f"#{GREEN_FILL[:6]}")
        self._lbl_pendientes = self._res_label(frame_res, "Pendientes:", "—", fg="#c0392b")
        self._lbl_imp_cob = self._res_label(frame_res, "Importe cobrado:", "—")
        self._lbl_imp_pen = self._res_label(frame_res, "Importe pendiente:", "—")
        self._lbl_avisos_importe = self._res_label(frame_res, "Importes duplicados:", "—", fg="#7D4800")
        self._lbl_avisos_factura = self._res_label(frame_res, "Nº facturas duplicados:", "—", fg="#7D4800")

        # ── Barra de progreso ─────────────────────────────────────────────────
        self._progress = ttk.Progressbar(self, mode="indeterminate", length=480)
        self._progress.pack(pady=(16, 0))

        # ── Estado ────────────────────────────────────────────────────────────
        self._lbl_estado = tk.Label(
            self, text="", font=("Segoe UI", 9), bg=BG, fg="#444", wraplength=480
        )
        self._lbl_estado.pack(pady=(8, 0))

    def _selector_frame_var(self, labelvar: tk.StringVar, var: tk.StringVar):
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="x", padx=20, pady=(0, 10))

        tk.Label(outer, textvariable=labelvar, font=("Segoe UI", 9, "bold"), bg=BG, anchor="w"
                 ).pack(fill="x")

        row = tk.Frame(outer, bg=BG)
        row.pack(fill="x")

        entry = tk.Entry(
            row, textvariable=var, font=("Segoe UI", 9),
            relief="solid", bd=1, bg="white", fg="#222",
        )
        entry.pack(side="left", fill="x", expand=True, ipady=5)

        tk.Button(
            row, text="Seleccionar",
            font=("Segoe UI", 9), relief="flat",
            bg="#dde3ea", activebackground="#c4cdd8",
            cursor="hand2", padx=10,
            command=lambda v=var: self._elegir_fichero(v),
        ).pack(side="left", padx=(6, 0), ipady=5)

    def _actualizar_labels(self):
        nombres = {
            "bankinter": "Bankinter",
            "abanca":    "Abanca",
            "bbva":      "BBVA",
            "lacaixa":   "La Caixa",
            "unicaja":   "Unicaja",
        }
        banco_nombre   = nombres.get(self._modo.get(), self._modo.get())
        listado_nombre = "SIDI" if self._tipo_listado.get() == "sidi" else "Pancho"
        self._lbl_banco.set(f"Extracto {banco_nombre}  (*.xlsx)")
        self._lbl_facturas.set(f"Facturas pendientes {listado_nombre}  (*.xls / .xlsx)")
        self._btn_ejecutar.config(
            text=f"Ejecutar conciliación {banco_nombre} · {listado_nombre}  →"
        )

    def _res_label(self, parent, texto: str, valor: str, fg: str = "#222"):
        frame = tk.Frame(parent, bg=BG)
        frame.pack(fill="x", pady=1)
        tk.Label(frame, text=texto, font=("Segoe UI", 9), bg=BG, fg="#555", width=22, anchor="w"
                 ).pack(side="left")
        lbl_val = tk.Label(frame, text=valor, font=("Segoe UI", 9, "bold"), bg=BG, fg=fg)
        lbl_val.pack(side="left")
        return lbl_val

    # ── Acciones ──────────────────────────────────────────────────────────────

    def _elegir_fichero(self, var: tk.StringVar):
        ruta = filedialog.askopenfilename(
            filetypes=[
                ("Ficheros Excel", "*.xlsx *.xls *.XLSX *.XLS"),
                ("Excel 2007+", "*.xlsx *.XLSX"),
                ("Excel 97-2003", "*.xls *.XLS"),
                ("Todos los ficheros", "*.*"),
            ]
        )
        if ruta:
            var.set(ruta)

    def _lanzar(self):
        banco = self._ruta_banco.get().strip()
        facturas = self._ruta_facturas.get().strip()

        if not banco:
            messagebox.showwarning("Falta fichero", "Selecciona el extracto del banco.")
            return
        if not facturas:
            messagebox.showwarning("Falta fichero", "Selecciona el fichero de facturas.")
            return

        self._btn_ejecutar.config(state="disabled")
        self._lbl_estado.config(text="Procesando…", fg="#444")
        self._progress.start(12)

        threading.Thread(
            target=self._ejecutar_hilo,
            args=(banco, facturas, self._modo.get(), self._tipo_listado.get()),
            daemon=True,
        ).start()

    def _ejecutar_hilo(self, banco: str, facturas: str, modo: str, tipo_listado: str):
        try:
            if tipo_listado == "sidi":
                if modo == "bankinter":
                    stats = conciliar_bankinter_sidi(banco, facturas)
                elif modo == "abanca":
                    stats = conciliar_abanca_sidi(banco, facturas)
                elif modo == "lacaixa":
                    stats = conciliar_lacaixa_sidi(banco, facturas)
                elif modo == "bbva":
                    stats = conciliar_bbva_sidi(banco, facturas)
                elif modo == "unicaja":
                    stats = conciliar_unicaja_sidi(banco, facturas)
                else:
                    stats = conciliar_bankinter_sidi(banco, facturas)
            else:
                if modo == "bankinter":
                    stats = conciliar_bankinter(banco, facturas)
                elif modo == "abanca":
                    stats = conciliar_abanca(banco, facturas)
                elif modo == "lacaixa":
                    stats = conciliar_lacaixa(banco, facturas)
                elif modo == "bbva":
                    stats = conciliar_bbva(banco, facturas)
                elif modo == "unicaja":
                    stats = conciliar_unicaja(banco, facturas)
                else:
                    stats = conciliar(banco, facturas)
            self.after(0, self._mostrar_resultado, stats)
        except Exception as exc:
            self.after(0, self._mostrar_error, str(exc))

    def _mostrar_resultado(self, stats: dict):
        self._progress.stop()
        self._btn_ejecutar.config(state="normal")

        self._lbl_total.config(text=str(stats["total"]))
        self._lbl_cobradas.config(text=str(stats["cobradas"]))
        self._lbl_pendientes.config(text=str(stats["pendientes"]))
        self._lbl_imp_cob.config(text=f"{stats['importe_cobradas']:,.2f} €")
        self._lbl_imp_pen.config(text=f"{stats['importe_pendientes']:,.2f} €")
        av_imp = stats.get("avisos_importe", 0)
        av_fac = stats.get("avisos_factura", 0)
        self._lbl_avisos_importe.config(text=str(av_imp) if av_imp == 0 else f"{av_imp}  ⚠")
        self._lbl_avisos_factura.config(text=str(av_fac) if av_fac == 0 else f"{av_fac}  ⚠")

        self._lbl_estado.config(
            text=f"✔ Resultado guardado en: {stats['ruta_resultado']}",
            fg="#27ae60",
        )

        if messagebox.askyesno(
            "Conciliación completada",
            "El fichero se ha generado correctamente.\n¿Desea abrirlo ahora?",
        ):
            self._abrir_fichero(stats["ruta_resultado"])

    def _mostrar_error(self, mensaje: str):
        self._progress.stop()
        self._btn_ejecutar.config(state="normal")
        self._lbl_estado.config(text=f"Error: {mensaje}", fg="#c0392b")
        messagebox.showerror("Error en la conciliación", mensaje)

    @staticmethod
    def _abrir_fichero(ruta: str):
        try:
            if sys.platform.startswith("win"):
                os.startfile(ruta)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.run(["open", ruta], check=True)
            else:
                subprocess.run(["xdg-open", ruta], check=True)
        except Exception as exc:
            messagebox.showerror("No se pudo abrir", str(exc))


# ── Punto de entrada ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
