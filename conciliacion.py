"""
Conciliación de facturas pendientes contra movimientos bancarios.
Dependencias: pandas, openpyxl, xlrd
"""

import threading
import subprocess
import sys
import os
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
RED_FILL = "FFC7CE"


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

    # ── Exportar Excel con formato ────────────────────────────────────────────
    carpeta = os.path.dirname(ruta_facturas)
    ruta_resultado = os.path.join(carpeta, "resultado_conciliacion.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active

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

    return {
        "total": n_total,
        "cobradas": len(cobradas),
        "pendientes": len(pendientes),
        "importe_cobradas": cobradas[col_importe_fac].sum(),
        "importe_pendientes": pendientes[col_importe_fac].sum(),
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

    # ── Banco: filtrar Transferencias / Recibos con HABER positivo ────────────
    categorias_validas = {"transferencias", "recibos"}
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
        nombre_norm = _normalizar(fac[col_nombre])

        coincidencia = banco_disponible[
            (banco_disponible[col_haber] == importe_fac) &
            banco_disponible[col_descripcion].apply(
                lambda x: nombre_norm in str(x).strip().lower()
            )
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

    # ── Exportar Excel con formato ────────────────────────────────────────────
    carpeta        = os.path.dirname(ruta_facturas)
    ruta_resultado = os.path.join(carpeta, "resultado_conciliacion.xlsx")

    df = df.map(_limpiar)
    df.to_excel(ruta_resultado, index=False, engine="openpyxl")

    wb = load_workbook(ruta_resultado)
    ws = wb.active

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

    return {
        "total":              n_total,
        "cobradas":           len(cobradas),
        "pendientes":         len(pendientes),
        "importe_cobradas":   cobradas[col_importe].sum(),
        "importe_pendientes": pendientes[col_importe].sum(),
        "ruta_resultado":     ruta_resultado,
    }


# ── Interfaz gráfica ─────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conciliación Bancaria")
        self.geometry("520x620")
        self.resizable(False, False)
        self.configure(bg=BG)

        self._ruta_banco    = tk.StringVar()
        self._ruta_facturas = tk.StringVar()
        self._modo          = tk.StringVar(value="bankinter")

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
        ]:
            tk.Radiobutton(
                frame_modo, text=texto, variable=self._modo, value=valor,
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
            text="Ejecutar conciliación  →",
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
        }
        banco_nombre = nombres.get(self._modo.get(), self._modo.get())
        self._lbl_banco.set(f"Extracto {banco_nombre}  (*.xlsx)")
        self._lbl_facturas.set("Facturas pendientes  (*.xls / .xlsx)")

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
            args=(banco, facturas, self._modo.get()),
            daemon=True,
        ).start()

    def _ejecutar_hilo(self, banco: str, facturas: str, modo: str):
        try:
            if modo == "bankinter":
                stats = conciliar_bankinter(banco, facturas)
            elif modo in ("abanca", "bbva", "lacaixa"):
                nombres = {"abanca": "Abanca", "bbva": "BBVA", "lacaixa": "La Caixa"}
                raise NotImplementedError(
                    f"El formato {nombres[modo]} aún no está implementado."
                )
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
