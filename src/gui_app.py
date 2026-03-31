import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import pandas as pd

from src.Qontact_report_reader import ReporteHorasExtras
from src.controlador_historico import ControladorHistorico
from src.separador_de_jornales import SeparadorDeJornales


class HorasExtrasGUI:
    TABLE_COLUMNS = [
        "ID",
        "ROW_STATUS",
        "NOMBRE_Y_APELLIDO",
        "INGRESO",
        "EGRESO",
        "COMENTARIOS",
        "HORAS_TRABAJADAS",
        "HORAS_NORMALES_DIURNAS",
        "HORAS_EXTRAS_NORMALES",
        "HORAS_NOCTURNAS",
    ]

    EDITABLE_COLUMNS = {
        "COMENTARIOS",
        "HORAS_TRABAJADAS",
        "HORAS_NORMALES_DIURNAS",
        "HORAS_EXTRAS_NORMALES",
        "HORAS_NOCTURNAS",
    }

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Sistema de pago de horas extras")
        self.root.geometry("1300x650")

        self.selected_excel_path = ""
        self.preview_df = pd.DataFrame(columns=self.TABLE_COLUMNS)
        self.temporal_loaded = False

        self.historico = ControladorHistorico()

        self._build_ui()

    def _build_ui(self):
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill="x")

        self.path_var = tk.StringVar(value="Sin archivo seleccionado")
        ttk.Label(top_frame, textvariable=self.path_var).pack(side="left", fill="x", expand=True)

        ttk.Button(top_frame, text="Seleccionar Excel", command=self.select_excel).pack(side="left", padx=4)
        ttk.Button(top_frame, text="Cargar temporal en historico", command=self.load_temporal).pack(side="left", padx=4)
        ttk.Button(top_frame, text="Confirmar cargado", command=self.confirm_loaded).pack(side="left", padx=4)

        table_frame = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        table_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(table_frame, columns=self.TABLE_COLUMNS, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        for col in self.TABLE_COLUMNS:
            self.tree.heading(col, text=col)
            width = 170
            if col in {"ID", "ROW_STATUS"}:
                width = 130
            if col == "COMENTARIOS":
                width = 240
            self.tree.column(col, width=width, anchor="w")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<Double-1>", self.on_cell_double_click)

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar reporte de Qontact",
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        if not path:
            return

        self.selected_excel_path = path
        self.path_var.set(path)

        try:
            reader = ReporteHorasExtras(path)
            reporte_df = reader.read()

            separador = SeparadorDeJornales(reporte_df)
            result_df = separador.build_result_df()

            faltantes = result_df[result_df["HS_JORNAL"].isna()]["NOMBRE_Y_APELLIDO"].unique().tolist()
            if faltantes:
                empleados = "\n".join(f"- {name}" for name in faltantes)
                messagebox.showerror(
                    "Configuracion incompleta",
                    "Hay empleados sin HS_JORNAL configurado:\n\n" + empleados,
                )
                return

            self.preview_df = result_df.copy()
            self.preview_df["ID"] = ""
            self.preview_df["ROW_STATUS"] = "NO_CONFIRMADO"
            if "COMENTARIOS" not in self.preview_df.columns:
                self.preview_df["COMENTARIOS"] = ""

            self.preview_df = self.preview_df[self.TABLE_COLUMNS]
            self.temporal_loaded = False

            self.refresh_table()
            messagebox.showinfo("Ok", "Archivo procesado. Ya podes revisar y editar antes de confirmar.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {exc}")

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, row in self.preview_df.iterrows():
            values = [self._display_value(row.get(col, "")) for col in self.TABLE_COLUMNS]
            self.tree.insert("", "end", iid=str(idx), values=values)

    def _display_value(self, value):
        if pd.isna(value):
            return ""
        if isinstance(value, pd.Timestamp):
            return value.strftime("%d/%m/%Y %H:%M:%S")
        if isinstance(value, float):
            return f"{value:.2f}"
        return str(value)

    def on_cell_double_click(self, _event):
        selected = self.tree.focus()
        if not selected:
            return

        col_id = self.tree.identify_column(self.tree.winfo_pointerx() - self.tree.winfo_rootx())
        if not col_id:
            return

        col_index = int(col_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.TABLE_COLUMNS):
            return

        column_name = self.TABLE_COLUMNS[col_index]
        if column_name not in self.EDITABLE_COLUMNS:
            messagebox.showwarning("No editable", f"La columna {column_name} no se puede editar.")
            return

        row_index = int(selected)
        current_value = self._display_value(self.preview_df.iloc[row_index][column_name])
        new_value = simpledialog.askstring("Editar", f"Nuevo valor para {column_name}:", initialvalue=current_value)

        if new_value is None:
            return

        try:
            if column_name == "COMENTARIOS":
                parsed = new_value.strip()
            else:
                parsed = float(new_value)
            self.preview_df.at[row_index, column_name] = parsed
            self.refresh_table()

            if self.temporal_loaded:
                self.historico.update_records(self.preview_df)
        except ValueError:
            messagebox.showerror("Valor invalido", "Para columnas numericas debes ingresar un numero valido.")

    def load_temporal(self):
        if self.preview_df.empty:
            messagebox.showwarning("Sin datos", "Primero selecciona y procesa un archivo Excel.")
            return

        try:
            if not self.temporal_loaded:
                ids = self.historico.add_temporal_records(self.preview_df)
                self.preview_df["ID"] = ids
                self.preview_df["ROW_STATUS"] = "NO_CONFIRMADO"
                self.temporal_loaded = True
            else:
                self.historico.update_records(self.preview_df)

            self.refresh_table()
            messagebox.showinfo("Ok", "Carga temporal guardada en historico.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar temporalmente: {exc}")

    def confirm_loaded(self):
        if self.preview_df.empty:
            messagebox.showwarning("Sin datos", "No hay datos para confirmar.")
            return

        try:
            if not self.temporal_loaded:
                ids = self.historico.add_temporal_records(self.preview_df)
                self.preview_df["ID"] = ids
                self.temporal_loaded = True

            self.historico.update_records(self.preview_df)
            self.historico.confirm_records(self.preview_df["ID"].tolist())
            self.preview_df["ROW_STATUS"] = "CONFIRMADO"

            self.refresh_table()
            messagebox.showinfo("Confirmado", "Registros confirmados en historico.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo confirmar la carga: {exc}")

    def run(self):
        self.root.mainloop()
