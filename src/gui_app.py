import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import pandas as pd

from src.workflow_service import HorasExtrasWorkflowService


class HorasExtrasGUI:
    SELECTION_COLUMN = "SELECCION"
    INTERNAL_SELECTION_COLUMN = "__SELECTED__"
    TABLE_COLUMNS = HorasExtrasWorkflowService.TABLE_COLUMNS
    DISPLAY_COLUMNS = [column for column in TABLE_COLUMNS if column != "ID"]

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
        self.temporal_loaded = True

        self.workflow = HorasExtrasWorkflowService()

        self._build_ui()
        self.load_historico_with_current_filters()

    def _build_ui(self):
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill="x")

        self.path_var = tk.StringVar(value="Sin archivo seleccionado")
        ttk.Label(top_frame, textvariable=self.path_var).grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 8))

        ttk.Button(top_frame, text="Seleccionar Excel", command=self.select_excel).grid(row=1, column=0, padx=4, pady=2)
        ttk.Button(top_frame, text="Confirmar marcadas", command=self.confirm_loaded).grid(row=1, column=1, padx=4, pady=2)
        ttk.Button(top_frame, text="Descartar marcadas", command=self.discard_selected).grid(row=1, column=2, padx=4, pady=2)

        ttk.Label(top_frame, text="ROW_STATUS").grid(row=1, column=3, padx=(12, 4), sticky="e")
        self.row_status_var = tk.StringVar(value="")
        self.row_status_combo = ttk.Combobox(
            top_frame,
            textvariable=self.row_status_var,
            values=["", "NO_CONFIRMADO", "CONFIRMADO", "ELIMINADO"],
            width=16,
            state="readonly",
        )
        self.row_status_combo.grid(row=1, column=4, padx=4, pady=2)

        ttk.Label(top_frame, text="Nombre").grid(row=1, column=5, padx=(12, 4), sticky="e")
        self.nombre_var = tk.StringVar(value="")
        self.nombre_entry = ttk.Entry(top_frame, textvariable=self.nombre_var, width=22)
        self.nombre_entry.grid(row=1, column=6, padx=4, pady=2)

        ttk.Label(top_frame, text="Desde").grid(row=2, column=3, padx=(12, 4), sticky="e")
        self.fecha_desde_var = tk.StringVar(value="")
        self.fecha_desde_entry = ttk.Entry(top_frame, textvariable=self.fecha_desde_var, width=16)
        self.fecha_desde_entry.grid(row=2, column=4, padx=4, pady=2)

        ttk.Label(top_frame, text="Hasta").grid(row=2, column=5, padx=(12, 4), sticky="e")
        self.fecha_hasta_var = tk.StringVar(value="")
        self.fecha_hasta_entry = ttk.Entry(top_frame, textvariable=self.fecha_hasta_var, width=16)
        self.fecha_hasta_entry.grid(row=2, column=6, padx=4, pady=2)

        ttk.Button(top_frame, text="Aplicar filtros", command=self.apply_filters).grid(row=2, column=7, padx=4, pady=2)
        ttk.Button(top_frame, text="Limpiar filtros", command=self.clear_filters).grid(row=2, column=8, padx=4, pady=2)

        top_frame.columnconfigure(0, weight=1)

        table_frame = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        table_frame.pack(fill="both", expand=True)

        tree_columns = [self.SELECTION_COLUMN] + self.DISPLAY_COLUMNS
        self.tree = ttk.Treeview(table_frame, columns=tree_columns, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        for col in tree_columns:
            self.tree.heading(col, text=col)
            width = 170
            if col == self.SELECTION_COLUMN:
                width = 90
            elif col in {"ID", "ROW_STATUS"}:
                width = 130
            if col == "COMENTARIOS":
                width = 240
            self.tree.column(col, width=width, anchor="w")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<Button-1>", self.on_single_click)
        self.tree.bind("<Double-1>", self.on_cell_double_click)

    def ensure_selection_column(self):
        if self.INTERNAL_SELECTION_COLUMN not in self.preview_df.columns:
            self.preview_df[self.INTERNAL_SELECTION_COLUMN] = False

    def load_historico_with_current_filters(self):
        self.preview_df = self.workflow.get_historico_filtered(
            row_status=self.row_status_var.get(),
            nombre_filtro=self.nombre_var.get(),
            fecha_desde=self.fecha_desde_var.get(),
            fecha_hasta=self.fecha_hasta_var.get(),
        )
        self.ensure_selection_column()
        self.temporal_loaded = True
        self.refresh_table()

    def apply_filters(self):
        try:
            self.load_historico_with_current_filters()
        except ValueError as exc:
            messagebox.showerror("Filtro invalido", str(exc))

    def clear_filters(self):
        self.row_status_var.set("")
        self.nombre_var.set("")
        self.fecha_desde_var.set("")
        self.fecha_hasta_var.set("")
        self.apply_filters()

    def reset_filters_for_no_confirmado(self):
        self.row_status_var.set("NO_CONFIRMADO")
        self.nombre_var.set("")
        self.fecha_desde_var.set("")
        self.fecha_hasta_var.set("")

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
            loaded_df = self.workflow.build_preview_from_excel(path)
            loaded_df, _ = self.workflow.load_temporal(loaded_df, False)

            self.reset_filters_for_no_confirmado()
            self.load_historico_with_current_filters()

            messagebox.showinfo(
                "Ok",
                "Archivo procesado y cargado temporalmente. Se aplico el filtro NO_CONFIRMADO.",
            )
        except ValueError as exc:
            messagebox.showerror("Configuracion incompleta", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {exc}")

    def refresh_table(self):
        self.ensure_selection_column()

        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, row in self.preview_df.iterrows():
            marcado = bool(row.get(self.INTERNAL_SELECTION_COLUMN, False))
            checkbox = "[x]" if marcado else "[ ]"
            values = [checkbox] + [self._display_value(row.get(col, "")) for col in self.DISPLAY_COLUMNS]
            self.tree.insert("", "end", iid=str(idx), values=values)

    def on_single_click(self, event):
        col_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)

        if col_id == "#1" and row_id:
            row_index = int(row_id)
            current = bool(self.preview_df.at[row_index, self.INTERNAL_SELECTION_COLUMN])
            self.preview_df.at[row_index, self.INTERNAL_SELECTION_COLUMN] = not current
            self.refresh_table()
            return "break"

        return None

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

        col_index = int(col_id.replace("#", "")) - 2
        if col_index < 0 or col_index >= len(self.DISPLAY_COLUMNS):
            return

        column_name = self.DISPLAY_COLUMNS[col_index]
        if column_name not in self.EDITABLE_COLUMNS:
            messagebox.showwarning("No editable", f"La columna {column_name} no se puede editar.")
            return

        row_index = int(selected)
        current_value = self._display_value(self.preview_df.iloc[row_index][column_name])
        new_value = simpledialog.askstring("Editar", f"Nuevo valor para {column_name}:", initialvalue=current_value)

        if new_value is None:
            return

        try:
            parsed = self.workflow.parse_edition_value(column_name, new_value)
            self.preview_df.at[row_index, column_name] = parsed
            self.refresh_table()

            if self.temporal_loaded:
                self.preview_df, self.temporal_loaded = self.workflow.load_temporal(self.preview_df, True)
        except ValueError:
            messagebox.showerror("Valor invalido", "Para columnas numericas debes ingresar un numero valido.")

    def confirm_loaded(self):
        if self.preview_df.empty:
            messagebox.showwarning("Sin datos", "No hay datos para confirmar.")
            return

        try:
            selected_ids = self.preview_df.loc[
                self.preview_df[self.INTERNAL_SELECTION_COLUMN],
                "ID",
            ].tolist()

            self.preview_df, self.temporal_loaded = self.workflow.confirm_selected(
                self.preview_df,
                selected_ids,
                self.temporal_loaded,
            )

            self.refresh_table()
            messagebox.showinfo("Confirmado", "Filas marcadas confirmadas en historico.")
        except ValueError as exc:
            messagebox.showwarning("Sin filas marcadas", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo confirmar la carga: {exc}")

    def discard_selected(self):
        if self.preview_df.empty:
            messagebox.showwarning("Sin datos", "No hay datos para descartar.")
            return

        try:
            selected_ids = self.preview_df.loc[
                self.preview_df[self.INTERNAL_SELECTION_COLUMN],
                "ID",
            ].tolist()

            self.preview_df, self.temporal_loaded = self.workflow.discard_selected(
                self.preview_df,
                selected_ids,
                self.temporal_loaded,
            )

            self.preview_df.loc[self.preview_df["ID"].isin(selected_ids), self.INTERNAL_SELECTION_COLUMN] = False
            self.refresh_table()
            messagebox.showinfo("Descartado", "Filas marcadas descartadas en historico.")
        except ValueError as exc:
            messagebox.showwarning("Sin filas marcadas", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo descartar la carga: {exc}")

    def run(self):
        self.root.mainloop()
