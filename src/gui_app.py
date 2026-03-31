import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import calendar
import datetime

import pandas as pd

from src.datos_empleados_reader import DatosEmpleados
from src.feriados import FeriadosReader
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
        self.root.title("Sistema de Pago de Horas Extras")
        self.root.geometry("1450x820")
        self.root.minsize(1200, 720)
        self.root.configure(bg="#eef2f7")

        self.workflow = HorasExtrasWorkflowService()
        self.datos_empleados = DatosEmpleados()
        self.feriados_reader = FeriadosReader()

        self.selected_excel_path = ""
        self.preview_df = pd.DataFrame(columns=self.TABLE_COLUMNS)
        self.temporal_loaded = True
        self.selected_empleado_nombre = ""
        self.selected_feriado_fecha = None

        self.current_view = "validacion"

        self._configure_style()
        self._build_layout()
        self._show_view("validacion")
        self.load_historico_with_current_filters()

    # =========================
    # ESTILOS
    # =========================
    def _configure_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "App.TFrame",
            background="#eef2f7"
        )

        style.configure(
            "Sidebar.TFrame",
            background="#dfe5ec"
        )

        style.configure(
            "Content.TFrame",
            background="#eef2f7"
        )

        style.configure(
            "Card.TFrame",
            background="#ffffff",
            relief="flat"
        )

        style.configure(
            "Header.TFrame",
            background="#3e4f68"
        )

        style.configure(
            "Header.TLabel",
            background="#3e4f68",
            foreground="white",
            font=("Segoe UI", 18, "bold")
        )

        style.configure(
            "Sidebar.TButton",
            font=("Segoe UI", 12),
            padding=12,
            anchor="w",
            background="#dfe5ec",
            foreground="#2e3a4d",
            borderwidth=0
        )
        style.map(
            "Sidebar.TButton",
            background=[("active", "#cfd8e3"), ("pressed", "#cfd8e3")]
        )

        style.configure(
            "SidebarActive.TButton",
            font=("Segoe UI", 12, "bold"),
            padding=12,
            anchor="w",
            background="#4a90d9",
            foreground="white",
            borderwidth=0
        )
        style.map(
            "SidebarActive.TButton",
            background=[("active", "#4a90d9"), ("pressed", "#4a90d9")]
        )

        style.configure(
            "CardTitle.TLabel",
            background="#ffffff",
            foreground="#344054",
            font=("Segoe UI", 15, "bold")
        )

        style.configure(
            "SectionLabel.TLabel",
            background="#ffffff",
            foreground="#475467",
            font=("Segoe UI", 11)
        )

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(14, 10),
            background="#2f80ed",
            foreground="white",
            borderwidth=0
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#2467c8")]
        )

        style.configure(
            "Success.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(14, 10),
            background="#2e9b50",
            foreground="white",
            borderwidth=0
        )
        style.map(
            "Success.TButton",
            background=[("active", "#257e41")]
        )

        style.configure(
            "Danger.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(14, 10),
            background="#c0392b",
            foreground="white",
            borderwidth=0
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#a93226")]
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            padding=(10, 8),
            background="#e7edf4",
            foreground="#344054",
            borderwidth=1
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#dbe4ee")]
        )

        style.configure(
            "CalendarIcon.TButton",
            font=("Segoe UI Emoji", 9),
            padding=(3, 1),
            background="#e7edf4",
            foreground="#344054",
            borderwidth=1,
            width=2,
        )
        style.map(
            "CalendarIcon.TButton",
            background=[("active", "#dbe4ee")]
        )

        style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=30,
            borderwidth=0,
            relief="flat"
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#f3f5f7",
            foreground="#344054",
            relief="flat"
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#e9edf2")]
        )

    # =========================
    # LAYOUT GENERAL
    # =========================
    def _build_layout(self):
        self.main_frame = ttk.Frame(self.root, style="App.TFrame")
        self.main_frame.pack(fill="both", expand=True)

        self._build_header()
        self._build_body()

    def _build_header(self):
        self.header_frame = ttk.Frame(self.main_frame, style="Header.TFrame", height=58)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)

        ttk.Label(
            self.header_frame,
            text="Sistema de Pago de Horas Extras",
            style="Header.TLabel"
        ).pack(side="left", padx=20, pady=12)

    def _build_body(self):
        self.body_frame = ttk.Frame(self.main_frame, style="App.TFrame")
        self.body_frame.pack(fill="both", expand=True)

        self._build_sidebar()
        self._build_content_area()

    def _build_sidebar(self):
        self.sidebar = ttk.Frame(self.body_frame, style="Sidebar.TFrame", width=230)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.nav_buttons = {}

        nav_items = [
            ("empleados", "Empleados"),
            ("feriados", "Feriados"),
            ("validacion", "Validación"),
            ("reportes", "Reportes"),
        ]

        for key, text in nav_items:
            btn = ttk.Button(
                self.sidebar,
                text=f"   {text}",
                style="Sidebar.TButton",
                command=lambda k=key: self._show_view(k)
            )
            btn.pack(fill="x", padx=0, pady=1)
            self.nav_buttons[key] = btn

    def _build_content_area(self):
        self.content = ttk.Frame(self.body_frame, style="Content.TFrame")
        self.content.pack(side="left", fill="both", expand=True, padx=18, pady=18)

        self.views = {}

        self.views["empleados"] = ttk.Frame(self.content, style="Content.TFrame")
        self.views["feriados"] = ttk.Frame(self.content, style="Content.TFrame")
        self.views["validacion"] = ttk.Frame(self.content, style="Content.TFrame")
        self.views["reportes"] = ttk.Frame(self.content, style="Content.TFrame")

        self._build_empleados_view()
        self._build_feriados_view()
        self._build_validacion_view()
        self._build_reportes_view()

    def _show_view(self, view_name):
        self.current_view = view_name

        for key, frame in self.views.items():
            frame.pack_forget()
            self.nav_buttons[key].configure(style="Sidebar.TButton")

        self.views[view_name].pack(fill="both", expand=True)
        self.nav_buttons[view_name].configure(style="SidebarActive.TButton")

        if view_name == "empleados":
            self.refresh_empleados_table()
        elif view_name == "feriados":
            self.refresh_feriados_table()

    # =========================
    # HELPERS VISUALES
    # =========================
    def _create_card(self, parent, title):
        outer = ttk.Frame(parent, style="Card.TFrame")
        outer.pack(fill="x", pady=(0, 16))

        title_label = ttk.Label(outer, text=title, style="CardTitle.TLabel")
        title_label.pack(anchor="w", padx=18, pady=(16, 8))

        separator = ttk.Separator(outer, orient="horizontal")
        separator.pack(fill="x", padx=18, pady=(0, 12))

        inner = ttk.Frame(outer, style="Card.TFrame")
        inner.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        return outer, inner

    # =========================
    # VISTA EMPLEADOS
    # =========================
    def _build_empleados_view(self):
        parent = self.views["empleados"]

        _, card = self._create_card(parent, "Configuración de empleados")

        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x", pady=(0, 12))

        self.emp_nombre_var = tk.StringVar()
        self.emp_valor_hs_jornal_var = tk.StringVar()
        self.emp_hs_jornal_var = tk.StringVar()
        self.emp_valor_hs_extras_var = tk.StringVar()

        ttk.Label(form, text="Nombre", style="SectionLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.emp_nombre_var, width=36).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(form, text="Valor hs jornal", style="SectionLabel.TLabel").grid(row=0, column=2, sticky="w", padx=(18, 8), pady=4)
        ttk.Entry(form, textvariable=self.emp_valor_hs_jornal_var, width=16).grid(row=0, column=3, sticky="w", pady=4)

        ttk.Label(form, text="Hs jornal", style="SectionLabel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.emp_hs_jornal_var, width=16).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(form, text="Valor hs extras", style="SectionLabel.TLabel").grid(row=1, column=2, sticky="w", padx=(18, 8), pady=4)
        ttk.Entry(form, textvariable=self.emp_valor_hs_extras_var, width=16).grid(row=1, column=3, sticky="w", pady=4)

        actions = ttk.Frame(form, style="Card.TFrame")
        actions.grid(row=0, column=4, rowspan=2, padx=(26, 0), sticky="e")

        ttk.Button(actions, text="Agregar", style="Primary.TButton", command=self.add_empleado).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Actualizar", style="Success.TButton", command=self.update_empleado).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Eliminar", style="Danger.TButton", command=self.remove_empleado).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Limpiar", style="Secondary.TButton", command=self.clear_empleado_form).pack(side="left")

        columns = ("NOMBRE_Y_APELLIDO", "VALOR_HS_JORNAL", "HS_JORNAL", "VALOR_HS_EXTRAS")
        self.empleados_tree = ttk.Treeview(card, columns=columns, show="headings", height=12)
        self.empleados_tree.pack(fill="x")
        self.empleados_tree.bind("<<TreeviewSelect>>", self.on_empleado_selected)

        widths = {
            "NOMBRE_Y_APELLIDO": 320,
            "VALOR_HS_JORNAL": 180,
            "HS_JORNAL": 140,
            "VALOR_HS_EXTRAS": 180,
        }

        for col in columns:
            self.empleados_tree.heading(col, text=col.replace("_", " "))
            self.empleados_tree.column(col, width=widths[col], anchor="w")

        self.refresh_empleados_table()

    # =========================
    # VISTA FERIADOS
    # =========================
    def _build_feriados_view(self):
        parent = self.views["feriados"]

        _, card = self._create_card(parent, "Carga de feriados")
        ttk.Label(
            card,
            text="Acá podés agregar, editar o eliminar feriados que luego serán considerados en el cálculo.",
            style="SectionLabel.TLabel"
        ).pack(anchor="w", pady=(0, 12))

        form = ttk.Frame(card, style="Card.TFrame")
        form.pack(fill="x")

        ttk.Label(form, text="Fecha:", style="SectionLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.feriado_fecha_var = tk.StringVar()
        feriado_fecha_frame = ttk.Frame(form, style="Card.TFrame")
        feriado_fecha_frame.grid(row=0, column=1, sticky="w", pady=4)
        ttk.Entry(feriado_fecha_frame, textvariable=self.feriado_fecha_var, width=15, state="readonly").pack(side="left")
        ttk.Button(
            feriado_fecha_frame,
            text="📅",
            style="CalendarIcon.TButton",
            command=lambda: self._open_date_picker(self.feriado_fecha_var),
        ).pack(side="left", padx=(6, 0))

        ttk.Label(form, text="Descripción:", style="SectionLabel.TLabel").grid(row=0, column=2, sticky="w", padx=(20, 8), pady=4)
        self.feriado_desc_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.feriado_desc_var, width=35).grid(row=0, column=3, sticky="w", pady=4)

        actions = ttk.Frame(form, style="Card.TFrame")
        actions.grid(row=0, column=4, padx=(20, 0), pady=4, sticky="w")

        ttk.Button(actions, text="Agregar", style="Primary.TButton", command=self.add_feriado).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Actualizar", style="Success.TButton", command=self.update_feriado).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Eliminar", style="Danger.TButton", command=self.remove_feriado).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Limpiar", style="Secondary.TButton", command=self.clear_feriado_form).pack(side="left")

        self.feriados_table = ttk.Treeview(card, columns=("FECHA_FERIADO", "DESCRIPCION_FERIADO"), show="headings", height=10)
        self.feriados_table.pack(fill="x", pady=(16, 0))
        self.feriados_table.heading("FECHA_FERIADO", text="Fecha")
        self.feriados_table.heading("DESCRIPCION_FERIADO", text="Descripción")
        self.feriados_table.column("FECHA_FERIADO", width=170, anchor="w")
        self.feriados_table.column("DESCRIPCION_FERIADO", width=600, anchor="w")
        self.feriados_table.bind("<<TreeviewSelect>>", self.on_feriado_selected)

        self.refresh_feriados_table()

    def _parse_float(self, raw_value: str, field_name: str) -> float:
        value = (raw_value or "").strip().replace(",", ".")
        if not value:
            raise ValueError(f"El campo '{field_name}' es obligatorio.")
        return float(value)

    def _parse_feriado_date(self, raw_value: str) -> pd.Timestamp:
        value = (raw_value or "").strip()
        if not value:
            raise ValueError("La fecha es obligatoria.")

        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            parsed = pd.to_datetime(value, format=fmt, errors="coerce")
            if not pd.isna(parsed):
                return parsed.normalize()

        raise ValueError("Formato de fecha invalido. Usa dd/mm/yyyy o yyyy-mm-dd.")

    def refresh_empleados_table(self):
        if not hasattr(self, "empleados_tree"):
            return

        for item in self.empleados_tree.get_children():
            self.empleados_tree.delete(item)

        try:
            empleados_df = self.datos_empleados.read().copy()
        except Exception as exc:
            messagebox.showerror("Empleados", f"No se pudieron cargar los empleados: {exc}")
            return

        if empleados_df.empty:
            return

        empleados_df = empleados_df.sort_values(by="NOMBRE_Y_APELLIDO")
        for _, row in empleados_df.iterrows():
            self.empleados_tree.insert(
                "",
                "end",
                values=(
                    str(row.get("NOMBRE_Y_APELLIDO", "")),
                    self._display_value(row.get("VALOR_HS_JORNAL", "")),
                    self._display_value(row.get("HS_JORNAL", "")),
                    self._display_value(row.get("VALOR_HS_EXTRAS", "")),
                ),
            )

    def on_empleado_selected(self, _event):
        selected = self.empleados_tree.selection()
        if not selected:
            return

        values = self.empleados_tree.item(selected[0], "values")
        if not values:
            return

        self.selected_empleado_nombre = str(values[0]).strip().upper()
        self.emp_nombre_var.set(str(values[0]))
        self.emp_valor_hs_jornal_var.set(str(values[1]))
        self.emp_hs_jornal_var.set(str(values[2]))
        self.emp_valor_hs_extras_var.set(str(values[3]))

    def clear_empleado_form(self):
        self.selected_empleado_nombre = ""
        self.emp_nombre_var.set("")
        self.emp_valor_hs_jornal_var.set("")
        self.emp_hs_jornal_var.set("")
        self.emp_valor_hs_extras_var.set("")

    def add_empleado(self):
        try:
            nombre = (self.emp_nombre_var.get() or "").strip()
            if not nombre:
                raise ValueError("El nombre es obligatorio.")

            valor_hs_jornal = self._parse_float(self.emp_valor_hs_jornal_var.get(), "valor hs jornal")
            hs_jornal = self._parse_float(self.emp_hs_jornal_var.get(), "hs jornal")
            valor_hs_extras = self._parse_float(self.emp_valor_hs_extras_var.get(), "valor hs extras")

            self.datos_empleados.add_employee(nombre, valor_hs_jornal, hs_jornal, valor_hs_extras)
            self.refresh_empleados_table()
            self.clear_empleado_form()
            messagebox.showinfo("Empleados", "Empleado agregado correctamente.")
        except ValueError as exc:
            messagebox.showerror("Empleados", str(exc))
        except Exception as exc:
            messagebox.showerror("Empleados", f"No se pudo agregar el empleado: {exc}")

    def update_empleado(self):
        try:
            nombre_original = self.selected_empleado_nombre
            if not nombre_original:
                raise ValueError("Seleccioná un empleado para actualizar.")

            nombre_form = (self.emp_nombre_var.get() or "").strip().upper()
            if nombre_form != nombre_original:
                raise ValueError("No se puede cambiar el nombre del empleado en la actualizacion.")

            valor_hs_jornal = self._parse_float(self.emp_valor_hs_jornal_var.get(), "valor hs jornal")
            hs_jornal = self._parse_float(self.emp_hs_jornal_var.get(), "hs jornal")
            valor_hs_extras = self._parse_float(self.emp_valor_hs_extras_var.get(), "valor hs extras")

            self.datos_empleados.update_employee_data(nombre_original, valor_hs_jornal, hs_jornal, valor_hs_extras)
            self.refresh_empleados_table()
            messagebox.showinfo("Empleados", "Empleado actualizado correctamente.")
        except ValueError as exc:
            messagebox.showerror("Empleados", str(exc))
        except Exception as exc:
            messagebox.showerror("Empleados", f"No se pudo actualizar el empleado: {exc}")

    def remove_empleado(self):
        try:
            nombre = self.selected_empleado_nombre
            if not nombre:
                raise ValueError("Seleccioná un empleado para eliminar.")

            if not messagebox.askyesno("Empleados", f"Eliminar a '{nombre}'?"):
                return

            self.datos_empleados.remove_employee(nombre)
            self.refresh_empleados_table()
            self.clear_empleado_form()
            messagebox.showinfo("Empleados", "Empleado eliminado correctamente.")
        except ValueError as exc:
            messagebox.showerror("Empleados", str(exc))
        except Exception as exc:
            messagebox.showerror("Empleados", f"No se pudo eliminar el empleado: {exc}")

    def refresh_feriados_table(self):
        if not hasattr(self, "feriados_table"):
            return

        for item in self.feriados_table.get_children():
            self.feriados_table.delete(item)

        try:
            feriados_df = self.feriados_reader.read().copy()
        except Exception as exc:
            messagebox.showerror("Feriados", f"No se pudieron cargar los feriados: {exc}")
            return

        if feriados_df.empty:
            return

        feriados_df = feriados_df.sort_values(by="FECHA_FERIADO")
        for _, row in feriados_df.iterrows():
            fecha = row.get("FECHA_FERIADO")
            fecha_txt = fecha.strftime("%d/%m/%Y") if not pd.isna(fecha) else ""
            self.feriados_table.insert(
                "",
                "end",
                values=(fecha_txt, str(row.get("DESCRIPCION_FERIADO", ""))),
            )

    def on_feriado_selected(self, _event):
        selected = self.feriados_table.selection()
        if not selected:
            return

        values = self.feriados_table.item(selected[0], "values")
        if not values:
            return

        self.feriado_fecha_var.set(str(values[0]))
        self.feriado_desc_var.set(str(values[1]))
        self.selected_feriado_fecha = self._parse_feriado_date(str(values[0]))

    def clear_feriado_form(self):
        self.selected_feriado_fecha = None
        self.feriado_fecha_var.set("")
        self.feriado_desc_var.set("")

    def add_feriado(self):
        try:
            fecha = self._parse_feriado_date(self.feriado_fecha_var.get())
            descripcion = (self.feriado_desc_var.get() or "").strip()
            if not descripcion:
                raise ValueError("La descripcion es obligatoria.")

            self.feriados_reader.add_date(fecha, descripcion)
            self.refresh_feriados_table()
            self.clear_feriado_form()
            messagebox.showinfo("Feriados", "Feriado agregado correctamente.")
        except ValueError as exc:
            messagebox.showerror("Feriados", str(exc))
        except Exception as exc:
            messagebox.showerror("Feriados", f"No se pudo agregar el feriado: {exc}")

    def update_feriado(self):
        try:
            if self.selected_feriado_fecha is None:
                raise ValueError("Seleccioná un feriado para actualizar.")

            nueva_fecha = self._parse_feriado_date(self.feriado_fecha_var.get())
            nueva_descripcion = (self.feriado_desc_var.get() or "").strip()
            if not nueva_descripcion:
                raise ValueError("La descripcion es obligatoria.")

            self.feriados_reader.update_date(self.selected_feriado_fecha, nueva_fecha, nueva_descripcion)
            self.refresh_feriados_table()
            self.clear_feriado_form()
            messagebox.showinfo("Feriados", "Feriado actualizado correctamente.")
        except ValueError as exc:
            messagebox.showerror("Feriados", str(exc))
        except Exception as exc:
            messagebox.showerror("Feriados", f"No se pudo actualizar el feriado: {exc}")

    def remove_feriado(self):
        try:
            if self.selected_feriado_fecha is None:
                raise ValueError("Seleccioná un feriado para eliminar.")

            fecha_txt = self.selected_feriado_fecha.strftime("%d/%m/%Y")
            if not messagebox.askyesno("Feriados", f"Eliminar el feriado '{fecha_txt}'?"):
                return

            self.feriados_reader.remove_date(self.selected_feriado_fecha)
            self.refresh_feriados_table()
            self.clear_feriado_form()
            messagebox.showinfo("Feriados", "Feriado eliminado correctamente.")
        except ValueError as exc:
            messagebox.showerror("Feriados", str(exc))
        except Exception as exc:
            messagebox.showerror("Feriados", f"No se pudo eliminar el feriado: {exc}")

    # =========================
    # VISTA VALIDACION
    # =========================
    def _build_validacion_view(self):
        parent = self.views["validacion"]

        self._build_upload_card(parent)
        self._build_filters_card(parent)
        self._build_detail_card(parent)
        self._build_bottom_actions(parent)

    def _build_upload_card(self, parent):
        _, card = self._create_card(parent, "Carga de reporte Qontact")

        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x")

        self.path_var = tk.StringVar(value="Sin archivo seleccionado")

        ttk.Button(
            row,
            text="Subir Archivo",
            style="Secondary.TButton",
            command=self.select_excel
        ).pack(side="left")

        ttk.Entry(
            row,
            textvariable=self.path_var,
            state="readonly",
            width=90
        ).pack(side="left", fill="x", expand=True, padx=(12, 0))

    def _build_filters_card(self, parent):
        _, card = self._create_card(parent, "Filtros")

        self.row_status_var = tk.StringVar(value="")
        self.nombre_var = tk.StringVar(value="")
        self.fecha_desde_var = tk.StringVar(value="")
        self.fecha_hasta_var = tk.StringVar(value="")

        ttk.Label(card, text="Estado", style="SectionLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.row_status_combo = ttk.Combobox(
            card,
            textvariable=self.row_status_var,
            values=["", "NO_CONFIRMADO", "CONFIRMADO"],
            width=18,
            state="readonly"
        )
        self.row_status_combo.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(card, text="Nombre", style="SectionLabel.TLabel").grid(row=0, column=2, sticky="w", padx=(18, 8), pady=4)
        self.nombre_entry = ttk.Entry(card, textvariable=self.nombre_var, width=26)
        self.nombre_entry.grid(row=0, column=3, sticky="w", pady=4)

        ttk.Label(card, text="Desde", style="SectionLabel.TLabel").grid(row=0, column=4, sticky="w", padx=(18, 8), pady=4)
        desde_frame = ttk.Frame(card, style="Card.TFrame")
        desde_frame.grid(row=0, column=5, sticky="w", pady=4)
        self.fecha_desde_entry = ttk.Entry(desde_frame, textvariable=self.fecha_desde_var, width=15, state="readonly")
        self.fecha_desde_entry.pack(side="left")
        ttk.Button(
            desde_frame,
            text="📅",
            style="CalendarIcon.TButton",
            command=lambda: self._open_date_picker(self.fecha_desde_var),
        ).pack(side="left", padx=(6, 0))

        ttk.Label(card, text="Hasta", style="SectionLabel.TLabel").grid(row=0, column=6, sticky="w", padx=(18, 8), pady=4)
        hasta_frame = ttk.Frame(card, style="Card.TFrame")
        hasta_frame.grid(row=0, column=7, sticky="w", pady=4)
        self.fecha_hasta_entry = ttk.Entry(hasta_frame, textvariable=self.fecha_hasta_var, width=15, state="readonly")
        self.fecha_hasta_entry.pack(side="left")
        ttk.Button(
            hasta_frame,
            text="📅",
            style="CalendarIcon.TButton",
            command=lambda: self._open_date_picker(self.fecha_hasta_var),
        ).pack(side="left", padx=(6, 0))

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=0, column=9, padx=(20, 0), sticky="e")
        card.columnconfigure(8, weight=1)

        ttk.Button(actions, text="Aplicar filtros", style="Primary.TButton", command=self.apply_filters).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Limpiar", style="Secondary.TButton", command=self.clear_filters).pack(side="left")

    def _build_validation_message(self, parent):
        self.validation_card, self.validation_inner = self._create_card(parent, "Validación")

        self.validation_message_var = tk.StringVar(
            value="Listo para cargar y validar registros."
        )

        self.validation_label = tk.Label(
            self.validation_inner,
            textvariable=self.validation_message_var,
            bg="#f8d7da",
            fg="#7a1f27",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
            padx=14,
            pady=12
        )
        self.validation_label.pack(fill="x")

    def _build_detail_card(self, parent):
        detail_outer, detail_inner = self._create_card(parent, "Detalle Calculado")
        detail_outer.pack_configure(fill="both", expand=True)

        detail_container = ttk.Frame(detail_inner, style="Card.TFrame")
        detail_container.pack(fill="both", expand=True)

        tree_columns = [self.SELECTION_COLUMN] + self.DISPLAY_COLUMNS
        self.tree = ttk.Treeview(detail_container, columns=tree_columns, show="headings", height=18)
        self.tree.pack(side="left", fill="both", expand=True)

        for col in tree_columns:
            self.tree.heading(col, text=col)
            width = 150
            if col == self.SELECTION_COLUMN:
                width = 90
            elif col == "COMENTARIOS":
                width = 220
            elif col in {"NOMBRE_Y_APELLIDO", "NOMBRE Y APELLIDO"}:
                width = 220
            self.tree.column(col, width=width, anchor="w")

        scrollbar = ttk.Scrollbar(detail_container, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<Button-1>", self.on_single_click)
        self.tree.bind("<Double-1>", self.on_cell_double_click)

    def _build_bottom_actions(self, parent):
        actions = ttk.Frame(parent, style="Content.TFrame")
        actions.pack(fill="x", pady=(0, 8))

        ttk.Button(
            actions,
            text="Descartar",
            style="Danger.TButton",
            command=self.discard_selected
        ).pack(side="right", padx=(10, 0))

        ttk.Button(
            actions,
            text="Confirmar",
            style="Success.TButton",
            command=self.confirm_loaded
        ).pack(side="right", padx=(10, 0))

        ttk.Button(
            actions,
            text="Editar",
            style="Primary.TButton",
            command=self._edit_selected_row
        ).pack(side="right", padx=(10, 0))

    def _open_date_picker(self, target_var: tk.StringVar):
        popup = tk.Toplevel(self.root)
        popup.title("Seleccionar fecha")
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()

        raw_value = (target_var.get() or "").strip()
        current_date = None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                current_date = datetime.datetime.strptime(raw_value, fmt).date()
                break
            except ValueError:
                continue
        if current_date is None:
            current_date = datetime.date.today()

        state = {"year": current_date.year, "month": current_date.month}

        header = ttk.Frame(popup, padding=(10, 10, 10, 6))
        header.pack(fill="x")

        days_container = ttk.Frame(popup, padding=(10, 0, 10, 10))
        days_container.pack(fill="both", expand=True)

        month_label = ttk.Label(header, text="", style="CardTitle.TLabel")

        def select_date(day: int):
            selected = datetime.date(state["year"], state["month"], day)
            target_var.set(selected.strftime("%d/%m/%Y"))
            popup.destroy()

        def render_calendar():
            for child in days_container.winfo_children():
                child.destroy()

            month_label.configure(text=f"{calendar.month_name[state['month']]} {state['year']}")

            weekday_labels = ["L", "M", "X", "J", "V", "S", "D"]
            for idx, label in enumerate(weekday_labels):
                ttk.Label(days_container, text=label, width=4, anchor="center").grid(row=0, column=idx, pady=(0, 4))

            month_rows = calendar.monthcalendar(state["year"], state["month"])
            for week_idx, week in enumerate(month_rows, start=1):
                for day_idx, day in enumerate(week):
                    if day == 0:
                        ttk.Label(days_container, text="", width=4).grid(row=week_idx, column=day_idx)
                        continue

                    ttk.Button(
                        days_container,
                        text=str(day),
                        style="Secondary.TButton",
                        width=3,
                        command=lambda d=day: select_date(d),
                    ).grid(row=week_idx, column=day_idx, padx=1, pady=1)

        def prev_month():
            if state["month"] == 1:
                state["month"] = 12
                state["year"] -= 1
            else:
                state["month"] -= 1
            render_calendar()

        def next_month():
            if state["month"] == 12:
                state["month"] = 1
                state["year"] += 1
            else:
                state["month"] += 1
            render_calendar()

        ttk.Button(header, text="<", style="Secondary.TButton", width=3, command=prev_month).pack(side="left")
        month_label.pack(side="left", padx=10)
        ttk.Button(header, text=">", style="Secondary.TButton", width=3, command=next_month).pack(side="left")

        footer = ttk.Frame(popup, padding=(10, 0, 10, 10))
        footer.pack(fill="x")
        ttk.Button(
            footer,
            text="Hoy",
            style="Secondary.TButton",
            command=lambda: target_var.set(datetime.date.today().strftime("%d/%m/%Y")),
        ).pack(side="left")
        ttk.Button(
            footer,
            text="Limpiar",
            style="Secondary.TButton",
            command=lambda: target_var.set(""),
        ).pack(side="left", padx=(8, 0))
        ttk.Button(footer, text="Cerrar", style="Primary.TButton", command=popup.destroy).pack(side="right")

        render_calendar()

    # =========================
    # VISTA REPORTES
    # =========================
    def _build_reportes_view(self):
        parent = self.views["reportes"]

        _, card = self._create_card(parent, "Reporte")

        ttk.Label(card, text="Seleccionar Empleado:", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        empleado_row = ttk.Frame(card, style="Card.TFrame")
        empleado_row.pack(fill="x", pady=(0, 18))

        ttk.Label(empleado_row, text="Empleado:", style="SectionLabel.TLabel").pack(side="left", padx=(0, 10))

        self.reporte_empleado_var = tk.StringVar()
        self.reporte_empleado_combo = ttk.Combobox(
            empleado_row,
            textvariable=self.reporte_empleado_var,
            values=[],
            state="readonly",
            width=50
        )
        self.reporte_empleado_combo.pack(side="left", fill="x", expand=True)

        ttk.Label(card, text="Seleccionar Período:", style="CardTitle.TLabel").pack(anchor="w", pady=(10, 8))

        period_row = ttk.Frame(card, style="Card.TFrame")
        period_row.pack(fill="x", pady=(0, 24))

        ttk.Label(period_row, text="Desde:", style="SectionLabel.TLabel").pack(side="left", padx=(0, 8))
        self.reporte_desde_var = tk.StringVar()
        reporte_desde_frame = ttk.Frame(period_row, style="Card.TFrame")
        reporte_desde_frame.pack(side="left")
        ttk.Entry(reporte_desde_frame, textvariable=self.reporte_desde_var, width=15, state="readonly").pack(side="left")
        ttk.Button(
            reporte_desde_frame,
            text="📅",
            style="CalendarIcon.TButton",
            command=lambda: self._open_date_picker(self.reporte_desde_var),
        ).pack(side="left", padx=(6, 0))

        ttk.Label(period_row, text="Hasta:", style="SectionLabel.TLabel").pack(side="left", padx=(30, 8))
        self.reporte_hasta_var = tk.StringVar()
        reporte_hasta_frame = ttk.Frame(period_row, style="Card.TFrame")
        reporte_hasta_frame.pack(side="left")
        ttk.Entry(reporte_hasta_frame, textvariable=self.reporte_hasta_var, width=15, state="readonly").pack(side="left")
        ttk.Button(
            reporte_hasta_frame,
            text="📅",
            style="CalendarIcon.TButton",
            command=lambda: self._open_date_picker(self.reporte_hasta_var),
        ).pack(side="left", padx=(6, 0))

        ttk.Button(
            card,
            text="Generar Reporte",
            style="Success.TButton",
            command=self.generar_reporte
        ).pack(anchor="center", pady=(8, 0))

    # =========================
    # DATOS / TABLA
    # =========================
    def ensure_selection_column(self):
        if self.INTERNAL_SELECTION_COLUMN not in self.preview_df.columns:
            self.preview_df[self.INTERNAL_SELECTION_COLUMN] = False

    def load_historico_with_current_filters(self):
        row_status = self.row_status_var.get() if hasattr(self, "row_status_var") else ""
        if row_status == "ELIMINADO":
            row_status = ""
            self.row_status_var.set("")

        self.preview_df = self.workflow.get_historico_filtered(
            row_status=row_status,
            nombre_filtro=self.nombre_var.get() if hasattr(self, "nombre_var") else "",
            fecha_desde=self.fecha_desde_var.get() if hasattr(self, "fecha_desde_var") else "",
            fecha_hasta=self.fecha_hasta_var.get() if hasattr(self, "fecha_hasta_var") else "",
        )

        # En la vista de validacion se ocultan siempre los registros eliminados.
        if "ROW_STATUS" in self.preview_df.columns:
            self.preview_df = self.preview_df[self.preview_df["ROW_STATUS"] != "ELIMINADO"].copy()

        self.ensure_selection_column()
        self.temporal_loaded = True
        self.refresh_table()
        self._refresh_reportes_empleados()

    def refresh_table(self):
        self.ensure_selection_column()

        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.preview_df.empty:
            return

        for idx, row in self.preview_df.iterrows():
            marcado = bool(row.get(self.INTERNAL_SELECTION_COLUMN, False))
            checkbox = "[x]" if marcado else "[ ]"
            values = [checkbox] + [self._display_value(row.get(col, "")) for col in self.DISPLAY_COLUMNS]
            self.tree.insert("", "end", iid=str(idx), values=values)

    def _refresh_reportes_empleados(self):
        if not hasattr(self, "reporte_empleado_combo"):
            return

        posibles = []
        for col in ["NOMBRE_Y_APELLIDO", "NOMBRE Y APELLIDO"]:
            if col in self.preview_df.columns:
                posibles = sorted(self.preview_df[col].dropna().astype(str).unique().tolist())
                break

        self.reporte_empleado_combo["values"] = posibles

    def _display_value(self, value):
        if pd.isna(value):
            return ""
        if isinstance(value, pd.Timestamp):
            return value.strftime("%d/%m/%Y %H:%M:%S")
        if isinstance(value, float):
            return f"{value:.2f}"
        return str(value)

    # =========================
    # EVENTOS
    # =========================
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

        new_value = simpledialog.askstring(
            "Editar",
            f"Nuevo valor para {column_name}:",
            initialvalue=current_value
        )

        if new_value is None:
            return

        try:
            parsed = self.workflow.parse_edition_value(column_name, new_value)
            self.preview_df.at[row_index, column_name] = parsed
            self.refresh_table()

            if self.temporal_loaded:
                self.preview_df, self.temporal_loaded = self.workflow.load_temporal(self.preview_df, True)
        except ValueError:
            messagebox.showerror("Valor inválido", "Para columnas numéricas debés ingresar un número válido.")

    def _edit_selected_row(self):
        selected = self.tree.focus()
        if not selected:
            messagebox.showwarning("Sin selección", "Seleccioná una fila para editar.")
            return
        self.on_cell_double_click(None)

    # =========================
    # ACCIONES PRINCIPALES
    # =========================
    def apply_filters(self):
        try:
            self.load_historico_with_current_filters()
        except ValueError as exc:
            messagebox.showerror("Filtro inválido", str(exc))

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
                "OK",
                "Archivo procesado y cargado temporalmente. Se aplicó el filtro NO_CONFIRMADO."
            )
        except ValueError as exc:
            messagebox.showerror("Configuración incompleta", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {exc}")

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
            messagebox.showinfo("Confirmado", "Filas marcadas confirmadas en histórico.")
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

            self.preview_df.loc[
                self.preview_df["ID"].isin(selected_ids),
                self.INTERNAL_SELECTION_COLUMN
            ] = False

            self.refresh_table()
            messagebox.showinfo("Descartado", "Filas marcadas descartadas en histórico.")
        except ValueError as exc:
            messagebox.showwarning("Sin filas marcadas", str(exc))
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo descartar la carga: {exc}")

    def generar_reporte(self):
        empleado = self.reporte_empleado_var.get().strip()
        desde = self.reporte_desde_var.get().strip()
        hasta = self.reporte_hasta_var.get().strip()

        if not empleado:
            messagebox.showwarning("Falta empleado", "Seleccioná un empleado.")
            return

        if not desde or not hasta:
            messagebox.showwarning("Falta período", "Indicá fecha desde y hasta.")
            return

        # Acá conectás con tu service real
        messagebox.showinfo(
            "Reporte",
            f"Generar reporte para:\n\nEmpleado: {empleado}\nDesde: {desde}\nHasta: {hasta}"
        )

    def run(self):
        self.root.mainloop()