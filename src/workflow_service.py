import pandas as pd

from src.Qontact_report_reader import ReporteHorasExtras
from src.controlador_historico import ControladorHistorico
from src.separador_de_jornales import SeparadorDeJornales


class HorasExtrasWorkflowService:
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

    def __init__(self):
        self.historico = ControladorHistorico()

    def get_historico_filtered(
        self,
        row_status: str = "",
        nombre_filtro: str = "",
        fecha_desde: str = "",
        fecha_hasta: str = "",
    ) -> pd.DataFrame:
        df = self.historico.read().copy()

        status = (row_status or "").strip().upper()
        if status:
            df = df[df["ROW_STATUS"].str.upper() == status]

        nombre = (nombre_filtro or "").strip().upper()
        if nombre:
            df = df[df["NOMBRE_Y_APELLIDO"].str.upper().str.contains(nombre, na=False)]

        ts_desde = self._parse_date_filter(fecha_desde, "fecha desde")
        ts_hasta = self._parse_date_filter(fecha_hasta, "fecha hasta")

        if ts_desde is not None:
            df = df[df["INGRESO"].dt.normalize() >= ts_desde]
        if ts_hasta is not None:
            df = df[df["INGRESO"].dt.normalize() <= ts_hasta]

        return df[self.TABLE_COLUMNS]

    def build_reporte_df(
        self,
        fecha_desde: str,
        fecha_hasta: str,
        empleado: str = "",
    ) -> pd.DataFrame:
        ts_desde = self._parse_date_filter(fecha_desde, "fecha desde")
        ts_hasta = self._parse_date_filter(fecha_hasta, "fecha hasta")

        if ts_desde is None or ts_hasta is None:
            raise ValueError("Las fechas desde y hasta son obligatorias.")
        if ts_desde > ts_hasta:
            raise ValueError("La fecha desde no puede ser mayor que la fecha hasta.")

        df = self.historico.read().copy()
        df = df[df["ROW_STATUS"].str.upper() == "CONFIRMADO"]
        df = df[df["INGRESO"].dt.normalize() >= ts_desde]
        df = df[df["INGRESO"].dt.normalize() <= ts_hasta]

        empleado_filtro = (empleado or "").strip().upper()
        if empleado_filtro:
            df = df[df["NOMBRE_Y_APELLIDO"].str.upper() == empleado_filtro]

        report_columns = [
            column_name
            for column_name in self.TABLE_COLUMNS
            if column_name not in {"ID", "ROW_STATUS"}
        ]

        if df.empty:
            return df[report_columns]

        df = df.sort_values(by=["NOMBRE_Y_APELLIDO", "INGRESO"]).reset_index(drop=True)
        return df[report_columns]

    @staticmethod
    def _parse_date_filter(raw_value: str, field_name: str) -> pd.Timestamp | None:
        value = (raw_value or "").strip()
        if not value:
            return None

        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            parsed = pd.to_datetime(value, format=fmt, errors="coerce")
            if not pd.isna(parsed):
                return parsed.normalize()

        raise ValueError(f"Formato invalido para {field_name}. Usa dd/mm/yyyy o yyyy-mm-dd.")

    def build_preview_from_excel(self, excel_path: str) -> pd.DataFrame:
        reporte_df = ReporteHorasExtras().read(excel_path)
        result_df = SeparadorDeJornales(reporte_df).build_result_df()

        faltantes_hs_jornal = result_df[result_df["HS_JORNAL"].isna()]["NOMBRE_Y_APELLIDO"].unique().tolist()
        if faltantes_hs_jornal:
            empleados = "\n".join(f"- {name}" for name in faltantes_hs_jornal)
            raise ValueError(
                "Hay empleados sin HS_JORNAL configurado:\n\n" + empleados
            )

        preview_df = result_df.copy()
        preview_df["ID"] = ""
        preview_df["ROW_STATUS"] = "NO_CONFIRMADO"

        if "COMENTARIOS" not in preview_df.columns:
            preview_df["COMENTARIOS"] = ""

        return preview_df[self.TABLE_COLUMNS]

    def load_temporal(self, preview_df: pd.DataFrame, already_loaded: bool) -> tuple[pd.DataFrame, bool]:
        updated_df = preview_df.copy()

        if not already_loaded:
            ids = self.historico.add_temporal_records(updated_df)
            updated_df["ID"] = ids
            updated_df["ROW_STATUS"] = "NO_CONFIRMADO"
            return updated_df, True

        self.historico.update_records(updated_df)
        return updated_df, True

    def confirm_loaded(self, preview_df: pd.DataFrame, already_loaded: bool) -> tuple[pd.DataFrame, bool]:
        updated_df = preview_df.copy()

        if not already_loaded:
            updated_df, already_loaded = self.load_temporal(updated_df, already_loaded)

        self.historico.update_records(updated_df)
        self.historico.confirm_records(updated_df["ID"].tolist())
        updated_df["ROW_STATUS"] = "CONFIRMADO"

        return updated_df, already_loaded

    def confirm_selected(
        self,
        preview_df: pd.DataFrame,
        selected_ids: list[str],
        already_loaded: bool,
    ) -> tuple[pd.DataFrame, bool]:
        updated_df = preview_df.copy()

        if not already_loaded:
            updated_df, already_loaded = self.load_temporal(updated_df, already_loaded)

        selected_ids = [record_id for record_id in selected_ids if str(record_id).strip()]
        if not selected_ids:
            raise ValueError("No hay filas marcadas para confirmar.")

        self.historico.update_records(updated_df)
        self.historico.confirm_records(selected_ids)

        updated_df.loc[updated_df["ID"].isin(selected_ids), "ROW_STATUS"] = "CONFIRMADO"
        return updated_df, already_loaded

    def discard_selected(
        self,
        preview_df: pd.DataFrame,
        selected_ids: list[str],
        already_loaded: bool,
    ) -> tuple[pd.DataFrame, bool]:
        updated_df = preview_df.copy()

        if not already_loaded:
            updated_df, already_loaded = self.load_temporal(updated_df, already_loaded)

        selected_ids = [record_id for record_id in selected_ids if str(record_id).strip()]
        if not selected_ids:
            raise ValueError("No hay filas marcadas para descartar.")

        self.historico.update_records(updated_df)
        for record_id in selected_ids:
            self.historico.remove_record(record_id)

        updated_df.loc[updated_df["ID"].isin(selected_ids), "ROW_STATUS"] = "ELIMINADO"
        return updated_df, already_loaded

    @staticmethod
    def parse_edition_value(column_name: str, value: str):
        if column_name == "COMENTARIOS":
            return value.strip()
        return float(value)
