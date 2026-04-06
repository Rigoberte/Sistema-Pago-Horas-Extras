import pandas as pd

from src.Qontact_report_reader import ReporteHorasExtras
from src.controlador_historico import ControladorHistorico
from src.datos_empleados_reader import DatosEmpleados
from src.separador_de_jornales import SeparadorDeJornales
from src.time_utils import round_timestamp_to_nearest_half_hour


class HorasExtrasWorkflowService:
    MULTIPLIER_HORAS_NORMALES = 1.0
    MULTIPLIER_HORAS_EXTRAS = 1.5
    MULTIPLIER_HORAS_NOCTURNAS = 2.0

    TABLE_COLUMNS = [
        "ID",
        "ROW_STATUS",
        "NOMBRE_Y_APELLIDO",
        "INGRESO",
        "EGRESO",
        "COMENTARIOS",
        "VALOR_HS_JORNAL",
        "IMPORTE",
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

        return self.recalculate_importes(df[self.TABLE_COLUMNS])

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
        df = self.recalculate_importes(df)
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

        preview_df = result_df.copy()
        preview_df["ID"] = ""
        preview_df["ROW_STATUS"] = "NO_CONFIRMADO"

        if "COMENTARIOS" not in preview_df.columns:
            preview_df["COMENTARIOS"] = ""

        preview_df = self.recalculate_importes(preview_df)

        return preview_df[self.TABLE_COLUMNS]

    def build_manual_record(self, row_data: dict) -> pd.DataFrame:
        row_data = row_data.copy()
        row_data["INGRESO"] = round_timestamp_to_nearest_half_hour(pd.to_datetime(row_data["INGRESO"]))
        row_data["EGRESO"] = round_timestamp_to_nearest_half_hour(pd.to_datetime(row_data["EGRESO"]))

        base_df = pd.DataFrame([row_data])
        base_df["NOMBRE_Y_APELLIDO"] = base_df["NOMBRE_Y_APELLIDO"].astype(str).str.strip().str.upper()

        # En carga manual las horas se calculan automaticamente desde ingreso/egreso.
        preview_df = SeparadorDeJornales(base_df).build_result_df()
        preview_df = self.recalculate_importes(preview_df)
        preview_df["ID"] = ""
        preview_df["ROW_STATUS"] = "NO_CONFIRMADO"

        if "COMENTARIOS" not in preview_df.columns:
            preview_df["COMENTARIOS"] = ""

        return preview_df[self.TABLE_COLUMNS]

    def load_temporal(self, preview_df: pd.DataFrame, already_loaded: bool) -> tuple[pd.DataFrame, bool]:
        updated_df = self.recalculate_importes(preview_df.copy())

        if not already_loaded:
            ids = self.historico.add_temporal_records(updated_df)
            updated_df["ID"] = ids
            updated_df["ROW_STATUS"] = "NO_CONFIRMADO"
            return updated_df, True

        self.historico.update_records(updated_df)
        return updated_df, True

    def confirm_loaded(self, preview_df: pd.DataFrame, already_loaded: bool) -> tuple[pd.DataFrame, bool]:
        updated_df = self.recalculate_importes(preview_df.copy())

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
        updated_df = self.recalculate_importes(preview_df.copy())

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
        updated_df = self.recalculate_importes(preview_df.copy())

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

    @staticmethod
    def _parse_numeric(series: pd.Series) -> pd.Series:
        return pd.to_numeric(series, errors="coerce").fillna(0.0)

    @staticmethod
    def _normalize_name(value: str) -> str:
        if pd.isna(value):
            return ""
        return " ".join(str(value).upper().strip().split())

    def _build_valor_jornal_map(self) -> dict[str, float]:
        empleados_df = DatosEmpleados().read().copy()
        empleados_df["__MATCH_KEY__"] = empleados_df["NOMBRE_Y_APELLIDO"].apply(self._normalize_name)

        duplicados = empleados_df[empleados_df["__MATCH_KEY__"].duplicated(keep=False)]
        if not duplicados.empty:
            nombres = "\n".join(f"- {name}" for name in sorted(duplicados["NOMBRE_Y_APELLIDO"].unique().tolist()))
            raise ValueError(
                "No se puede hacer un match unico porque hay empleados duplicados:\n\n" + nombres
            )

        return dict(
            zip(
                empleados_df["__MATCH_KEY__"],
                pd.to_numeric(empleados_df["VALOR_HS_JORNAL"], errors="coerce"),
            )
        )

    def recalculate_importes(self, df: pd.DataFrame, enrich_valor_jornal: bool = False) -> pd.DataFrame:
        result_df = df.copy()
        if result_df.empty:
            for column in self.TABLE_COLUMNS:
                if column not in result_df.columns:
                    result_df[column] = "" if column in {"ID", "ROW_STATUS", "NOMBRE_Y_APELLIDO", "COMENTARIOS"} else 0.0
            return result_df

        for hour_column in [
            "HORAS_NORMALES_DIURNAS",
            "HORAS_EXTRAS_NORMALES",
            "HORAS_NOCTURNAS",
        ]:
            if hour_column not in result_df.columns:
                result_df[hour_column] = 0.0

        if "VALOR_HS_JORNAL" not in result_df.columns:
            result_df["VALOR_HS_JORNAL"] = 0.0

        if enrich_valor_jornal:
            valor_map = self._build_valor_jornal_map()
            result_df["__MATCH_KEY__"] = result_df["NOMBRE_Y_APELLIDO"].apply(self._normalize_name)

            result_df["VALOR_HS_JORNAL"] = result_df.apply(
                lambda row: valor_map.get(row["__MATCH_KEY__"], float("nan")),
                axis=1,
            )

            no_match = result_df[result_df["VALOR_HS_JORNAL"].isna() | (pd.to_numeric(result_df["VALOR_HS_JORNAL"], errors="coerce").isna())]
            if not no_match.empty:
                nombres = "\n".join(f"- {name}" for name in sorted(no_match["NOMBRE_Y_APELLIDO"].astype(str).unique().tolist()))
                raise ValueError(
                    "No se encontro un empleado para estos nombres/apellidos:\n\n" + nombres
                )

            result_df = result_df.drop(columns=["__MATCH_KEY__"])

        valor_hs_jornal = self._parse_numeric(result_df["VALOR_HS_JORNAL"])
        horas_normales = self._parse_numeric(result_df["HORAS_NORMALES_DIURNAS"])
        horas_extras = self._parse_numeric(result_df["HORAS_EXTRAS_NORMALES"])
        horas_nocturnas = self._parse_numeric(result_df["HORAS_NOCTURNAS"])

        importe_calculado = valor_hs_jornal * (
            horas_normales * self.MULTIPLIER_HORAS_NORMALES
            + horas_extras * self.MULTIPLIER_HORAS_EXTRAS
            + horas_nocturnas * self.MULTIPLIER_HORAS_NOCTURNAS
        )

        if "ROW_STATUS" in result_df.columns and "IMPORTE" in result_df.columns:
            estados = result_df["ROW_STATUS"].astype(str).str.upper()
            existentes = pd.to_numeric(result_df["IMPORTE"], errors="coerce")
            es_dinamico = estados == "NO_CONFIRMADO"
            result_df["IMPORTE"] = existentes.where(~es_dinamico & existentes.notna(), importe_calculado)
        else:
            result_df["IMPORTE"] = importe_calculado

        return result_df
