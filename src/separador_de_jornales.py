import datetime

import pandas as pd

from src.Qontact_report_reader import ReporteHorasExtras
from src.datos_empleados_reader import DatosEmpleados
from src.feriados import FeriadosReader


class SeparadorDeJornales:
    LEGACY_COLUMNS = [
        "HORAS_NORMALES_DIURNAS",
        "HORAS_NORMALES_NOCTURNAS",
        "HORAS_EXTRAS_DIURNAS",
        "HORAS_EXTRAS_NOCTURNAS",
        "HORAS_EXTRAS_DIURNAS_FERIADO",
        "HORAS_EXTRAS_NOCTURNAS_FERIADO",
    ]
    NEW_COLUMNS = [
        "HORAS_NORMALES",
        "HORAS_EXTRAS_50",
        "HORAS_EXTRAS_100",
    ]

    def __init__(self, reporte_horas_extras_df: pd.DataFrame | None = None):
        self.reporte_horas_extras_df: pd.DataFrame = (
            reporte_horas_extras_df if reporte_horas_extras_df is not None else pd.DataFrame()
        )
        self.datos_empleados_df: pd.DataFrame = DatosEmpleados().read()
        self.feriados_reader: FeriadosReader = FeriadosReader()

    def _ensure_compatibility_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        compatible = df.copy()
        for column in self.NEW_COLUMNS:
            if column not in compatible.columns:
                compatible[column] = 0.0
        for column in self.LEGACY_COLUMNS:
            if column not in compatible.columns:
                compatible[column] = 0.0

        if "HORAS_NORMALES" not in compatible.columns:
            compatible["HORAS_NORMALES"] = (
                compatible.get("HORAS_NORMALES_DIURNAS", 0.0)
                + compatible.get("HORAS_NORMALES_NOCTURNAS", 0.0)
            )
        if "HORAS_EXTRAS_50" not in compatible.columns:
            compatible["HORAS_EXTRAS_50"] = (
                compatible.get("HORAS_EXTRAS_DIURNAS", 0.0)
                + compatible.get("HORAS_EXTRAS_NOCTURNAS", 0.0)
                + compatible.get("HORAS_EXTRAS_DIURNAS_FERIADO", 0.0) * 0.0
            )
        if "HORAS_EXTRAS_100" not in compatible.columns:
            compatible["HORAS_EXTRAS_100"] = (
                compatible.get("HORAS_EXTRAS_DIURNAS_FERIADO", 0.0)
                + compatible.get("HORAS_EXTRAS_NOCTURNAS_FERIADO", 0.0)
            )

        return compatible

    def build_result_df(self, reporte_horas_extras_df: pd.DataFrame | None = None) -> pd.DataFrame:
        reporte_base = reporte_horas_extras_df if reporte_horas_extras_df is not None else self.reporte_horas_extras_df
        reporte_horas_extras_df = self._match_empleados_unico(reporte_base)

        if reporte_horas_extras_df.empty:
            empty_columns = ["HORAS_TRABAJADAS", *self.NEW_COLUMNS, *self.LEGACY_COLUMNS]
            for column_name in empty_columns:
                reporte_horas_extras_df[column_name] = 0.0
            return reporte_horas_extras_df

        reporte_horas_extras_df["HORAS_TRABAJADAS"] = (
            reporte_horas_extras_df["EGRESO"] - reporte_horas_extras_df["INGRESO"]
        ).dt.total_seconds() / 3600

        resultados = reporte_horas_extras_df.apply(
            lambda row: self.split_hours(
                row["INGRESO"],
                row["EGRESO"],
                row["HS_JORNAL"],
            ),
            axis=1,
            result_type="expand",
        )

        resultados.columns = self.NEW_COLUMNS
        reporte_horas_extras_df = pd.concat([reporte_horas_extras_df, resultados], axis=1)
        reporte_horas_extras_df = self._ensure_compatibility_columns(reporte_horas_extras_df)
        return reporte_horas_extras_df

    @staticmethod
    def _normalize_name(value: str) -> str:
        if pd.isna(value):
            return ""
        return " ".join(str(value).upper().strip().split())

    def _match_empleados_unico(self, reporte_df: pd.DataFrame) -> pd.DataFrame:
        empleados_df = self.datos_empleados_df.copy()
        empleados_df["__MATCH_KEY__"] = empleados_df["NOMBRE_Y_APELLIDO"].apply(self._normalize_name)

        duplicados = empleados_df[empleados_df["__MATCH_KEY__"].duplicated(keep=False)]
        if not duplicados.empty:
            nombres = "\n".join(f"- {name}" for name in sorted(duplicados["NOMBRE_Y_APELLIDO"].unique().tolist()))
            raise ValueError(
                "No se puede hacer un match unico porque hay empleados duplicados:\n\n" + nombres
            )

        reporte = reporte_df.copy()
        reporte["__MATCH_KEY__"] = reporte["NOMBRE_Y_APELLIDO"].apply(self._normalize_name)

        for column_name in ["HS_JORNAL", "VALOR_HS_JORNAL"]:
            if column_name not in reporte.columns:
                reporte[column_name] = pd.NA

        merged_df = reporte.merge(
            empleados_df[["__MATCH_KEY__", "HS_JORNAL", "VALOR_HS_JORNAL"]],
            on="__MATCH_KEY__",
            how="left",
            suffixes=("", "_EMP"),
        )

        merged_df["HS_JORNAL"] = pd.to_numeric(merged_df["HS_JORNAL"], errors="coerce").combine_first(
            pd.to_numeric(merged_df["HS_JORNAL_EMP"], errors="coerce")
        )
        merged_df["VALOR_HS_JORNAL"] = pd.to_numeric(merged_df["VALOR_HS_JORNAL"], errors="coerce").combine_first(
            pd.to_numeric(merged_df["VALOR_HS_JORNAL_EMP"], errors="coerce")
        )

        faltantes = merged_df[merged_df["HS_JORNAL"].isna()]["NOMBRE_Y_APELLIDO"].dropna().astype(str).unique().tolist()
        if faltantes:
            empleados = "\n".join(f"- {name}" for name in sorted(faltantes))
            raise ValueError(
                "No se encontro un empleado para estos nombres/apellidos:\n\n" + empleados
            )

        return merged_df.drop(columns=["__MATCH_KEY__", "HS_JORNAL_EMP", "VALOR_HS_JORNAL_EMP"])

    def split_jornales(self):
        reporte_horas_extras_df = self.build_result_df()

        print(reporte_horas_extras_df[[
            "NOMBRE_Y_APELLIDO",
            "EDIFICIO",
            "INGRESO",
            "EGRESO",
            "HORAS_TRABAJADAS",
            "HORAS_NORMALES",
            "HORAS_EXTRAS_50",
            "HORAS_EXTRAS_100",
        ]])

    def is_holiday_or_weekend(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() >= 5 or self.feriados_reader.is_holiday(dt)

    def is_sunday_or_holiday(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() == 6 or self.feriados_reader.is_holiday(dt)

    def is_saturday_extra_50(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() == 5 and dt.time() < datetime.time(13, 0)

    def is_saturday_extra_100(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() == 5 and dt.time() >= datetime.time(13, 0)

    def split_hours(
        self,
        ingreso: pd.Timestamp,
        egreso: pd.Timestamp,
        hs_jornal: float,
    ) -> tuple[float, float, float]:
        """
        Retorna:
        - horas_normales
        - horas_extras_50
        - horas_extras_100
        """
        if egreso <= ingreso:
            return 0.0, 0.0, 0.0

        horas_normales = 0.0
        horas_extras_50 = 0.0
        horas_extras_100 = 0.0
        horas_por_dia: dict[pd.Timestamp, float] = {}
        actual = ingreso

        while actual < egreso:
            dia = actual.normalize()
            dia_fin = dia + pd.Timedelta(days=1)
            corte = min(egreso, dia_fin)
            duracion = (corte - actual).total_seconds() / 3600

            es_domingo_o_feriado = self.is_sunday_or_holiday(actual)
            es_sabado = actual.weekday() == 5
            es_sabado_50 = es_sabado and actual.time() < datetime.time(13, 0)
            es_sabado_100 = es_sabado and actual.time() >= datetime.time(13, 0)

            if es_domingo_o_feriado:
                horas_extras_100 += duracion
            elif es_sabado:
                if es_sabado_50:
                    limite_13 = dia + pd.Timedelta(hours=13)
                    resto = min(corte, limite_13)
                    horas_extras_50 += max((resto - actual).total_seconds() / 3600, 0.0)
                    actual = resto
                    continue
                if es_sabado_100:
                    horas_extras_100 += duracion
            else:
                horas_dia = horas_por_dia.get(dia, 0.0)
                normales_disponibles = max(hs_jornal - horas_dia, 0.0)
                parte_normal = min(duracion, normales_disponibles)
                parte_extra_50 = max(duracion - parte_normal, 0.0)
                horas_normales += parte_normal
                horas_extras_50 += parte_extra_50
                horas_por_dia[dia] = horas_dia + duracion

            actual = corte

            if es_sabado and actual > (dia + pd.Timedelta(hours=13)) and actual <= corte:
                sabado_100 = min(corte, actual)
                horas_extras_100 += max((corte - max(actual, dia + pd.Timedelta(hours=13))).total_seconds() / 3600, 0.0)
                actual = corte

        return horas_normales, horas_extras_50, horas_extras_100
