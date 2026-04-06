import datetime

import pandas as pd

from src.Qontact_report_reader import ReporteHorasExtras
from src.datos_empleados_reader import DatosEmpleados
from src.feriados import FeriadosReader


class SeparadorDeJornales:
    def __init__(self, reporte_horas_extras_df: pd.DataFrame | None = None):
        self.reporte_horas_extras_df: pd.DataFrame = (
            reporte_horas_extras_df if reporte_horas_extras_df is not None else ReporteHorasExtras().read()
        )
        self.datos_empleados_df: pd.DataFrame = DatosEmpleados().read()
        self.feriados_reader: FeriadosReader = FeriadosReader()

    def build_result_df(self, reporte_horas_extras_df: pd.DataFrame | None = None) -> pd.DataFrame:
        reporte_base = reporte_horas_extras_df if reporte_horas_extras_df is not None else self.reporte_horas_extras_df
        reporte_horas_extras_df = self._match_empleados_unico(reporte_base)

        if reporte_horas_extras_df.empty:
            empty_columns = [
                "HORAS_TRABAJADAS",
                "HORAS_NORMALES_DIURNAS",
                "HORAS_NORMALES_NOCTURNAS",
                "HORAS_EXTRAS_DIURNAS",
                "HORAS_EXTRAS_NOCTURNAS",
                "HORAS_EXTRAS_DIURNAS_FERIADO",
                "HORAS_EXTRAS_NOCTURNAS_FERIADO",
            ]
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
                bool(row.get("IGNORAR_PERIODO_NOCTURNO", False)),
            ),
            axis=1,
            result_type="expand",
        )

        resultados.columns = [
            "HORAS_NORMALES_DIURNAS",
            "HORAS_NORMALES_NOCTURNAS",
            "HORAS_EXTRAS_DIURNAS",
            "HORAS_EXTRAS_NOCTURNAS",
            "HORAS_EXTRAS_DIURNAS_FERIADO",
            "HORAS_EXTRAS_NOCTURNAS_FERIADO",
        ]

        reporte_horas_extras_df = pd.concat([reporte_horas_extras_df, resultados], axis=1)
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

        for column_name in ["HS_JORNAL", "VALOR_HS_JORNAL", "IGNORAR_PERIODO_NOCTURNO"]:
            if column_name not in reporte.columns:
                reporte[column_name] = pd.NA

        merged_df = reporte.merge(
            empleados_df[["__MATCH_KEY__", "HS_JORNAL", "VALOR_HS_JORNAL", "IGNORAR_PERIODO_NOCTURNO"]],
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
        merged_df["IGNORAR_PERIODO_NOCTURNO"] = merged_df["IGNORAR_PERIODO_NOCTURNO"].combine_first(
            merged_df["IGNORAR_PERIODO_NOCTURNO_EMP"]
        ).fillna(False).astype(bool)

        faltantes = merged_df[merged_df["HS_JORNAL"].isna()]["NOMBRE_Y_APELLIDO"].dropna().astype(str).unique().tolist()
        if faltantes:
            empleados = "\n".join(f"- {name}" for name in sorted(faltantes))
            raise ValueError(
                "No se encontro un empleado para estos nombres/apellidos:\n\n" + empleados
            )

        return merged_df.drop(columns=["__MATCH_KEY__", "HS_JORNAL_EMP", "VALOR_HS_JORNAL_EMP", "IGNORAR_PERIODO_NOCTURNO_EMP"])

    def split_jornales(self):
        reporte_horas_extras_df = self.build_result_df()

        print(reporte_horas_extras_df[[
            "NOMBRE_Y_APELLIDO",
            "INGRESO",
            "EGRESO",
            "HORAS_TRABAJADAS",
            "HORAS_NORMALES_DIURNAS",
            "HORAS_NORMALES_NOCTURNAS",
            "HORAS_EXTRAS_DIURNAS",
            "HORAS_EXTRAS_NOCTURNAS",
            "HORAS_EXTRAS_DIURNAS_FERIADO",
            "HORAS_EXTRAS_NOCTURNAS_FERIADO",
        ]])

    def is_night(self, dt: pd.Timestamp) -> bool:
        t = dt.time()
        return t >= datetime.time(21, 0) or t < datetime.time(6, 0)

    def is_holiday_or_weekend(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() >= 5 or self.feriados_reader.is_holiday(dt)

    def is_sunday_or_holiday(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() == 6 or self.feriados_reader.is_holiday(dt)

    def is_saturday_daytime_after_thirteen(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() == 5 and datetime.time(13, 0) <= dt.time() < datetime.time(21, 0)

    def next_boundary(self, dt: pd.Timestamp) -> pd.Timestamp:
        """
        Devuelve el próximo corte de franja:
        - 06:00
        - 13:00
        - 21:00
        - o medianoche si corresponde
        """
        current_date = dt.normalize()
        t = dt.time()

        six_am = current_date + pd.Timedelta(hours=6)
        one_pm = current_date + pd.Timedelta(hours=13)
        nine_pm = current_date + pd.Timedelta(hours=21)
        midnight = current_date + pd.Timedelta(days=1)

        if t < datetime.time(6, 0):
            return six_am
        if t < datetime.time(13, 0):
            return one_pm
        if t < datetime.time(21, 0):
            return nine_pm
        return midnight

    def split_hours(
        self,
        ingreso: pd.Timestamp,
        egreso: pd.Timestamp,
        hs_jornal: float,
        ignorar_periodo_nocturno: bool = False,
    ) -> tuple[float, float, float, float, float, float]:
        """
        Retorna:
        - horas_normales_diurnas
        - horas_normales_nocturnas
        - horas_extras_diurnas
        - horas_extras_nocturnas
        - horas_extras_diurnas_feriado
        - horas_extras_nocturnas_feriado
        """
        if egreso <= ingreso:
            return 0.0, 0.0, 0.0, 0.0, 0.0, 0.0

        horas_normales_diurnas = 0.0
        horas_normales_nocturnas = 0.0
        horas_extras_diurnas = 0.0
        horas_extras_nocturnas = 0.0
        horas_extras_diurnas_feriado = 0.0
        horas_extras_nocturnas_feriado = 0.0

        horas_acumuladas = 0.0
        actual = ingreso

        while actual < egreso:
            corte = min(self.next_boundary(actual), egreso)
            duracion = (corte - actual).total_seconds() / 3600

            es_noche_real = self.is_night(actual)
            nocturna = es_noche_real and not ignorar_periodo_nocturno
            domingo_o_feriado = self.is_sunday_or_holiday(actual)
            sabado_tarde = self.is_saturday_daytime_after_thirteen(actual)

            if domingo_o_feriado or sabado_tarde:
                parte_normal = 0.0
                parte_excedente = duracion
            else:
                horas_restantes_normales = max(hs_jornal - horas_acumuladas, 0.0)
                parte_normal = min(duracion, horas_restantes_normales)
                parte_excedente = duracion - parte_normal

            if domingo_o_feriado:
                if nocturna:
                    horas_extras_nocturnas_feriado += duracion
                else:
                    horas_extras_diurnas_feriado += duracion
            elif sabado_tarde:
                horas_extras_diurnas += duracion
            else:
                if nocturna:
                    horas_normales_nocturnas += parte_normal
                    horas_extras_nocturnas += parte_excedente
                else:
                    horas_normales_diurnas += parte_normal
                    horas_extras_diurnas += parte_excedente

            horas_acumuladas += duracion
            actual = corte

        return (
            horas_normales_diurnas,
            horas_normales_nocturnas,
            horas_extras_diurnas,
            horas_extras_nocturnas,
            horas_extras_diurnas_feriado,
            horas_extras_nocturnas_feriado,
        )