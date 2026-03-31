import pandas as pd
import datetime

from src.Qontact_report_reader import ReporteHorasExtras
from src.datos_empleados_reader import DatosEmpleados
from src.feriados import FeriadosReader

class SeparadorDeJornales:
    def __init__(self, reporte_horas_extras_df: pd.DataFrame | None = None):
        self.reporte_horas_extras_df: pd.DataFrame = reporte_horas_extras_df if reporte_horas_extras_df is not None else ReporteHorasExtras().read()
        self.datos_empleados_df: pd.DataFrame = DatosEmpleados().read()
        self.feriados_reader: FeriadosReader = FeriadosReader()

    def build_result_df(self, reporte_horas_extras_df: pd.DataFrame | None = None) -> pd.DataFrame:
        reporte_base = reporte_horas_extras_df if reporte_horas_extras_df is not None else self.reporte_horas_extras_df

        reporte_horas_extras_df = reporte_base.merge(
            self.datos_empleados_df,
            on="NOMBRE_Y_APELLIDO",
            how="left"
        )

        reporte_horas_extras_df["HORAS_TRABAJADAS"] = (
            reporte_horas_extras_df["EGRESO"] - reporte_horas_extras_df["INGRESO"]
        ).dt.total_seconds() / 3600

        resultados = reporte_horas_extras_df.apply(
            lambda row: self.split_hours(row["INGRESO"], row["EGRESO"], row["HS_JORNAL"]),
            axis=1,
            result_type="expand"
        )

        resultados.columns = [
            "HORAS_NORMALES_DIURNAS",
            "HORAS_EXTRAS_NORMALES",
            "HORAS_NOCTURNAS"
        ]

        reporte_horas_extras_df = pd.concat([reporte_horas_extras_df, resultados], axis=1)
        return reporte_horas_extras_df

    def split_jornales(self):
        reporte_horas_extras_df = self.build_result_df()
    
        print(reporte_horas_extras_df[[
            "NOMBRE_Y_APELLIDO",
            "INGRESO",
            "EGRESO",
            "HORAS_TRABAJADAS",
            "HORAS_NORMALES_DIURNAS",
            "HORAS_EXTRAS_NORMALES",
            "HORAS_NOCTURNAS"
        ]])


    def is_night(self, dt: pd.Timestamp) -> bool:
        t = dt.time()
        return t >= datetime.time(21, 0) or t < datetime.time(6, 0)
    
    def is_holiday_or_weekend(self, dt: pd.Timestamp) -> bool:
        return dt.weekday() >= 5 or self.feriados_reader.is_holiday(dt)


    def next_boundary(self, dt: pd.Timestamp) -> pd.Timestamp:
        """
        Devuelve el próximo corte de franja:
        - 06:00
        - 21:00
        - o el día siguiente si corresponde
        """
        current_date = dt.normalize()
        t = dt.time()

        six_am = current_date + pd.Timedelta(hours=6)
        nine_pm = current_date + pd.Timedelta(hours=21)

        if t < datetime.time(6, 0):
            return six_am
        elif t < datetime.time(21, 0):
            return nine_pm
        else:
            return current_date + pd.Timedelta(days=1, hours=6)


    def split_hours(self, ingreso: pd.Timestamp, egreso: pd.Timestamp, hs_jornal: float) -> tuple[float, float, float]:
        """
        Retorna:
        - horas_normales_diurnas
        - horas_extras_normales
        - horas_nocturnas

        Regla:
        - Todo lo trabajado entre 21:00 y 06:00 es nocturno
        - Las extras normales son solo las horas DIURNAS que queden
            luego de superar las hs_jornal horas trabajadas
        - Una hora nocturna NO puede ser también extra normal
        """
        if egreso <= ingreso:
            return 0.0, 0.0, 0.0

        horas_normales_diurnas = 0.0
        horas_extras_normales = 0.0
        horas_nocturnas = 0.0

        horas_acumuladas = 0.0
        actual = ingreso

        while actual < egreso:
            corte = min(self.next_boundary(actual), egreso)
            duracion = (corte - actual).total_seconds() / 3600

            nocturna = self.is_night(actual)
            feriado_o_findesemana = self.is_holiday_or_weekend(actual)

            # Esta fracción cae dentro de las primeras 9 horas del turno
            horas_restantes_normales = max(hs_jornal - horas_acumuladas, 0.0)
            parte_normal = min(duracion, horas_restantes_normales)
            parte_excedente = duracion - parte_normal

            if nocturna:
                # Todo lo nocturno se paga como nocturno, aunque esté después de las 9 horas
                horas_nocturnas += duracion
            elif feriado_o_findesemana:
                # Todo lo diurno en feriado o finde se paga como extra normal
                horas_extras_normales += duracion
            else:
                # Solo lo diurno que excede las 9 horas es extra normal
                horas_normales_diurnas += parte_normal
                horas_extras_normales += parte_excedente

            horas_acumuladas += duracion
            actual = corte

        return horas_normales_diurnas, horas_extras_normales, horas_nocturnas