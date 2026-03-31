import pandas as pd

class ReporteHorasExtras:
    def __init__(self):
        folder_path = r"C:\Users\inaki.costa\Downloads\GitHub_Repositories\Sistema-Pago-Horas-Extras\data"
        self.excel_path = folder_path + r"\ReporteHS.xlsx"

    def read(self) -> pd.DataFrame:
        reporte_horas_extras_df = pd.read_excel(self.excel_path)
        
        columns_rename = {
            "NOMBRE Y APELLIDO": "NOMBRE_Y_APELLIDO",
            "DESDE": "FECHA_INGRESO",
            "HASTA": "FECHA_EGRESO",
            "INGRESO": "HORA_INGRESO",
            "EGRESO": "HORA_EGRESO"
        }

        reporte_horas_extras_df = reporte_horas_extras_df.loc[reporte_horas_extras_df["NOMBRE Y APELLIDO"].notna(), columns_rename.keys()]
        reporte_horas_extras_df.rename(columns=columns_rename, inplace=True)

        reporte_horas_extras_df["NOMBRE_Y_APELLIDO"] = reporte_horas_extras_df["NOMBRE_Y_APELLIDO"].str.strip().str.upper()

        reporte_horas_extras_df["FECHA_INGRESO"] = pd.to_datetime(reporte_horas_extras_df["FECHA_INGRESO"], format="%d/%m/%Y")
        reporte_horas_extras_df["FECHA_EGRESO"] = pd.to_datetime(reporte_horas_extras_df["FECHA_EGRESO"], format="%d/%m/%Y")

        reporte_horas_extras_df["HORA_INGRESO"] = self.parse_time_column(reporte_horas_extras_df["HORA_INGRESO"])
        reporte_horas_extras_df["HORA_EGRESO"] = self.parse_time_column(reporte_horas_extras_df["HORA_EGRESO"])

        reporte_horas_extras_df["INGRESO"] = reporte_horas_extras_df.apply(
            lambda row: pd.Timestamp.combine(row["FECHA_INGRESO"], row["HORA_INGRESO"]),
            axis=1
        )
        reporte_horas_extras_df["EGRESO"] = reporte_horas_extras_df.apply(
            lambda row: pd.Timestamp.combine(row["FECHA_EGRESO"], row["HORA_EGRESO"]),
            axis=1
        )

        return reporte_horas_extras_df

    def parse_time_column(self, series: pd.Series) -> pd.Series:
        values = series.astype(str).str.strip()

        parsed = pd.to_datetime(values, format="%H:%M:%S", errors="coerce")
        parsed = parsed.fillna(pd.to_datetime(values, format="%H:%M", errors="coerce"))

        return parsed.dt.time