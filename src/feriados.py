import pandas as pd

class FeriadosReader:
    def __init__(self):
        folder_path = r"C:\Users\inaki.costa\Downloads\GitHub_Repositories\Sistema-Pago-Horas-Extras\data"
        self.excel_path = folder_path + r"\Feriados.xlsx"

    def read(self) -> pd.DataFrame:
        feriados_df = pd.read_excel(self.excel_path)
        feriados_df = feriados_df.rename(columns={
            "FECHA": "FECHA_FERIADO",
            "DESCRIPCION": "DESCRIPCION_FERIADO"
        })

        feriados_df["FECHA_FERIADO"] = pd.to_datetime(feriados_df["FECHA_FERIADO"], format="%d/%m/%Y")

        return feriados_df
    
    def add_date(self, fecha_feriado: pd.Timestamp, descripcion_feriado: str):
        df = self.read()
        
        nuevo_feriado = {
            "FECHA_FERIADO": fecha_feriado,
            "DESCRIPCION_FERIADO": descripcion_feriado.strip()
        }

        if df["FECHA_FERIADO"].eq(nuevo_feriado["FECHA_FERIADO"]).any():
            raise ValueError(f"La fecha '{fecha_feriado.strftime('%d/%m/%Y')}' ya existe en el sistema.")

        df = df.append(nuevo_feriado, ignore_index=True)

        df.to_excel(self.excel_path, index=False)

    def remove_date(self, fecha_feriado: pd.Timestamp):
        df = self.read()
        
        if df["FECHA_FERIADO"].eq(fecha_feriado).any():
            df = df[df["FECHA_FERIADO"] != fecha_feriado]
        else:
            raise ValueError(f"La fecha '{fecha_feriado.strftime('%d/%m/%Y')}' no existe en el sistema.")

        df.to_excel(self.excel_path, index=False)

    def is_holiday(self, date: pd.Timestamp) -> bool:
        df = self.read()
        return df["FECHA_FERIADO"].eq(date.normalize()).any()
    
    def update_date(self, old_fecha_feriado: pd.Timestamp, new_fecha_feriado: pd.Timestamp, new_descripcion_feriado: str):
        df = self.read()
        
        if df["FECHA_FERIADO"].eq(old_fecha_feriado).any():
            df.loc[df["FECHA_FERIADO"] == old_fecha_feriado, ["FECHA_FERIADO", "DESCRIPCION_FERIADO"]] = [
                new_fecha_feriado,
                new_descripcion_feriado.strip()
            ]
        else:
            raise ValueError(f"La fecha '{old_fecha_feriado.strftime('%d/%m/%Y')}' no existe en el sistema.")

        df.to_excel(self.excel_path, index=False)