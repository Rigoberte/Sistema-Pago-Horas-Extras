import pandas as pd
import uuid


class ControladorHistorico:
    def __init__(self):
        folder_path = r"C:\Users\inaki.costa\Downloads\GitHub_Repositories\Sistema-Pago-Horas-Extras\data"
        self.excel_path = folder_path + r"\Historico.xlsx"

    def read(self) -> pd.DataFrame:
        historico_df = pd.read_excel(self.excel_path)
        
        historico_df["INGRESO"] = pd.to_datetime(historico_df["INGRESO"], format="%d/%m/%Y %H:%M:%S")
        historico_df["EGRESO"] = pd.to_datetime(historico_df["EGRESO"], format="%d/%m/%Y %H:%M:%S")

        return historico_df
    
    def add_record(self, nombre_y_apellido: str, ingreso: pd.Timestamp, horas_trabajadas: float, horas_normales_diurnas: float, horas_extras_normales: float, horas_nocturnas: float):
        historico_df = self.read()

        nuevo_registro = {
            "ID": str(uuid.uuid4()),
            "ROW_STATUS": "NO_CONFIRMADO",
            "NOMBRE_Y_APELLIDO": nombre_y_apellido,
            "INGRESO": ingreso,
            "HORAS_TRABAJADAS": horas_trabajadas,
            "HORAS_NORMALES_DIURNAS": horas_normales_diurnas,
            "HORAS_EXTRAS_NORMALES": horas_extras_normales,
            "HORAS_NOCTURNAS": horas_nocturnas,
        }

        historico_df = historico_df.append(nuevo_registro, ignore_index=True)
        historico_df.to_excel(self.excel_path, index=False)

    def remove_record(self, record_id: str):
        historico_df = self.read()
        historico_df.loc[historico_df["ID"] == record_id, "ROW_STATUS"] = "ELIMINADO"
        historico_df.to_excel(self.excel_path, index=False)

    def confirm_record(self, record_id: str):
        historico_df = self.read()
        historico_df.loc[historico_df["ID"] == record_id, "ROW_STATUS"] = "CONFIRMADO"
        historico_df.to_excel(self.excel_path, index=False)
