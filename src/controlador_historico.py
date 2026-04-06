import pandas as pd
import uuid
from pathlib import Path


class ControladorHistorico:
    REQUIRED_COLUMNS = [
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

    TEXT_COLUMNS = [
        "ID",
        "ROW_STATUS",
        "NOMBRE_Y_APELLIDO",
        "COMENTARIOS",
    ]

    NUMERIC_COLUMNS = [
        "VALOR_HS_JORNAL",
        "IMPORTE",
        "HORAS_TRABAJADAS",
        "HORAS_NORMALES_DIURNAS",
        "HORAS_EXTRAS_NORMALES",
        "HORAS_NOCTURNAS",
    ]

    def __init__(self):
        folder_path = Path(__file__).resolve().parent.parent / "data"
        self.excel_path = str(folder_path / "Historico.xlsx")

    def read(self) -> pd.DataFrame:
        historico_df = pd.read_excel(self.excel_path)

        for column in self.REQUIRED_COLUMNS:
            if column not in historico_df.columns:
                historico_df[column] = ""

        historico_df["INGRESO"] = pd.to_datetime(historico_df["INGRESO"], dayfirst=True, errors="coerce")
        historico_df["EGRESO"] = pd.to_datetime(historico_df["EGRESO"], dayfirst=True, errors="coerce")

        for column in self.TEXT_COLUMNS:
            historico_df[column] = historico_df[column].fillna("").astype(str)

        for column in self.NUMERIC_COLUMNS:
            historico_df[column] = pd.to_numeric(historico_df[column], errors="coerce")

        historico_df = historico_df[self.REQUIRED_COLUMNS]

        return historico_df

    def save(self, historico_df: pd.DataFrame):
        historico_df.to_excel(self.excel_path, index=False)

    def add_temporal_records(self, records_df: pd.DataFrame) -> list[str]:
        historico_df = self.read()

        nuevos = records_df.copy()
        nuevos["ID"] = [str(uuid.uuid4()) for _ in range(len(nuevos))]
        nuevos["ROW_STATUS"] = "NO_CONFIRMADO"

        for column in self.REQUIRED_COLUMNS:
            if column not in nuevos.columns:
                nuevos[column] = ""

        nuevos = nuevos[self.REQUIRED_COLUMNS]
        historico_df = pd.concat([historico_df, nuevos], ignore_index=True)
        self.save(historico_df)

        return nuevos["ID"].tolist()

    def update_records(self, records_df: pd.DataFrame):
        historico_df = self.read()

        if "ID" not in records_df.columns:
            raise ValueError("Se requiere columna ID para actualizar registros.")

        editable_columns = [
            "COMENTARIOS",
            "VALOR_HS_JORNAL",
            "IMPORTE",
            "HORAS_TRABAJADAS",
            "HORAS_NORMALES_DIURNAS",
            "HORAS_EXTRAS_NORMALES",
            "HORAS_NOCTURNAS",
        ]

        updates = records_df.set_index("ID")
        historico_df_indexed = historico_df.set_index("ID")
        common_ids = historico_df_indexed.index.intersection(updates.index)

        for column in editable_columns:
            if column in updates.columns:
                historico_df_indexed.loc[common_ids, column] = updates.loc[common_ids, column]

        historico_df = historico_df_indexed.reset_index()
        self.save(historico_df)
    
    def remove_record(self, record_id: str):
        historico_df = self.read()
        historico_df.loc[historico_df["ID"] == record_id, "ROW_STATUS"] = "ELIMINADO"
        self.save(historico_df)

    def confirm_records(self, record_ids: list[str]):
        historico_df = self.read()
        historico_df.loc[historico_df["ID"].isin(record_ids), "ROW_STATUS"] = "CONFIRMADO"
        self.save(historico_df)

    def confirm_record(self, record_id: str):
        self.confirm_records([record_id])
