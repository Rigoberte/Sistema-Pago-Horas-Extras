import pandas as pd
from pathlib import Path

class FeriadosReader:
    def __init__(self):
        folder_path = Path(__file__).resolve().parent.parent / "data"
        self.excel_path = str(folder_path / "Feriados.xlsx")

    def read(self) -> pd.DataFrame:
        feriados_df = pd.read_excel(self.excel_path)
        feriados_df = feriados_df.rename(columns={
            "FECHA": "FECHA_FERIADO",
            "DESCRIPCION": "DESCRIPCION_FERIADO"
        })

        feriados_df["FECHA_FERIADO"] = pd.to_datetime(feriados_df["FECHA_FERIADO"], dayfirst=True, errors="coerce").dt.normalize()

        return feriados_df

    @staticmethod
    def _normalize_date(fecha: pd.Timestamp) -> pd.Timestamp:
        return pd.to_datetime(fecha, errors="coerce").normalize()
    
    def add_date(self, fecha_feriado: pd.Timestamp, descripcion_feriado: str):
        df = self.read()
        fecha_feriado = self._normalize_date(fecha_feriado)
        descripcion_feriado = descripcion_feriado.strip()
        
        nuevo_feriado = {
            "FECHA_FERIADO": fecha_feriado,
            "DESCRIPCION_FERIADO": descripcion_feriado
        }

        if df["FECHA_FERIADO"].dt.normalize().eq(nuevo_feriado["FECHA_FERIADO"]).any():
            raise ValueError(f"La fecha '{fecha_feriado.strftime('%d/%m/%Y')}' ya existe en el sistema.")

        df = pd.concat([df, pd.DataFrame([nuevo_feriado])], ignore_index=True)

        df.to_excel(self.excel_path, index=False)

    def remove_date(self, fecha_feriado: pd.Timestamp):
        df = self.read()
        fecha_feriado = self._normalize_date(fecha_feriado)
        
        mask = df["FECHA_FERIADO"].dt.normalize().eq(fecha_feriado)
        if mask.any():
            df = df[~mask]
        else:
            raise ValueError(f"La fecha '{fecha_feriado.strftime('%d/%m/%Y')}' no existe en el sistema.")

        df.to_excel(self.excel_path, index=False)

    def is_holiday(self, date_input: pd.Timestamp) -> bool:
        df = self.read()
        date = self._normalize_date(date_input)
        return df["FECHA_FERIADO"].dt.normalize().eq(date).any()
    
    def update_date(
        self,
        old_fecha_feriado: pd.Timestamp,
        old_descripcion_feriado: str,
        new_fecha_feriado: pd.Timestamp,
        new_descripcion_feriado: str,
    ):
        df = self.read()
        old_fecha_feriado = self._normalize_date(old_fecha_feriado)
        new_fecha_feriado = self._normalize_date(new_fecha_feriado)
        old_descripcion_feriado = (old_descripcion_feriado or "").strip()
        new_descripcion_feriado = (new_descripcion_feriado or "").strip()

        if old_fecha_feriado != new_fecha_feriado:
            date_exists = df["FECHA_FERIADO"].dt.normalize().eq(new_fecha_feriado).any()
            if date_exists:
                raise ValueError(f"La fecha '{new_fecha_feriado.strftime('%d/%m/%Y')}' ya existe en el sistema.")
        
        mask_selected = (
            df["FECHA_FERIADO"].dt.normalize().eq(old_fecha_feriado)
            & df["DESCRIPCION_FERIADO"].fillna("").astype(str).str.strip().eq(old_descripcion_feriado)
        )

        if not mask_selected.any():
            raise ValueError(f"La fecha '{old_fecha_feriado.strftime('%d/%m/%Y')}' no existe en el sistema.")

        selected_index = df.index[mask_selected][0]
        df.at[selected_index, "FECHA_FERIADO"] = new_fecha_feriado
        df.at[selected_index, "DESCRIPCION_FERIADO"] = new_descripcion_feriado

        df.to_excel(self.excel_path, index=False)