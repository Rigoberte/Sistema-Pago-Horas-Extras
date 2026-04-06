import pandas as pd
from pathlib import Path

class DatosEmpleados:
    REQUIRED_COLUMNS = [
        "NOMBRE_Y_APELLIDO",
        "VALOR_HS_JORNAL",
        "HS_JORNAL",
        "IGNORAR_PERIODO_NOCTURNO",
    ]

    def __init__(self):
        folder_path = Path(__file__).resolve().parent.parent / "data"
        self.excel_path = str(folder_path / "DatosEmpleados.xlsx")

    def read(self) -> pd.DataFrame:
        datos_empleados_df = pd.read_excel(self.excel_path)
        datos_empleados_df = datos_empleados_df.rename(columns={
            "NOMBRE Y APELLIDO": "NOMBRE_Y_APELLIDO",
            "VALOR HS JORNAL": "VALOR_HS_JORNAL",
            "HS JORNAL": "HS_JORNAL",
            "IGNORAR PERIODO NOCTURNO": "IGNORAR_PERIODO_NOCTURNO",
        })

        for column in self.REQUIRED_COLUMNS:
            if column not in datos_empleados_df.columns:
                datos_empleados_df[column] = ""

        datos_empleados_df = datos_empleados_df[self.REQUIRED_COLUMNS]

        datos_empleados_df["NOMBRE_Y_APELLIDO"] = datos_empleados_df["NOMBRE_Y_APELLIDO"].str.strip().str.upper()
        datos_empleados_df["IGNORAR_PERIODO_NOCTURNO"] = datos_empleados_df["IGNORAR_PERIODO_NOCTURNO"].apply(self._to_bool)

        return datos_empleados_df

    @staticmethod
    def _to_bool(value) -> bool:
        if isinstance(value, bool):
            return value
        if pd.isna(value):
            return False

        text = str(value).strip().lower()
        if text in {"true", "1", "si", "sí", "yes", "y", "t"}:
            return True
        if text in {"false", "0", "no", "n", "f", ""}:
            return False
        return False
    
    def add_employee(self, nombre_y_apellido: str, valor_hs_jornal: float, hs_jornal: float, ignorar_periodo_nocturno: bool = False):
        df = self.read()
        
        nuevo_empleado = {
            "NOMBRE_Y_APELLIDO": nombre_y_apellido.strip().upper(),
            "VALOR_HS_JORNAL": valor_hs_jornal,
            "HS_JORNAL": hs_jornal,
            "IGNORAR_PERIODO_NOCTURNO": bool(ignorar_periodo_nocturno),
        }

        if df["NOMBRE_Y_APELLIDO"].str.upper().str.strip().eq(nuevo_empleado["NOMBRE_Y_APELLIDO"]).any():
            raise ValueError(f"El empleado '{nombre_y_apellido}' ya existe en el sistema.")

        df = df.append(nuevo_empleado, ignore_index=True)

        df.to_excel(self.excel_path, index=False)
    
    def update_employee_data(
        self,
        nombre_y_apellido: str,
        valor_hs_jornal: float,
        hs_jornal: float,
        ignorar_periodo_nocturno: bool = False,
    ):
        df = self.read()
        
        if not df.loc[df["NOMBRE_Y_APELLIDO"] == nombre_y_apellido].empty:
            df.loc[
                df["NOMBRE_Y_APELLIDO"] == nombre_y_apellido,
                ["VALOR_HS_JORNAL", "HS_JORNAL", "IGNORAR_PERIODO_NOCTURNO"],
            ] = [
                valor_hs_jornal,
                hs_jornal,
                bool(ignorar_periodo_nocturno),
            ]
        else:
            raise ValueError(f"El empleado '{nombre_y_apellido}' no existe en el sistema.")

        df.to_excel(self.excel_path, index=False)

    def remove_employee(self, nombre_y_apellido: str):
        df = self.read()
        
        if not df.loc[df["NOMBRE_Y_APELLIDO"] == nombre_y_apellido].empty:
            df = df[df["NOMBRE_Y_APELLIDO"] != nombre_y_apellido]
        else:
            raise ValueError(f"El empleado '{nombre_y_apellido}' no existe en el sistema.")

        df.to_excel(self.excel_path, index=False)