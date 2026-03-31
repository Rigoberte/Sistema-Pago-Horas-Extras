import pandas as pd

class DatosEmpleados:
    def __init__(self):
        folder_path = r"C:\Users\inaki.costa\Downloads\GitHub_Repositories\Sistema-Pago-Horas-Extras\data"
        self.excel_path = folder_path + r"\DatosEmpleados.xlsx"

    def read(self) -> pd.DataFrame:
        datos_empleados_df = pd.read_excel(self.excel_path)
        datos_empleados_df = datos_empleados_df.rename(columns={
            "NOMBRE Y APELLIDO": "NOMBRE_Y_APELLIDO",
            "VALOR HS JORNAL": "VALOR_HS_JORNAL",
            "HS JORNAL": "HS_JORNAL",
            "VALOR HS EXTRAS": "VALOR_HS_EXTRAS",
        })

        datos_empleados_df["NOMBRE_Y_APELLIDO"] = datos_empleados_df["NOMBRE_Y_APELLIDO"].str.strip().str.upper()

        return datos_empleados_df
    
    def add_employee(self, nombre_y_apellido: str, valor_hs_jornal: float, hs_jornal: float, valor_hs_extras: float):
        df = self.read()
        
        nuevo_empleado = {
            "NOMBRE_Y_APELLIDO": nombre_y_apellido.strip().upper(),
            "VALOR_HS_JORNAL": valor_hs_jornal,
            "HS_JORNAL": hs_jornal,
            "VALOR_HS_EXTRAS": valor_hs_extras
        }

        if df["NOMBRE_Y_APELLIDO"].str.upper().str.strip().eq(nuevo_empleado["NOMBRE_Y_APELLIDO"]).any():
            raise ValueError(f"El empleado '{nombre_y_apellido}' ya existe en el sistema.")

        df = df.append(nuevo_empleado, ignore_index=True)

        df.to_excel(self.excel_path, index=False)
    
    def update_employee_data(self, nombre_y_apellido: str, valor_hs_jornal: float, hs_jornal: float, valor_hs_extras: float):
        df = self.read()
        
        if not df.loc[df["NOMBRE_Y_APELLIDO"] == nombre_y_apellido].empty:
            df.loc[df["NOMBRE_Y_APELLIDO"] == nombre_y_apellido, ["VALOR_HS_JORNAL", "HS_JORNAL", "VALOR_HS_EXTRAS"]] = [
                valor_hs_jornal,
                hs_jornal,
                valor_hs_extras
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