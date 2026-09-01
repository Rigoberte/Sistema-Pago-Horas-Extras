import unittest

import pandas as pd

from src.separador_de_jornales import SeparadorDeJornales
from src.workflow_service import HorasExtrasWorkflowService


class TestSeparadorDeJornales(unittest.TestCase):
    def setUp(self):
        self.splitter = SeparadorDeJornales()

    def test_lunes_viernes_hasta_9_horas_es_normal_y_luego_50(self):
        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-31 08:00"),
            pd.Timestamp("2026-08-31 10:30"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 2.5)
        self.assertAlmostEqual(result[1], 0.0)
        self.assertAlmostEqual(result[2], 0.0)

        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-31 07:00"),
            pd.Timestamp("2026-08-31 18:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 9.0)
        self.assertAlmostEqual(result[1], 2.0)
        self.assertAlmostEqual(result[2], 0.0)

    def test_sabado_hasta_13_hs_es_extra_50_y_despues_100(self):
        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-29 08:00"),
            pd.Timestamp("2026-08-29 15:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 0.0)
        self.assertAlmostEqual(result[1], 5.0)
        self.assertAlmostEqual(result[2], 2.0)

    def test_sabado_cruza_13_hs_se_divide_en_50_y_100(self):
        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-29 10:00"),
            pd.Timestamp("2026-08-29 15:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 0.0)
        self.assertAlmostEqual(result[1], 3.0)
        self.assertAlmostEqual(result[2], 2.0)

    def test_domingo_y_feriado_son_extra_100(self):
        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-30 08:00"),
            pd.Timestamp("2026-08-30 12:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 0.0)
        self.assertAlmostEqual(result[1], 0.0)
        self.assertAlmostEqual(result[2], 4.0)

    def test_recalcula_compatibilidad_importada_con_total_sin_desglose(self):
        workflow = HorasExtrasWorkflowService()
        df = pd.DataFrame([
            {
                "VALOR_HS_JORNAL": 1000,
                "HORAS_TRABAJADAS": 9.0,
                "HORAS_NORMALES": "",
                "HORAS_EXTRAS_50": "",
                "HORAS_EXTRAS_100": "",
            }
        ])

        result = workflow.recalculate_importes(df)

        self.assertAlmostEqual(float(result["HORAS_NORMALES"].iloc[0]), 9.0)
        self.assertAlmostEqual(float(result["HORAS_EXTRAS_50"].iloc[0]), 0.0)
        self.assertAlmostEqual(float(result["HORAS_EXTRAS_100"].iloc[0]), 0.0)

    def test_no_reescribe_filas_existentes_al_agregar_una_carga_manual(self):
        workflow = HorasExtrasWorkflowService()
        df = pd.DataFrame([
            {
                "NOMBRE_Y_APELLIDO": "JUAN PEREZ",
                "INGRESO": pd.Timestamp("2026-08-31 07:00"),
                "EGRESO": pd.Timestamp("2026-08-31 18:00"),
                "VALOR_HS_JORNAL": 1000,
                "HORAS_TRABAJADAS": 11.0,
                "HORAS_NORMALES": 9.0,
                "HORAS_EXTRAS_50": 2.0,
                "HORAS_EXTRAS_100": 0.0,
            },
            {
                "NOMBRE_Y_APELLIDO": "ANA LOPEZ",
                "INGRESO": pd.Timestamp("2026-08-31 07:00"),
                "EGRESO": pd.Timestamp("2026-08-31 17:00"),
                "VALOR_HS_JORNAL": 1000,
                "HORAS_TRABAJADAS": 10.0,
                "HORAS_NORMALES": 0.0,
                "HORAS_EXTRAS_50": 0.0,
                "HORAS_EXTRAS_100": 0.0,
            },
        ])

        result = workflow.recalculate_importes(df)

        self.assertAlmostEqual(float(result["HORAS_NORMALES"].iloc[0]), 9.0)
        self.assertAlmostEqual(float(result["HORAS_EXTRAS_50"].iloc[0]), 2.0)
        self.assertAlmostEqual(float(result["HORAS_NORMALES"].iloc[1]), 9.0)
        self.assertAlmostEqual(float(result["HORAS_EXTRAS_50"].iloc[1]), 1.0)
        self.assertAlmostEqual(float(result["HORAS_EXTRAS_100"].iloc[1]), 0.0)


if __name__ == "__main__":
    unittest.main()
