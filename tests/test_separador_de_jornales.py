import unittest

import pandas as pd

from src.separador_de_jornales import SeparadorDeJornales


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
        self.assertAlmostEqual(result[3], 0.0)
        self.assertAlmostEqual(result[4], 0.0)
        self.assertAlmostEqual(result[5], 0.0)

        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-31 07:00"),
            pd.Timestamp("2026-08-31 18:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 9.0)
        self.assertAlmostEqual(result[2], 2.0)

    def test_sabado_hasta_13_hs_es_extra_50_y_despues_100(self):
        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-29 08:00"),
            pd.Timestamp("2026-08-29 15:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 0.0)
        self.assertAlmostEqual(result[2], 5.0)
        self.assertAlmostEqual(result[4], 2.0)

    def test_domingo_y_feriado_son_extra_100(self):
        result = self.splitter.split_hours(
            pd.Timestamp("2026-08-30 08:00"),
            pd.Timestamp("2026-08-30 12:00"),
            hs_jornal=9.0,
        )
        self.assertAlmostEqual(result[0], 0.0)
        self.assertAlmostEqual(result[4], 4.0)


if __name__ == "__main__":
    unittest.main()
