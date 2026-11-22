"""Tests de la lógica de Parole. Ejecutar con:  py -m unittest discover tests"""

import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from calculos import (calcular_condena, convertir_dias, edad_en, fecha_larga,  # noqa: E402
                      nacimiento_desde_ci, parsear_fecha, rellenar_plantilla)


class TestFechas(unittest.TestCase):
    def test_fecha_larga(self):
        self.assertEqual(fecha_larga(datetime(2026, 3, 15)), "15 de marzo de 2026")
        self.assertEqual(fecha_larga(None), "")

    def test_parsear_fecha_invalida(self):
        with self.assertRaises(ValueError):
            parsear_fecha("31/02/2026")

    def test_convertir_dias(self):
        self.assertEqual(convertir_dias(0), "0 días")
        self.assertEqual(convertir_dias(1), "1 día")
        self.assertEqual(convertir_dias(400), "1 año, 1 mes, 5 días")
        self.assertEqual(convertir_dias(-31), "-1 mes, 1 día")


class TestCondena(unittest.TestCase):
    def setUp(self):
        self.entrada = datetime(2024, 1, 1)
        self.salida = datetime(2027, 1, 1)  # 1096 días
        self.hoy = datetime(2026, 1, 1)

    def test_fechas_sin_rebaja(self):
        r = calcular_condena(self.entrada, self.salida, hoy=self.hoy)
        self.assertEqual(r.condena_efectiva_dias, 1096)
        self.assertEqual(r.mitad, datetime(2025, 7, 2))
        self.assertEqual(r.tercio, datetime(2024, 12, 31))
        self.assertEqual(r.salida_con_rebaja, self.salida)

    def test_rebaja_adelanta_todas_las_fechas(self):
        r = calcular_condena(self.entrada, self.salida, rebaja=30, hoy=self.hoy)
        self.assertEqual(r.salida_con_rebaja, datetime(2026, 12, 2))
        self.assertEqual((r.mitad - r.mitad_con_rebaja).days, 30)
        self.assertEqual(r.dias_cumplidos, r.dias_transcurridos + 30)

    def test_preventiva_reconstruye_condena_original(self):
        r = calcular_condena(self.entrada, self.salida, prev_dias=100, hoy=self.hoy)
        self.assertEqual(r.condena_original_dias, 1196)
        self.assertEqual((r.salida_original_con_preventiva - self.salida).days, 100)

    def test_antes_de_entrar_y_despues_de_salir(self):
        antes = calcular_condena(self.entrada, self.salida, hoy=datetime(2023, 1, 1))
        self.assertEqual(antes.dias_transcurridos, 0)
        despues = calcular_condena(self.entrada, self.salida, hoy=datetime(2030, 1, 1))
        self.assertEqual(despues.dias_restantes, 0)

    def test_datos_invalidos(self):
        with self.assertRaises(ValueError):
            calcular_condena(self.salida, self.entrada)
        with self.assertRaises(ValueError):
            calcular_condena(self.entrada, self.salida, rebaja=-5)


class TestCarnet(unittest.TestCase):
    def test_carnet_completo(self):
        self.assertEqual(nacimiento_desde_ci("85031512345"), datetime(1985, 3, 15))
        self.assertEqual(nacimiento_desde_ci("05110278912"), datetime(2005, 11, 2))

    def test_carnet_incompleto_usa_solo_el_anio(self):
        hoy = datetime(2026, 6, 1)
        self.assertEqual(nacimiento_desde_ci("85", hoy), datetime(1985, 1, 1))
        self.assertIsNone(nacimiento_desde_ci("ab123"))

    def test_edad_respeta_el_cumpleanos(self):
        nac = datetime(2000, 6, 15)
        self.assertEqual(edad_en(nac, datetime(2026, 6, 14)), 25)
        self.assertEqual(edad_en(nac, datetime(2026, 6, 15)), 26)


class TestPlantilla(unittest.TestCase):
    def test_sustituye_marcadores_y_negritas(self):
        from docx import Document
        doc = Document()
        rellenar_plantilla(
            doc,
            "El recluso (nombre) (art. 58) cumple el (MitadOriginal).",
            {"(nombre)": "Juan Pérez", "(MitadOriginal)": "2 de julio de 2025"},
            negrita=["(nombre)"],
        )
        texto = doc.paragraphs[0].text
        self.assertEqual(texto, "El recluso Juan Pérez (art. 58) cumple el 2 de julio de 2025.")
        negritas = [run.text for run in doc.paragraphs[0].runs if run.bold]
        self.assertEqual(negritas, ["Juan Pérez"])

    def test_lineas_vacias_seguidas_cuentan_como_una(self):
        from docx import Document
        doc = Document()
        rellenar_plantilla(doc, "A\n\n\n\nB", {})
        self.assertEqual([p.text for p in doc.paragraphs], ["A", "", "B"])


if __name__ == "__main__":
    unittest.main()
