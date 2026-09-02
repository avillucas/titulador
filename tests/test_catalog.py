import unittest
from src.catalog import TrayectoCatalog, format_trayecto_data

class TestTrayectoCatalog(unittest.TestCase):
    def setUp(self):
        self.catalog = TrayectoCatalog()

    def test_catalog_loaded(self):
        trayectos = self.catalog.get_all_trayectos()
        self.assertGreater(len(trayectos), 0)

    def test_find_by_code(self):
        soldador = self.catalog.find_by_code("MT30")
        self.assertIsNotNone(soldador)
        self.assertEqual(soldador.get("codigo"), "MT30")

    def test_format_trayecto_data_resolucion_and_hours(self):
        item = {
            "codigo": "TEST01",
            "trayecto_titulo": "Prueba de Trayecto",
            "resolucion_jurisdiccional": "RESFC-2019-6908-GDEBA-DGCYE",
            "resolucion_validez_nacional": "RESOL-2022-2139-APN-ME",
            "modulos": [
                {
                    "Código de Planificación": "MOD1.1",
                    "Denominación del Módulo": "Módulo Uno",
                    "Hs. Cat. Módulo": "30",
                    "Hs. Reloj Módulo": "20"
                },
                {
                    "Código de Planificación": "MOD2.1",
                    "Denominación del Módulo": "Módulo Dos",
                    "Hs. Cat. Módulo": "60",
                    "Hs. Reloj Módulo": "40"
                }
            ]
        }
        formatted = format_trayecto_data(item)
        self.assertEqual(formatted["codigo"], "TEST01")
        self.assertEqual(formatted["horas_cursada"], "60")
        self.assertEqual(
            formatted["resolucion"],
            "Resolución Nro. RESFC-2019-6908-GDEBA-DGCYE / RESOL-2022-2139-APN-ME--"
        )
        self.assertEqual(len(formatted["modulos"]), 2)
        self.assertEqual(formatted["modulos"][0], "Módulo Uno. (MOD1.1)")

    def test_format_trayecto_data_jurisdictional_only(self):
        item = {
            "codigo": "TEST02",
            "trayecto_titulo": "Prueba Solo Jurisdiccional",
            "resolucion_jurisdiccional": "RESFC-2019-510-GDEBA-DGCYE",
            "modulos": []
        }
        formatted = format_trayecto_data(item)
        self.assertEqual(formatted["resolucion"], "Resolución Nro. RESFC-2019-510-GDEBA-DGCYE--")

    def test_search_trayecto(self):
        match = self.catalog.search_trayecto("Soldador Básico")
        self.assertIsNotNone(match)
        self.assertEqual(match.get("codigo"), "MT30")

if __name__ == "__main__":
    unittest.main()
