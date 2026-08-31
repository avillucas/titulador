import unittest
import datetime
from src.date_utils import (
    parse_date_string,
    add_one_month,
    calculate_default_emision_date,
    parse_emision_input,
    MESES_ESPANOL
)

class TestDateUtils(unittest.TestCase):
    def test_parse_date_string_spanish(self):
        result = parse_date_string("14 de Julio de 2025")
        self.assertEqual(result, datetime.date(2025, 7, 14))

        result = parse_date_string("01 de Enero de 2024")
        self.assertEqual(result, datetime.date(2024, 1, 1))

        result = parse_date_string("31 de Diciembre de 2023")
        self.assertEqual(result, datetime.date(2023, 12, 31))

    def test_parse_date_string_numeric(self):
        self.assertEqual(parse_date_string("14/07/2025"), datetime.date(2025, 7, 14))
        self.assertEqual(parse_date_string("14-07-2025"), datetime.date(2025, 7, 14))
        self.assertEqual(parse_date_string("2025-07-14"), datetime.date(2025, 7, 14))

    def test_add_one_month_standard(self):
        dt = datetime.date(2025, 7, 14)
        next_dt = add_one_month(dt)
        self.assertEqual(next_dt, datetime.date(2025, 8, 14))

    def test_add_one_month_end_of_year(self):
        dt = datetime.date(2025, 12, 15)
        next_dt = add_one_month(dt)
        self.assertEqual(next_dt, datetime.date(2026, 1, 15))

    def test_add_one_month_month_end_clipping(self):
        # Jan 31 + 1 month -> Feb 28 (non-leap year 2025)
        dt = datetime.date(2025, 1, 31)
        next_dt = add_one_month(dt)
        self.assertEqual(next_dt, datetime.date(2025, 2, 28))

        # Jan 31 + 1 month -> Feb 29 (leap year 2024)
        dt_leap = datetime.date(2024, 1, 31)
        next_dt_leap = add_one_month(dt_leap)
        self.assertEqual(next_dt_leap, datetime.date(2024, 2, 29))

    def test_calculate_default_emision_date(self):
        dia, mes, ano = calculate_default_emision_date("14 de Julio de 2025")
        self.assertEqual(dia, "14")
        self.assertEqual(mes, "Agosto")
        self.assertEqual(ano, "25")

        # Invalid date string should fallback safely
        dia_f, mes_f, ano_f = calculate_default_emision_date("fecha invalida")
        self.assertEqual((dia_f, mes_f, ano_f), ("02", "Septiembre", "25"))

    def test_parse_emision_input(self):
        self.assertEqual(parse_emision_input("20/10/2025"), ("20", "Octubre", "25"))
        self.assertEqual(parse_emision_input("5 de Mayo de 2026"), ("05", "Mayo", "26"))
        self.assertIsNone(parse_emision_input("invalido"))

if __name__ == "__main__":
    unittest.main()
