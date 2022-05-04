import unittest
from pathlib import Path

import xlwings as xw
import pandas as pd


class TestTable(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.book = xw.Book()
        cls.sheet = cls.book.sheets[0]
        cls.sheet["A1"].value = [["a", "b"], [1, 2]]
        cls.test_table = cls.sheet.tables.add(source=cls.sheet["A1"].expand())

    @classmethod
    def tearDownClass(cls):
        cls.book.close()

    def test_add_table_no_name(self):
        self.assertEqual(self.sheet.tables[0].name, "Table1")

    def test_add_table_with_name(self):
        self.sheet["A4"].value = [["a", "b"], [1, 2]]
        self.sheet.tables.add(source=self.sheet["A4"].expand(), name="AABBCC")
        self.assertEqual(self.sheet.tables["AABBCC"].name, "AABBCC")

    def test_data_body_range(self):
        self.assertEqual(self.test_table.data_body_range, self.sheet["A2:B2"])

    def test_display_name(self):
        origin_display_name = self.test_table.display_name
        self.test_table.display_name = "ABCDE"
        self.assertEqual(self.test_table.display_name, "ABCDE")
        self.test_table.display_name = origin_display_name

    def test_header_row_range(self):
        self.assertEqual(self.test_table.header_row_range, self.sheet["A1:B1"])
        self.test_table.show_headers = False
        self.assertIsNone(self.test_table.header_row_range)
        self.test_table.show_headers = True

    def test_insert_row_range(self):
        table = self.sheet.tables.add(self.sheet["A10"])
        self.assertEqual(table.insert_row_range, self.sheet["A11"])

    def test_insert_row_range_none(self):
        self.assertIsNone(self.test_table.insert_row_range)

    def test_name(self):
        original_name = self.test_table.name
        self.test_table.name = "XYZ"
        self.assertEqual(self.test_table.name, "XYZ")
        self.assertEqual(self.sheet.tables["XYZ"].name, "XYZ")
        self.test_table.name = original_name

    def test_parent(self):
        self.assertEqual(self.test_table.parent, self.sheet)

    def test_show_autofilter(self):
        self.assertTrue(self.test_table.show_autofilter)
        self.test_table.show_autofilter = False
        self.assertFalse(self.test_table.show_autofilter)
        self.test_table.show_autofilter = True

    def test_show_headers(self):
        self.assertTrue(self.test_table.show_headers)
        self.test_table.show_headers = False
        self.assertFalse(self.test_table.show_headers)
        self.test_table.show_headers = True

    def test_show_table_style_columns_stripes(self):
        self.assertFalse(self.test_table.show_table_style_column_stripes)
        self.test_table.show_table_style_column_stripes = True
        self.assertTrue(self.test_table.show_table_style_column_stripes)
        self.test_table.show_table_style_column_stripes = False

    def test_show_table_style_first_column(self):
        self.assertFalse(self.test_table.show_table_style_first_column)
        self.test_table.show_table_style_first_column = True
        self.assertTrue(self.test_table.show_table_style_first_column)
        self.test_table.show_table_style_first_column = False

    def test_show_table_style_last_column(self):
        self.assertFalse(self.test_table.show_table_style_last_column)
        self.test_table.show_table_style_last_column = True
        self.assertTrue(self.test_table.show_table_style_last_column)
        self.test_table.show_table_style_last_column = False

    def test_show_table_style_row_stripes(self):
        self.assertTrue(self.test_table.show_table_style_row_stripes)
        self.test_table.show_table_style_row_stripes = False
        self.assertFalse(self.test_table.show_table_style_row_stripes)
        self.test_table.show_table_style_row_stripes = True

    def test_show_totals(self):
        self.assertFalse(self.test_table.show_totals)
        self.test_table.show_totals = True
        self.assertTrue(self.test_table.show_totals)
        self.test_table.show_totals = False

    def test_table_style(self):
        self.assertEqual(self.test_table.table_style, "TableStyleMedium2")
        self.test_table.table_style = "TableStyleMedium1"
        self.assertEqual(self.test_table.table_style, "TableStyleMedium1")
        self.test_table.table_style = "TableStyleMedium2"

    def test_totals_row_range(self):
        self.assertIsNone(self.test_table.totals_row_range)
        self.test_table.show_totals = True
        self.assertEqual(self.test_table.totals_row_range, self.sheet["A3:B3"])
        self.test_table.show_totals = False

    def test_resize(self):
        self.assertEqual(self.test_table.range.address, "$A$1:$B$2")
        self.test_table.resize(self.sheet["A1:C3"])
        self.assertEqual(self.test_table.range.address, "$A$1:$C$3")
        self.test_table.resize(self.sheet["$A$1:$B$2"])
        self.assertEqual(self.test_table.range.address, "$A$1:$B$2")


class TestTableUpdate(unittest.TestCase):
    def test_table_update(self):
        df = pd.DataFrame(
            {
                "a": [1, 2, 3, 4, 5],
                "b": [11, 22, 33, 44, 55],
                "c": [111, 222, 333, 444, 555],
                "d": [1111, 2222, 3333, 4444, 5555],
            }
        )
        book = xw.Book(Path("tables.xlsx").resolve())
        sheet = book.sheets["template"].copy()
        sheet.tables[0].update(df)
        sheet.tables[1].update(df)
        sheet.tables[2].update(df)
        sheet.tables[3].update(df, index=False)
        self.assertEqual(sheet["A1:E50"].value, book.sheets["expected"]["A1:E50"].value)
        sheet.book.close()


if __name__ == "__main__":
    unittest.main()
