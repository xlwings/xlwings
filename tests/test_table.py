import asyncio
import unittest
from datetime import date, datetime
from pathlib import Path

import pandas as pd

import xlwings as xw

this_dir = Path(__file__).parent


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

    def test_autofilter_apply_and_clear(self):
        book = xw.Book()
        sheet = book.sheets[0]
        values = [
            ["Region", "Amount", "Status", "When"],
            ["East", 5, "Open", datetime(2026, 1, 1)],
            ["West", 10, None, datetime(2026, 2, 1)],
            ["North", 15, "Closed", datetime(2026, 3, 1)],
            ["East", 20, None, datetime(2026, 4, 1)],
        ]
        try:
            target = sheet["A1:D5"]
            target.value = values
            formulas = target.formula
            table = sheet.tables.add(target)
            table.autofilter.apply_values(1, ["East", "West"])
            table.autofilter.apply_comparison(2, "greater_than_or_equal", 10)
            table.autofilter.apply_comparison(3, "not_equal_to", None)
            table.autofilter.apply_comparison(4, "less_than", date(2026, 4, 1))
            table.autofilter.apply_bottom_percent(2, 50)
            criteria = table.autofilter.criteria
            self.assertEqual(criteria[0].type, "values")
            self.assertEqual(criteria[1].type, "bottom_percent")
            self.assertIn(criteria[1].percent, (None, 50))
            self.assertEqual(criteria[2].operator, "not_equal_to")
            self.assertEqual(criteria[3].operator, "less_than")
            self.assertEqual(target.value, values)
            self.assertEqual(target.formula, formulas)
            table.autofilter.clear(2)
            table.autofilter.clear()
            self.assertEqual(target.value, values)
        finally:
            book.close()

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
        book = xw.Book(this_dir / "tables.xlsx")
        sheet = book.sheets["template"].copy()
        sheet.tables[0].update(df)
        sheet.tables[1].update(df)
        sheet.tables[2].update(df)
        sheet.tables[3].update(df, index=False)
        self.assertEqual(sheet["A1:E50"].value, book.sheets["expected"]["A1:E50"].value)
        sheet.book.close()


class TestTableRows(unittest.TestCase):
    """Integration checks against desktop Excel on Windows and macOS."""

    def setUp(self):
        self.book = xw.Book()
        self.sheet = self.book.sheets[0]
        self.sheet["B2:C4"].value = [["Item", "Amount"], ["first", 10], ["second", 20]]
        self.table = self.sheet.tables.add(self.sheet["B2:C4"], name="RowStructureTest")

    def tearDown(self):
        self.book.close()

    def test_append_insert_delete_and_readback(self):
        self.sheet["A3"].value = "left"
        self.sheet["E3"].value = "right"
        self.sheet["C3"].formula = "=5*2"
        original_style = self.table.table_style
        rows = self.table.rows
        self.assertEqual(len(rows), 2)
        self.assertEqual(asyncio.run(rows.get_count()), 2)
        self.assertEqual(rows[0].range.address, "$B$3:$C$3")
        self.assertEqual(asyncio.run(rows[0].get_range()).address, "$B$3:$C$3")

        appended = rows.add(["third", 30])
        self.assertEqual(appended.index, 3)
        rows.add(["middle", 15], index=2)
        self.assertEqual(len(rows), 4)
        self.assertEqual(
            self.table.data_body_range.value,
            [["first", 10], ["middle", 15], ["second", 20], ["third", 30]],
        )
        self.assertEqual(self.sheet["C3"].formula, "=5*2")
        self.assertEqual(self.sheet["A3"].value, "left")
        self.assertEqual(self.sheet["E3"].value, "right")

        rows[1].delete()
        self.assertEqual(
            self.table.data_body_range.value,
            [["first", 10], ["second", 20], ["third", 30]],
        )
        self.assertEqual(self.table.table_style, original_style)

    def test_refuses_to_shift_neighbors(self):
        self.sheet["B6"].value = "keep"
        with self.assertRaisesRegex(ValueError, "below the table"):
            self.table.rows.add(["third", 30])
        with self.assertRaisesRegex(ValueError, "below the table"):
            self.table.rows[0].delete()
        self.assertEqual(self.sheet["B6"].value, "keep")
        self.assertEqual(len(self.table.rows), 2)

    def test_empty_table_and_totals_row(self):
        self.table.show_totals = True
        initial_style = self.table.table_style
        self.table.rows.add(["third", 30])
        self.assertEqual(len(self.table.rows), 3)
        self.assertTrue(self.table.show_totals)
        self.assertIsNotNone(self.table.totals_row_range)
        self.table.rows[1].delete()
        self.assertTrue(self.table.show_totals)
        self.assertEqual(self.table.table_style, initial_style)
        self.assertEqual(len(self.table.rows), 2)

        empty_sheet = self.book.sheets.add("EmptyRows")
        empty = empty_sheet.tables.add(empty_sheet["G2:H2"], name="EmptyRowsTest")
        self.assertEqual(len(empty.rows), 0)
        empty.rows.add(["only", 1])
        self.assertEqual(len(empty.rows), 1)
        self.assertEqual(empty.data_body_range.value, ["only", 1])


if __name__ == "__main__":
    unittest.main()
