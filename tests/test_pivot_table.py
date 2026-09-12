import sys
import unittest
from pathlib import Path
from typing import Any

import xlwings as xw

this_dir = Path(__file__).parent


def _invalid(value: object) -> Any:
    """An argument that the Literal types reject statically: these tests exercise
    the runtime validation, which type checkers can't stand in for."""
    return value


class TestPivotTable(unittest.TestCase):
    """Works with the existing pivot table in pivot_table.xlsx, so it runs on
    macOS too, where pivot tables can't be created."""

    @classmethod
    def setUpClass(cls):
        cls.app = xw.App(visible=False)
        cls.book = cls.app.books.open(this_dir / "pivot_table.xlsx")
        cls.data = cls.book.sheets["Data"]
        cls.sheet = cls.book.sheets["Pivot"]
        cls.pt = cls.sheet.pivot_tables["PivotTable1"]
        cls.pt.refresh()

    @classmethod
    def tearDownClass(cls):
        cls.book.close()
        cls.app.quit()

    def test_collection(self):
        pts = self.sheet.pivot_tables
        self.assertEqual(len(pts), 1)
        self.assertEqual(pts.count, 1)
        self.assertEqual(pts[0].name, "PivotTable1")
        self.assertEqual(pts(1).name, "PivotTable1")
        self.assertEqual([pt.name for pt in pts], ["PivotTable1"])
        self.assertIn("PivotTable1", pts)
        self.assertNotIn("nope", pts)
        with self.assertRaises(KeyError):
            pts["nope"]
        self.assertEqual(pts.parent, self.sheet)
        self.assertEqual(len(self.data.pivot_tables), 0)
        self.assertEqual(
            repr(self.pt),
            "<PivotTable 'PivotTable1' in <Sheet [pivot_table.xlsx]Pivot>>",
        )
        self.assertEqual(self.pt, pts[0])

    def test_api_parent(self):
        self.assertIsNotNone(self.pt.api)
        self.assertEqual(self.pt.parent, self.sheet)

    def test_field_names(self):
        self.assertEqual(
            self.pt.field_names, ["Region", "Product", "Year", "Sales", "Qty"]
        )
        # the "Values" pseudo field that appears with 2+ value fields is excluded,
        # from the areas too (Excel puts it into the columns area)
        qty = self.pt.values.add("Qty")
        try:
            self.assertEqual(
                self.pt.field_names, ["Region", "Product", "Year", "Sales", "Qty"]
            )
            self.assertEqual(len(self.pt.columns), 0)
            self.assertEqual(list(self.pt.columns), [])
            self.assertEqual([f.name for f in self.pt.rows], ["Region"])
        finally:
            qty.remove()

    def test_name(self):
        self.pt.name = "MyPivot"
        try:
            self.assertEqual(self.pt.name, "MyPivot")
            self.assertEqual(self.sheet.pivot_tables["MyPivot"].name, "MyPivot")
        finally:
            self.pt.name = "PivotTable1"
        self.assertEqual(self.pt.name, "PivotTable1")

    def test_rows(self):
        rows = self.pt.rows
        self.assertEqual([f.name for f in rows], ["Region"])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].name, "Region")
        self.assertEqual(rows["Region"].name, "Region")
        self.assertEqual(rows(1).name, "Region")
        self.assertIn("Region", rows)
        self.assertNotIn("Year", rows)
        with self.assertRaises(KeyError):
            rows["Year"]
        self.assertEqual(rows.parent, self.pt)
        self.assertEqual(rows[0].parent, self.pt)
        self.assertIsNotNone(rows[0].api)
        self.assertEqual(
            repr(rows[0]),
            "<PivotField 'Region' in <PivotTable 'PivotTable1' in "
            "<Sheet [pivot_table.xlsx]Pivot>>>",
        )

    def test_fields_add_move_remove(self):
        pt = self.pt
        year = pt.columns.add("Year")
        self.assertEqual(year.name, "Year")
        self.assertEqual([f.name for f in pt.columns], ["Year"])
        product = pt.rows.add("Product")
        self.assertEqual([f.name for f in pt.rows], ["Region", "Product"])
        # already there: position is kept
        pt.rows.add("Region")
        self.assertEqual([f.name for f in pt.rows], ["Region", "Product"])
        # moving between areas appends
        pt.rows.add("Year")
        self.assertEqual([f.name for f in pt.rows], ["Region", "Product", "Year"])
        self.assertEqual(len(pt.columns), 0)
        pt.filters.add("Year")
        self.assertEqual([f.name for f in pt.filters], ["Year"])
        self.assertEqual([f.name for f in pt.rows], ["Region", "Product"])
        # a retained wrapper follows the field
        year.remove()
        self.assertEqual(len(pt.filters), 0)
        product.remove()
        self.assertEqual([f.name for f in pt.rows], ["Region"])
        with self.assertRaises(KeyError):
            pt.rows.add("Nope")

    def test_values(self):
        values = self.pt.values
        self.assertEqual(len(values), 1)
        self.assertEqual([v.name for v in values], ["Sum of Sales"])
        self.assertEqual(values[0].name, "Sum of Sales")
        self.assertEqual(values["Sum of Sales"].name, "Sum of Sales")
        self.assertIn("Sum of Sales", values)
        self.assertNotIn("Sum of Qty", values)
        with self.assertRaises(KeyError):
            values["Sum of Qty"]
        self.assertEqual(values.parent, self.pt)
        self.assertEqual(values[0].parent, self.pt)
        self.assertEqual(values[0].source_field, "Sales")
        self.assertEqual(values[0].function, "sum")
        self.assertEqual(values[0].number_format, "General")
        self.assertIsNotNone(values[0].api)
        self.assertEqual(
            repr(values[0]),
            "<PivotValueField 'Sum of Sales' in <PivotTable 'PivotTable1' in "
            "<Sheet [pivot_table.xlsx]Pivot>>>",
        )

    def test_values_add(self):
        pt = self.pt
        qty = pt.values.add("Qty")
        try:
            self.assertEqual(qty.name, "Sum of Qty")
            self.assertEqual(qty.source_field, "Qty")
            self.assertEqual(qty.function, "sum")
            self.assertEqual(
                [v.name for v in pt.values], ["Sum of Sales", "Sum of Qty"]
            )
            # with 2+ value fields, the report gains a "Values" caption row
            self.assertEqual(
                pt.range.value[1], ["Row Labels", "Sum of Sales", "Sum of Qty"]
            )
        finally:
            qty.remove()
        self.assertEqual(len(pt.values), 1)
        full = pt.values.add(
            "Sales", function="count", name="Sales Count", number_format="0.0"
        )
        try:
            self.assertEqual(full.name, "Sales Count")
            self.assertEqual(full.source_field, "Sales")
            self.assertEqual(full.function, "count")
            self.assertEqual(full.number_format, "0.0")
            self.assertEqual(pt.values["Sales Count"].function, "count")
            self.assertEqual(pt.values[1].name, "Sales Count")
            # the same source field twice
            again = pt.values.add("Sales", function="average")
            try:
                self.assertEqual(len(pt.values), 3)
                self.assertEqual(again.function, "average")
                self.assertEqual(again.source_field, "Sales")
            finally:
                again.remove()
        finally:
            full.remove()
        self.assertEqual([v.name for v in pt.values], ["Sum of Sales"])
        with self.assertRaises(KeyError):
            pt.values.add("Nope")
        with self.assertRaises(ValueError):
            pt.values.add("Qty", function=_invalid("total"))

    def test_value_field_setters(self):
        field = self.pt.values.add("Qty")
        try:
            field.function = "average"
            self.assertEqual(field.function, "average")
            # Excel renames the automatic caption along with the function
            self.assertEqual(field.name, "Average of Qty")
            self.assertEqual(self.pt.values[1].name, "Average of Qty")
            field.number_format = "#,##0.00"
            self.assertEqual(field.number_format, "#,##0.00")
            field.name = "Qty Average"
            self.assertEqual(field.name, "Qty Average")
            self.assertEqual(self.pt.values["Qty Average"].function, "average")
            field.function = "max"
            self.assertEqual(field.function, "max")
            with self.assertRaises(ValueError):
                field.function = _invalid("total")
        finally:
            field.remove()
        self.assertEqual(len(self.pt.values), 1)

    def test_layout(self):
        pt = self.pt
        self.assertEqual(pt.layout, "compact")
        try:
            pt.layout = "tabular"
            self.assertEqual(pt.layout, "tabular")
            pt.layout = "outline"
            self.assertEqual(pt.layout, "outline")
        finally:
            pt.layout = "compact"
        self.assertEqual(pt.layout, "compact")
        with self.assertRaises(ValueError):
            pt.layout = _invalid("fancy")

    def test_value_field_aliases(self):
        field = self.pt.values.add("Qty")
        # Reacquire the pivot and field through different collection lookups.
        other = self.sheet.pivot_tables["PivotTable1"].values["Sum of Qty"]
        try:
            other.function = "average"
            self.assertEqual(field.name, "Average of Qty")
            field.number_format = "0.00"
            self.assertEqual(other.number_format, "0.00")
            other.name = "Quantity average"
            self.assertEqual(field.name, "Quantity average")
            field.function = "max"
            self.assertEqual(other.function, "max")
        finally:
            field.remove()
        self.assertEqual([v.name for v in self.pt.values], ["Sum of Sales"])

    @unittest.skipUnless(sys.platform == "darwin", "macOS alias lifetime")
    def test_removed_value_alias_does_not_target_reused_caption(self):
        field = self.pt.values.add("Qty", name="Quantity")
        alias = self.sheet.pivot_tables[0].values["Quantity"]
        field.remove()
        replacement = self.pt.values.add("Qty", name="Quantity")
        try:
            with self.assertRaises(KeyError):
                alias.remove()
            replacement.number_format = "0.00"
            self.assertEqual(self.pt.values["Quantity"].number_format, "0.00")
        finally:
            replacement.remove()

    def test_grand_totals(self):
        pt = self.pt
        self.assertTrue(pt.show_row_grand_totals)
        self.assertTrue(pt.show_column_grand_totals)
        try:
            pt.show_row_grand_totals = False
            self.assertFalse(pt.show_row_grand_totals)
            # the "Grand Total" row at the bottom belongs to the column totals
            pt.show_column_grand_totals = False
            self.assertFalse(pt.show_column_grand_totals)
            self.assertEqual(pt.range.value[-1][0], "South")
        finally:
            pt.show_row_grand_totals = True
            pt.show_column_grand_totals = True
        self.assertEqual(pt.range.value[-1][0], "Grand Total")

    def test_ranges(self):
        pt = self.pt
        self.assertEqual(pt.range.address, "$A$3:$B$6")
        self.assertEqual(
            pt.range.value,
            [
                ["Row Labels", "Sum of Sales"],
                ["North", 300.0],
                ["South", 700.0],
                ["Grand Total", 1000.0],
            ],
        )
        body = pt.data_body_range
        assert body is not None
        self.assertEqual(body.address, "$B$4:$B$6")
        self.assertEqual(body.sheet, self.sheet)

    def test_data_body_range_without_values(self):
        pt = self.pt
        pt.values[0].remove()
        try:
            self.assertEqual(len(pt.values), 0)
            self.assertIsNone(pt.data_body_range)
        finally:
            pt.values.add("Sales", function="sum")
        self.assertEqual(pt.values[0].name, "Sum of Sales")

    def test_refresh(self):
        pt = self.pt
        self.data["D2"].value = 150
        try:
            pt.refresh()
            self.assertEqual(pt.range.value[1], ["North", 350.0])
        finally:
            self.data["D2"].value = 100
            pt.refresh()
        self.assertEqual(pt.range.value[1], ["North", 300.0])

    def test_add_validation_errors(self):
        # rejected before any engine call, so this runs on macOS too
        source = self.data["A1"].expand()
        with self.assertRaises(ValueError):
            self.sheet.pivot_tables.add(source, self.data["H1"])
        with self.assertRaises(TypeError):
            self.sheet.pivot_tables.add(_invalid("A1:E5"), self.sheet["A30"])
        with self.assertRaises(ValueError):
            self.sheet.pivot_tables.add(
                source, self.sheet["A30"], layout=_invalid("wide")
            )
        with self.assertRaises(ValueError):
            self.sheet.pivot_tables.add(
                source, self.sheet["A30"], values=_invalid({"Sales": "total"})
            )
        with self.assertRaises(TypeError):
            self.sheet.pivot_tables.add(source, self.sheet["A30"], rows=_invalid([1]))
        self.assertEqual(len(self.sheet.pivot_tables), 1)


class TestPivotTableDelete(unittest.TestCase):
    def test_delete(self):
        app = xw.App(visible=False)
        try:
            book = app.books.open(this_dir / "pivot_table.xlsx")
            sheet = book.sheets["Pivot"]
            pt = sheet.pivot_tables[0]
            pt.refresh()
            pt.delete()
            self.assertEqual(len(sheet.pivot_tables), 0)
            self.assertEqual(sheet["A3"].value, None)
            book.close()
        finally:
            app.quit()


class TestPivotTablesAdd(unittest.TestCase):
    """Creating pivot tables isn't possible on macOS."""

    @classmethod
    def setUpClass(cls):
        cls.app = xw.App(visible=False)
        cls.book = cls.app.books.add()
        cls.data = cls.book.sheets[0]
        cls.data.name = "Data"
        cls.data["A1"].value = [
            ["Region", "Product", "Year", "Sales", "Qty"],
            ["North", "A", 2023, 100, 1],
            ["North", "B", 2023, 200, 2],
            ["South", "A", 2024, 300, 3],
            ["South", "B", 2024, 400, 4],
        ]
        cls.report = cls.book.sheets.add("Report", after=cls.data)

    @classmethod
    def tearDownClass(cls):
        cls.book.close()
        cls.app.quit()

    def setUp(self):
        if sys.platform.startswith("darwin"):
            with self.assertRaises(NotImplementedError):
                self.report.pivot_tables.add(
                    self.data["A1"].expand(), self.report["A3"]
                )
            self.skipTest("Creating pivot tables isn't supported on macOS")

    def test_add_one_shot(self):
        pt = self.report.pivot_tables.add(
            source=self.data["A1"].expand(),
            destination=self.report["A3"],
            name="OneShot",
            rows=["Region", "Product"],
            columns="Year",
            filters="Product",
            values={"Sales": "sum", "Qty": "count"},
            layout="tabular",
        )
        self.assertEqual(pt.name, "OneShot")
        self.assertEqual(self.report.pivot_tables["OneShot"].name, "OneShot")
        self.assertEqual(pt.parent, self.report)
        self.assertEqual(pt.field_names, ["Region", "Product", "Year", "Sales", "Qty"])
        # "Product" was moved from rows to filters
        self.assertEqual([f.name for f in pt.rows], ["Region"])
        self.assertEqual([f.name for f in pt.columns], ["Year"])
        self.assertEqual([f.name for f in pt.filters], ["Product"])
        self.assertEqual([v.name for v in pt.values], ["Sum of Sales", "Count of Qty"])
        self.assertEqual([v.function for v in pt.values], ["sum", "count"])
        self.assertEqual(pt.layout, "tabular")
        # the filters area goes above the destination cell, which stays the
        # top-left of the report body; range excludes the filters area
        self.assertEqual(pt.range.address[:4], "$A$3")
        # last row: Grand Total, per-year values, then the overall totals of
        # both value fields
        self.assertEqual(pt.range.value[-1][0], "Grand Total")
        self.assertEqual(pt.range.value[-1][-2:], [1000.0, 4.0])

    def test_add_steps_and_table_source(self):
        table = self.data.tables.add(self.data["A1"].expand(), name="SourceTable")
        pt = self.report.pivot_tables.add(source=table, destination=self.report["J3"])
        self.assertEqual(pt.name[: len("PivotTable")], "PivotTable")
        self.assertEqual(len(pt.rows), 0)
        self.assertEqual(len(pt.values), 0)
        self.assertIsNone(pt.data_body_range)
        pt.rows.add("Region")
        pt.values.add("Sales", function="sum", name="Total", number_format="#,##0")
        pt.values.add("Sales", function="count")
        self.assertEqual([v.name for v in pt.values], ["Total", "Count of Sales"])
        self.assertEqual(pt.values["Total"].number_format, "#,##0")
        self.assertEqual(
            pt.range.value,
            [
                ["Row Labels", "Total", "Count of Sales"],
                ["North", 300.0, 2.0],
                ["South", 700.0, 2.0],
                ["Grand Total", 1000.0, 4.0],
            ],
        )
        name = pt.name
        pt.delete()
        self.assertNotIn(name, [p.name for p in self.report.pivot_tables])

    def test_add_duplicate_name(self):
        source = self.data["A1"].expand()
        pt = self.report.pivot_tables.add(source, self.report["A30"], name="Dup")
        with self.assertRaises(xw.XlwingsError):
            self.report.pivot_tables.add(source, self.report["A40"], name="Dup")
        pt.delete()


if __name__ == "__main__":
    unittest.main()
