from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest

import xlwings as xw

this_dir = Path(__file__).parent

SOURCE_DATA = [
    ["Region", "Product", "Year", "Sales", "Qty"],
    ["North", "A", 2023, 100, 1],
    ["North", "B", 2023, 200, 2],
    ["South", "A", 2024, 300, 3],
    ["South", "B", 2024, 400, 4],
]


def _invalid(value: object) -> Any:
    """An argument that the Literal types reject statically: these tests exercise
    the runtime validation, which type checkers can't stand in for."""
    return value


@dataclass
class Fixture:
    book: xw.Book
    data: xw.Sheet
    sheet: xw.Sheet
    pt: xw.PivotTable | None = None


@pytest.fixture(scope="module")
def fx():
    """The existing pivot table in pivot_table.xlsx, so the tests run on macOS
    too, where pivot tables can't be created. Tests restore what they change."""
    app = xw.App(visible=False)
    book = app.books.open(this_dir / "pivot_table.xlsx")
    sheet = book.sheets["Pivot"]
    pt = sheet.pivot_tables["PivotTable1"]
    pt.refresh()
    yield Fixture(book=book, data=book.sheets["Data"], sheet=sheet, pt=pt)
    book.close()
    app.quit()


@pytest.fixture(scope="module")
def report():
    """A fresh book with source data on 'Data' and an empty 'Report' sheet, for
    the creation tests. Creating pivot tables isn't possible on macOS."""
    app = xw.App(visible=False)
    book = app.books.add()
    data = book.sheets[0]
    data.name = "Data"
    data["A1"].value = SOURCE_DATA
    report = book.sheets.add("Report", after=data)
    try:
        if sys.platform.startswith("darwin"):
            with pytest.raises(NotImplementedError):
                report.pivot_tables.add(data["A1"].expand(), report["A3"])
            pytest.skip("Creating pivot tables isn't supported on macOS")
        yield Fixture(book=book, data=data, sheet=report)
    finally:
        book.close()
        app.quit()


def test_collection(fx):
    pts = fx.sheet.pivot_tables
    assert len(pts) == 1
    assert pts.count == 1
    assert pts[0].name == "PivotTable1"
    assert pts(1).name == "PivotTable1"
    assert [pt.name for pt in pts] == ["PivotTable1"]
    assert "PivotTable1" in pts
    assert "nope" not in pts
    with pytest.raises(KeyError):
        pts["nope"]
    assert pts.parent == fx.sheet
    assert len(fx.data.pivot_tables) == 0
    assert (
        repr(fx.pt) == "<PivotTable 'PivotTable1' in <Sheet [pivot_table.xlsx]Pivot>>"
    )
    assert fx.pt == pts[0]


def test_api_parent(fx):
    assert fx.pt.api is not None
    assert fx.pt.parent == fx.sheet


def test_field_names(fx):
    assert fx.pt.field_names == ["Region", "Product", "Year", "Sales", "Qty"]
    # the "Values" pseudo field that appears with 2+ value fields is excluded,
    # from the areas too (Excel puts it into the columns area)
    qty = fx.pt.values.add("Qty")
    try:
        assert fx.pt.field_names == ["Region", "Product", "Year", "Sales", "Qty"]
        assert len(fx.pt.columns) == 0
        assert list(fx.pt.columns) == []
        assert [f.name for f in fx.pt.rows] == ["Region"]
    finally:
        qty.remove()


def test_name(fx):
    fx.pt.name = "MyPivot"
    try:
        assert fx.pt.name == "MyPivot"
        assert fx.sheet.pivot_tables["MyPivot"].name == "MyPivot"
    finally:
        fx.pt.name = "PivotTable1"
    assert fx.pt.name == "PivotTable1"


def test_rows(fx):
    rows = fx.pt.rows
    assert [f.name for f in rows] == ["Region"]
    assert len(rows) == 1
    assert rows[0].name == "Region"
    assert rows["Region"].name == "Region"
    assert rows(1).name == "Region"
    assert "Region" in rows
    assert "Year" not in rows
    with pytest.raises(KeyError):
        rows["Year"]
    assert rows.parent == fx.pt
    assert rows[0].parent == fx.pt
    assert rows[0].api is not None
    assert repr(rows[0]) == (
        "<PivotField 'Region' in <PivotTable 'PivotTable1' in "
        "<Sheet [pivot_table.xlsx]Pivot>>>"
    )


def test_fields_add_move_remove(fx):
    pt = fx.pt
    year = pt.columns.add("Year")
    assert year.name == "Year"
    assert [f.name for f in pt.columns] == ["Year"]
    product = pt.rows.add("Product")
    assert [f.name for f in pt.rows] == ["Region", "Product"]
    # already there: position is kept
    pt.rows.add("Region")
    assert [f.name for f in pt.rows] == ["Region", "Product"]
    # moving between areas appends
    pt.rows.add("Year")
    assert [f.name for f in pt.rows] == ["Region", "Product", "Year"]
    assert len(pt.columns) == 0
    pt.filters.add("Year")
    assert [f.name for f in pt.filters] == ["Year"]
    assert [f.name for f in pt.rows] == ["Region", "Product"]
    # a retained wrapper follows the field
    year.remove()
    assert len(pt.filters) == 0
    product.remove()
    assert [f.name for f in pt.rows] == ["Region"]
    with pytest.raises(KeyError):
        pt.rows.add("Nope")


def test_values(fx):
    values = fx.pt.values
    assert len(values) == 1
    assert [v.name for v in values] == ["Sum of Sales"]
    assert values[0].name == "Sum of Sales"
    assert values["Sum of Sales"].name == "Sum of Sales"
    assert "Sum of Sales" in values
    assert "Sum of Qty" not in values
    with pytest.raises(KeyError):
        values["Sum of Qty"]
    assert values.parent == fx.pt
    assert values[0].parent == fx.pt
    assert values[0].source_field == "Sales"
    assert values[0].function == "sum"
    assert values[0].number_format == "General"
    assert values[0].api is not None
    assert repr(values[0]) == (
        "<PivotValueField 'Sum of Sales' in <PivotTable 'PivotTable1' in "
        "<Sheet [pivot_table.xlsx]Pivot>>>"
    )


def test_values_add(fx):
    pt = fx.pt
    qty = pt.values.add("Qty")
    try:
        assert qty.name == "Sum of Qty"
        assert qty.source_field == "Qty"
        assert qty.function == "sum"
        assert [v.name for v in pt.values] == ["Sum of Sales", "Sum of Qty"]
        # with 2+ value fields, the report gains a "Values" caption row
        assert pt.range.value[1] == ["Row Labels", "Sum of Sales", "Sum of Qty"]
    finally:
        qty.remove()
    assert len(pt.values) == 1
    full = pt.values.add(
        "Sales", function="count", name="Sales Count", number_format="0.0"
    )
    try:
        assert full.name == "Sales Count"
        assert full.source_field == "Sales"
        assert full.function == "count"
        assert full.number_format == "0.0"
        assert pt.values["Sales Count"].function == "count"
        assert pt.values[1].name == "Sales Count"
        # the same source field twice
        again = pt.values.add("Sales", function="average")
        try:
            assert len(pt.values) == 3
            assert again.function == "average"
            assert again.source_field == "Sales"
        finally:
            again.remove()
    finally:
        full.remove()
    assert [v.name for v in pt.values] == ["Sum of Sales"]
    with pytest.raises(KeyError):
        pt.values.add("Nope")
    with pytest.raises(ValueError):
        pt.values.add("Qty", function=_invalid("total"))


def test_value_field_setters(fx):
    field = fx.pt.values.add("Qty")
    try:
        field.function = "average"
        assert field.function == "average"
        # Excel renames the automatic caption along with the function
        assert field.name == "Average of Qty"
        assert fx.pt.values[1].name == "Average of Qty"
        field.number_format = "#,##0.00"
        assert field.number_format == "#,##0.00"
        field.name = "Qty Average"
        assert field.name == "Qty Average"
        assert fx.pt.values["Qty Average"].function == "average"
        field.function = "max"
        assert field.function == "max"
        with pytest.raises(ValueError):
            field.function = _invalid("total")
    finally:
        field.remove()
    assert len(fx.pt.values) == 1


def test_layout(fx):
    pt = fx.pt
    assert pt.layout == "compact"
    try:
        pt.layout = "tabular"
        assert pt.layout == "tabular"
        pt.layout = "outline"
        assert pt.layout == "outline"
    finally:
        pt.layout = "compact"
    assert pt.layout == "compact"
    with pytest.raises(ValueError):
        pt.layout = _invalid("fancy")


def test_value_field_aliases(fx):
    field = fx.pt.values.add("Qty")
    # Reacquire the pivot and field through different collection lookups.
    other = fx.sheet.pivot_tables["PivotTable1"].values["Sum of Qty"]
    try:
        other.function = "average"
        assert field.name == "Average of Qty"
        field.number_format = "0.00"
        assert other.number_format == "0.00"
        other.name = "Quantity average"
        assert field.name == "Quantity average"
        field.function = "max"
        assert other.function == "max"
    finally:
        field.remove()
    assert [v.name for v in fx.pt.values] == ["Sum of Sales"]


@pytest.mark.skipif(sys.platform != "darwin", reason="macOS alias lifetime")
def test_removed_value_alias_does_not_target_reused_caption(fx):
    field = fx.pt.values.add("Qty", name="Quantity")
    alias = fx.sheet.pivot_tables[0].values["Quantity"]
    field.remove()
    replacement = fx.pt.values.add("Qty", name="Quantity")
    try:
        with pytest.raises(KeyError):
            alias.remove()
        replacement.number_format = "0.00"
        assert fx.pt.values["Quantity"].number_format == "0.00"
    finally:
        replacement.remove()


def test_grand_totals(fx):
    pt = fx.pt
    assert pt.show_row_grand_totals
    assert pt.show_column_grand_totals
    try:
        pt.show_row_grand_totals = False
        assert not pt.show_row_grand_totals
        # the "Grand Total" row at the bottom belongs to the column totals
        pt.show_column_grand_totals = False
        assert not pt.show_column_grand_totals
        assert pt.range.value[-1][0] == "South"
    finally:
        pt.show_row_grand_totals = True
        pt.show_column_grand_totals = True
    assert pt.range.value[-1][0] == "Grand Total"


def test_ranges(fx):
    pt = fx.pt
    assert pt.range.address == "$A$3:$B$6"
    assert pt.range.value == [
        ["Row Labels", "Sum of Sales"],
        ["North", 300.0],
        ["South", 700.0],
        ["Grand Total", 1000.0],
    ]
    body = pt.data_body_range
    assert body is not None
    assert body.address == "$B$4:$B$6"
    assert body.sheet == fx.sheet


def test_data_body_range_without_values(fx):
    pt = fx.pt
    pt.values[0].remove()
    try:
        assert len(pt.values) == 0
        assert pt.data_body_range is None
    finally:
        pt.values.add("Sales", function="sum")
    assert pt.values[0].name == "Sum of Sales"


def test_refresh(fx):
    pt = fx.pt
    fx.data["D2"].value = 150
    try:
        pt.refresh()
        assert pt.range.value[1] == ["North", 350.0]
    finally:
        fx.data["D2"].value = 100
        pt.refresh()
    assert pt.range.value[1] == ["North", 300.0]


def test_add_validation_errors(fx):
    # rejected before any engine call, so this runs on macOS too
    source = fx.data["A1"].expand()
    with pytest.raises(ValueError):
        fx.sheet.pivot_tables.add(source, fx.data["H1"])
    with pytest.raises(TypeError):
        fx.sheet.pivot_tables.add(_invalid("A1:E5"), fx.sheet["A30"])
    with pytest.raises(ValueError):
        fx.sheet.pivot_tables.add(source, fx.sheet["A30"], layout=_invalid("wide"))
    with pytest.raises(ValueError):
        fx.sheet.pivot_tables.add(
            source, fx.sheet["A30"], values=_invalid({"Sales": "total"})
        )
    with pytest.raises(TypeError):
        fx.sheet.pivot_tables.add(source, fx.sheet["A30"], rows=_invalid([1]))
    assert len(fx.sheet.pivot_tables) == 1


def test_delete():
    # its own book: the module fixture keeps its pivot table
    app = xw.App(visible=False)
    try:
        book = app.books.open(this_dir / "pivot_table.xlsx")
        sheet = book.sheets["Pivot"]
        pt = sheet.pivot_tables[0]
        pt.refresh()
        pt.delete()
        assert len(sheet.pivot_tables) == 0
        assert sheet["A3"].value is None
        book.close()
    finally:
        app.quit()


def test_add_one_shot(report):
    pt = report.sheet.pivot_tables.add(
        source=report.data["A1"].expand(),
        destination=report.sheet["A3"],
        name="OneShot",
        rows=["Region", "Product"],
        columns="Year",
        filters="Product",
        values={"Sales": "sum", "Qty": "count"},
        layout="tabular",
    )
    assert pt.name == "OneShot"
    assert report.sheet.pivot_tables["OneShot"].name == "OneShot"
    assert pt.parent == report.sheet
    assert pt.field_names == ["Region", "Product", "Year", "Sales", "Qty"]
    # "Product" was moved from rows to filters
    assert [f.name for f in pt.rows] == ["Region"]
    assert [f.name for f in pt.columns] == ["Year"]
    assert [f.name for f in pt.filters] == ["Product"]
    assert [v.name for v in pt.values] == ["Sum of Sales", "Count of Qty"]
    assert [v.function for v in pt.values] == ["sum", "count"]
    assert pt.layout == "tabular"
    # the filters area goes above the destination cell, which stays the
    # top-left of the report body; range excludes the filters area
    assert pt.range.address[:4] == "$A$3"
    # last row: Grand Total, per-year values, then the overall totals of
    # both value fields
    assert pt.range.value[-1][0] == "Grand Total"
    assert pt.range.value[-1][-2:] == [1000.0, 4.0]


def test_add_steps_and_table_source(report):
    table = report.data.tables.add(report.data["A1"].expand(), name="SourceTable")
    pt = report.sheet.pivot_tables.add(source=table, destination=report.sheet["J3"])
    assert pt.name.startswith("PivotTable")
    assert len(pt.rows) == 0
    assert len(pt.values) == 0
    assert pt.data_body_range is None
    pt.rows.add("Region")
    pt.values.add("Sales", function="sum", name="Total", number_format="#,##0")
    pt.values.add("Sales", function="count")
    assert [v.name for v in pt.values] == ["Total", "Count of Sales"]
    assert pt.values["Total"].number_format == "#,##0"
    assert pt.range.value == [
        ["Row Labels", "Total", "Count of Sales"],
        ["North", 300.0, 2.0],
        ["South", 700.0, 2.0],
        ["Grand Total", 1000.0, 4.0],
    ]
    name = pt.name
    pt.delete()
    assert name not in [p.name for p in report.sheet.pivot_tables]


def test_add_duplicate_name(report):
    source = report.data["A1"].expand()
    pt = report.sheet.pivot_tables.add(source, report.sheet["A30"], name="Dup")
    with pytest.raises(xw.XlwingsError):
        report.sheet.pivot_tables.add(source, report.sheet["A40"], name="Dup")
    pt.delete()
