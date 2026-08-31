import asyncio
import datetime as dt
import json
import os
import re
from pathlib import Path
from unittest import mock

import pytest

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    from dateutil import tz
except ImportError:
    tz = None

import xlwings as xw

this_dir = Path(__file__).resolve().parent

# "calamine", "remote", or "excel"
engine = os.environ.get("XLWINGS_ENGINE") or "remote"
# "xlsm", "xlsb", or "xls"
file_extension = os.environ.get("XLWINGS_FILE_EXTENSION") or "xlsm"


data = {
    "client": "Microsoft Office Scripts",
    "version": xw.__version__,
    "book": {
        "name": f"engines.{file_extension}",
        "active_sheet_index": 0,
        "selection": "B3:B4",
        "calculation": "Automatic",
    },
    "names": [
        {
            "name": "one",
            "sheet_index": 0,
            "address": "A1",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
        {
            "name": "two",  # VBA/GS send: "'Sheet 1'!two"
            "sheet_index": 0,
            "address": "C7:D8",
            "scope_sheet_name": "Sheet 1",
            "scope_sheet_index": 0,
            "book_scope": False,
        },
        {
            "name": "two",  # VBA/GS send: "Sheet2!two"
            "sheet_index": 2,
            "address": "B3",
            "book_scope": False,
            "scope_sheet_name": "Sheet2",
            "scope_sheet_index": 1,
        },
        {
            "name": "two",
            "sheet_index": 1,
            "address": "A1:A2",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
    ],
    "sheets": [
        {
            "name": "Sheet 1",
            "visibility": "Visible",
            "print_area": "$A$1:$B$3",
            "values": [
                ["a", "b", "c", ""],
                [1.1, 2.2, 3.3, "2021-01-01T00:00:00.000Z"],
                [4.4, 5.5, 6.6, ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["Column1", "Column2", "", ""],
                [1.1, 2.2, "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                [1.1, 2.2, 3.3, ""],
                [4.4, 5.5, 6.6, ""],
                ["Total", "", 9.9, ""],
            ],
            # Office.js Range.address is relative ("A1"), while xlwings'
            # Range.address is absolute ("$A$1"). The engine normalizes them
            # when matching notes to cells.
            "notes": [{"address": "A1"}],
            "charts": [
                {
                    "name": "mychart1",
                    "chart_type": "Line",
                    "left": 10,
                    "top": 20,
                    "width": 300,
                    "height": 200,
                }
            ],
            "shapes": [
                {
                    "name": "myshape1",
                    "type": "GeometricShape",
                    "left": 11,
                    "top": 22,
                    "width": 33,
                    "height": 44,
                },
                {
                    "name": "mypic1",
                    "type": "Image",
                    "left": 50,
                    "top": 60,
                    "width": 20,
                    "height": 10,
                },
            ],
            "pictures": [
                {
                    "name": "mypic1",
                    "height": 10,
                    "width": 20,
                    "left": 50,
                    "top": 60,
                    "lock_aspect_ratio": True,
                },
                {
                    "name": "mypic2",
                    "height": 30,
                    "width": 40,
                    "left": 70,
                    "top": 80,
                    "lock_aspect_ratio": False,
                },
            ],
            "tables": [
                {
                    "name": "Table1",
                    "range_address": "A10:B11",
                    "header_row_range_address": "A10:B10",
                    "data_body_range_address": "A11:B11",
                    "total_row_range_address": None,
                    "show_headers": True,
                    "show_totals": False,
                    "table_style": "TableStyleMedium2",
                    "show_autofilter": True,
                    "show_table_style_first_column": True,
                    "show_table_style_last_column": False,
                    "show_table_style_row_stripes": True,
                    "show_table_style_column_stripes": False,
                },
                {
                    "name": "Table2",
                    "range_address": "A15:C17",
                    "header_row_range_address": None,
                    "data_body_range_address": "A15:C16",
                    "total_row_range_address": "A17:C17",
                    "show_headers": False,
                    "show_totals": True,
                    "table_style": "TableStyleLight1",
                    "show_autofilter": False,
                    "show_table_style_first_column": False,
                    "show_table_style_last_column": True,
                    "show_table_style_row_stripes": False,
                    "show_table_style_column_stripes": True,
                },
            ],
        },
        {
            "name": "Sheet2",
            "visibility": "Hidden",
            "values": [["aa", "bb"], [11.1, 22.2]],
            "pictures": [],
            "tables": [],
        },
        {
            "name": "Sheet3",
            "visibility": "VeryHidden",
            "values": [
                ["", "string"],
                [-1.1, 1.1],
                [True, False],
                ["2021-10-01T00:00:00.000Z", "2021-12-31T23:35:00.000Z"],
            ],
            "pictures": [],
            "tables": [],
        },
    ],
}


@pytest.fixture(scope="module")
def book():
    if engine == "remote":
        book = xw.Book(json=data)
    elif engine == "calamine":
        book = xw.Book(this_dir / f"engines.{file_extension}", mode="r")
    else:
        book = xw.Book(this_dir / f"engines.{file_extension}")
    yield book
    book.close()


@pytest.fixture(autouse=True)
def clear_json(book):
    book.impl._json = {"actions": []}


# range.value
def test_range_index(book):
    sheet = book.sheets[0]
    assert sheet.range((1, 1)).value == "a"
    assert sheet.range((1, 1), (3, 1)).value == ["a", 1.1, 4.4]
    assert sheet.range((1, 3), (3, 3)).value == ["c", 3.3, 6.6]
    assert sheet.range((1, 1), (3, 3)).value == [
        ["a", "b", "c"],
        [1.1, 2.2, 3.3],
        [4.4, 5.5, 6.6],
    ]
    assert sheet.range((2, 2), (3, 3)).value == [[2.2, 3.3], [5.5, 6.6]]


def test_range_a1(book):
    sheet = book.sheets[0]
    assert sheet.range("A1").value == "a"
    assert sheet.range("A1:A3").value == ["a", 1.1, 4.4]
    assert sheet.range("C1:C3").value == ["c", 3.3, 6.6]
    assert sheet.range("A1:C3").value == [
        ["a", "b", "c"],
        [1.1, 2.2, 3.3],
        [4.4, 5.5, 6.6],
    ]
    assert sheet.range("B2:C3").value == [[2.2, 3.3], [5.5, 6.6]]


def test_range_shortcut_address(book):
    sheet = book.sheets[0]
    assert sheet["A1"].value == "a"
    assert sheet["A1:A3"].value == ["a", 1.1, 4.4]
    assert sheet["C1:C3"].value == ["c", 3.3, 6.6]
    assert sheet["A1:C3"].value == [["a", "b", "c"], [1.1, 2.2, 3.3], [4.4, 5.5, 6.6]]
    assert sheet["B2:C3"].value == [[2.2, 3.3], [5.5, 6.6]]


def test_range_shortcut_index(book):
    sheet = book.sheets[0]
    assert sheet[0, 0].value == "a"
    assert sheet[0:3, 0].value == ["a", 1.1, 4.4]
    assert sheet[0:3, 2].value == ["c", 3.3, 6.6]
    assert sheet[0:3, 0:3].value == [["a", "b", "c"], [1.1, 2.2, 3.3], [4.4, 5.5, 6.6]]
    assert sheet[1:3, 1:3].value == [[2.2, 3.3], [5.5, 6.6]]


def test_range_from_range(book):
    sheet = book.sheets[0]
    assert sheet.range(sheet.range((1, 1)), sheet.range((3, 1))).value == [
        "a",
        1.1,
        4.4,
    ]
    assert sheet.range(sheet.range("C1"), sheet.range("C3")).value == ["c", 3.3, 6.6]
    assert sheet.range(sheet.range("A1"), sheet.range("C3")).value == [
        ["a", "b", "c"],
        [1.1, 2.2, 3.3],
        [4.4, 5.5, 6.6],
    ]
    assert sheet.range(sheet.range("B2"), sheet.range("C3")).value == [
        [2.2, 3.3],
        [5.5, 6.6],
    ]


def test_range_round_indexing(book):
    sheet = book.sheets[0]
    assert sheet["B2:C3"](1, 1).value == 2.2
    assert sheet["B2:C3"](1, 1).address == "$B$2"
    assert sheet["B2:C3"](2, 1).value == 5.5
    assert sheet["B2:C3"](2, 1).address == "$B$3"


def test_range_square_indexing_2d(book):
    sheet = book.sheets[0]
    assert sheet["B2:C3"][0, 0].value == 2.2
    assert sheet["B2:C3"][0, 0].address == "$B$2"
    assert sheet["B2:C3"][1, 0].value == 5.5
    assert sheet["B2:C3"][1, 0].address == "$B$3"


def test_range_square_indexing_1d(book):
    sheet1 = book.sheets[0]
    r = sheet1.range("A1:B2")
    assert r[0].address, "$A$1"
    assert r(1).address, "$A$1"


def test_range_slice1(book):
    r = book.sheets[0].range("B2:D4")
    assert r[0:, 1:].address == "$C$2:$D$4"


def test_range_resize(book):
    sheet1 = book.sheets[0]
    assert sheet1["A1"].resize(row_size=2, column_size=3).address == "$A$1:$C$2"
    assert (
        sheet1["A1"].resize(row_size=4, column_size=5).address == "$A$1:$E$4"
    )  # outside of used range


def test_range_offset(book):
    sheet1 = book.sheets[0]
    assert sheet1["A1"].offset(row_offset=2, column_offset=3).address == "$D$3"
    assert sheet1["A1"].offset(row_offset=10, column_offset=10).address == "$K$11"


def test_last_cell(book):
    sheet1 = book.sheets[0]
    assert sheet1["B3:F5"].last_cell.row == 5
    assert sheet1["B3:F5"].last_cell.column == 6


def test_expand(book):
    sheet1 = book.sheets[0]
    assert sheet1["A1"].expand().address == "$A$1:$C$3"
    assert sheet1["A1"].expand().value == [
        ["a", "b", "c"],
        [1.1, 2.2, 3.3],
        [4.4, 5.5, 6.6],
    ]
    assert sheet1["B1"].expand().address == "$B$1:$C$3"
    assert sheet1["B1"].expand().value == [["b", "c"], [2.2, 3.3], [5.5, 6.6]]
    assert sheet1["C3"].expand().address == "$C$3"
    assert sheet1["C3"].expand().value == 6.6

    # Edge case (no more rows/cols after expanded range
    sheet2 = book.sheets[1]
    assert sheet2["A1"].expand().value == [["aa", "bb"], [11.1, 22.2]]
    assert sheet2["A1"].expand().address == "$A$1:$B$2"


def test_completely_outside_usedrange(book):
    sheet = book.sheets[0]
    assert sheet["D5"].value is None
    assert sheet["D5:D6"].value == [None, None]
    assert sheet["D5:E7"].value == [[None, None], [None, None], [None, None]]


def test_partly_outside_usedrange(book):
    sheet = book.sheets[0]
    assert sheet["A4:A5"].value == [None, None]
    assert sheet["A3:A5"].value == [4.4, None, None]
    assert sheet["A4:B6"].value == [[None, None], [None, None], [None, None]]
    assert sheet["D4:F4"].value == [None, None, None]
    assert sheet["D4:F5"].value == [[None, None, None], [None, None, None]]


def test_len(book):
    assert len(book.sheets[0]["A1:C4"]) == 12


def test_count(book):
    assert len(book.sheets[0]["A1:C4"]) == book.sheets[0]["A1:C4"].count


# Conversion
@pytest.mark.skipif(not np, reason="requires NumPy")
def test_numpy_array(book):
    sheet = book.sheets[0]
    np.testing.assert_array_equal(
        sheet["B2:C3"].options(np.array).value, np.array([[2.2, 3.3], [5.5, 6.6]])
    )


@pytest.mark.skipif(not pd, reason="requires pandas")
def test_pandas_df(book):
    sheet = book.sheets[0]
    pd.testing.assert_frame_equal(
        sheet["A1:C3"].options(pd.DataFrame, index=False).value,
        pd.DataFrame(data=[[1.1, 2.2, 3.3], [4.4, 5.5, 6.6]], columns=["a", "b", "c"]),
    )


def test_read_basic_types(book):
    sheet = book.sheets[2]
    assert sheet["A1:B4"].value == [
        [None, "string"],
        [-1.1, 1.1],
        [True, False],
        [dt.datetime(2021, 10, 1, 0, 0), dt.datetime(2021, 12, 31, 23, 35)],
    ]


def test_read_basic_types_no_datetime(book):
    sheet = book.sheets[2]
    assert sheet["A1:B3"].value == [
        [None, "string"],
        [-1.1, 1.1],
        [True, False],
    ]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.skipif(not tz, reason="requires dateutil")
def test_write_basic_types(book):
    sheet = book.sheets[0]
    sheet["Z10"].value = [
        [None, "string"],
        [-1.1, 1.1],
        [True, False],
        [
            dt.date(2021, 10, 1),
            dt.datetime(2021, 12, 31, 23, 35, tzinfo=tz.gettz("Europe/Paris")),
        ],
    ]
    assert (
        json.dumps(book.json()["actions"][0]["values"])
        == '[["", "string"], [-1.1, 1.1], [true, false], '
        '["2021-10-01", "2021-12-31 23:35:00"]]'
    )


# sheets
def test_sheet_access(book):
    assert book.sheets[0] == book.sheets["Sheet 1"]
    assert book.sheets[1] == book.sheets["Sheet2"]
    assert book.sheets[0].name == "Sheet 1"
    assert book.sheets[1].name == "Sheet2"


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_sheet_active(book):
    assert book.sheets.active == book.sheets[0]


def test_sheets_iteration(book):
    for ix, sheet in enumerate(book.sheets):
        assert sheet.name == "Sheet 1" if ix == 0 else "Sheet2"


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_sheet_page_setup_print_area(book):
    assert book.sheets[0].page_setup.print_area == "$A$1:$B$3"
    # a sheet without one reports None, not an empty string
    assert book.sheets[1].page_setup.print_area is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_page_setup_print_area_setter():
    book = xw.Book(json=json.loads(json.dumps(data)))
    page_setup = book.sheets[0].page_setup
    page_setup.print_area = "$C$1:$D$5"
    action = book.json()["actions"][-1]
    assert action["func"] == "setPrintArea"
    assert action["args"] == ["$C$1:$D$5"]
    assert action["sheet_position"] == 0
    # written through, so a read-after-write in the same script is correct
    assert page_setup.print_area == "$C$1:$D$5"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_page_setup_print_area_clear():
    # The public API documents `print_area = None` as the way to clear it.
    book = xw.Book(json=json.loads(json.dumps(data)))
    page_setup = book.sheets[0].page_setup
    page_setup.print_area = None
    assert book.json()["actions"][-1]["args"] == [None]
    assert page_setup.print_area is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_page_setup_print_area_missing_from_payload():
    # Clients that predate the payload field report no print area rather than
    # raising, which is what None means anyway.
    payload = json.loads(json.dumps(data))
    del payload["sheets"][0]["print_area"]
    book = xw.Book(json=payload)
    assert book.sheets[0].page_setup.print_area is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_visible(book):
    # Office.js' "VeryHidden" maps to False like "Hidden" does, as xlwings'
    # public API is a bool.
    assert book.sheets[0].visible is True
    assert book.sheets[1].visible is False
    assert book.sheets[2].visible is False


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_visible_set_hidden(book):
    book.sheets[0].visible = False
    assert book.json()["actions"][0]["func"] == "setSheetVisibility"
    assert book.json()["actions"][0]["args"] == ["Hidden"]
    assert book.json()["actions"][0]["sheet_position"] == 0
    # written through, so a read-after-write in the same script is correct
    assert book.sheets[0].visible is False
    book.sheets[0].visible = True


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_visible_set_visible(book):
    book.sheets[1].visible = True
    assert book.json()["actions"][0]["func"] == "setSheetVisibility"
    assert book.json()["actions"][0]["args"] == ["Visible"]
    assert book.json()["actions"][0]["sheet_position"] == 1
    assert book.sheets[1].visible is True
    book.sheets[1].visible = False


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_delete_keeps_active_index_valid():
    # Deleting a sheet must leave active_sheet_index pointing at an existing
    # sheet, otherwise Sheets.active raises IndexError.
    def fresh():
        return xw.Book(json=json.loads(json.dumps(data)))

    # delete a sheet after the active one: index unchanged
    book = fresh()
    book.sheets[0].activate()
    book.sheets[2].delete()
    assert book.sheets.active.name == "Sheet 1"

    # delete a sheet before the active one: index shifts down with it
    book = fresh()
    book.sheets[2].activate()
    book.sheets[0].delete()
    assert book.sheets.active.name == "Sheet3"

    # delete the active sheet itself: the next one takes its place
    book = fresh()
    book.sheets[1].activate()
    book.sheets[1].delete()
    assert book.sheets.active.name == "Sheet3"

    # delete the active sheet when it's the last one: falls back to the
    # new last sheet rather than running off the end
    book = fresh()
    book.sheets[2].activate()
    book.sheets[2].delete()
    assert book.sheets.active.name == "Sheet2"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_delete_after_add_keeps_active_index_valid():
    # Sheets.add() makes the new sheet active; deleting it must restore a
    # valid index (the original IndexError repro).
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets.add(name="freshsheet")
    sheet.delete()
    assert book.sheets.active.name in ("Sheet 1", "Sheet2", "Sheet3")
    assert len(book.sheets) == 3


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_visible_added_sheet():
    # Sheets.add() seeds "visibility", so the getter works before the next
    # round-trip rather than raising KeyError. Uses its own book: adding a
    # sheet mutates the module-scoped fixture's active_sheet_index.
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets.add(name="freshsheet")
    assert sheet.visible is True
    sheet.visible = False
    assert book.json()["actions"][-1]["func"] == "setSheetVisibility"
    assert book.json()["actions"][-1]["args"] == ["Hidden"]
    assert sheet.visible is False


# book name
def test_book(book):
    assert book.name == f"engines.{file_extension}"


@pytest.mark.skipif(engine in ["calamine", "excel"], reason="calamine engine")
def test_book_selection(book):
    assert book.selection.address == "$B$3:$B$4"


# pictures
@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_len(book):
    assert len(book.sheets[0].pictures) == 2


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_name(book):
    assert book.sheets[0].pictures[0].name == "mypic1"
    assert book.sheets[0].pictures[1].name == "mypic2"
    assert book.sheets[0].pictures(1).name == "mypic1"
    assert book.sheets[0].pictures(2).name == "mypic2"


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_width(book):
    assert book.sheets[0].pictures[0].width == 20
    assert book.sheets[0].pictures[1].width == 40


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_height(book):
    assert book.sheets[0].pictures[0].height == 10
    assert book.sheets[0].pictures[1].height == 30


# Collections that create an object locally, so that the public API can read it
# back before the next round-trip. Each entry is the collection's attribute on
# the sheet, plus a callable that adds one. Their seeded keys have to keep
# matching the payload the client sends, which the fixture mirrors -- three
# separate getters have shipped KeyError bugs from that drifting apart, and a
# fourth (print_area) was latent.
#
# A shared seeding helper wouldn't prevent this: the dicts are genuinely
# different shapes and their values are per-class (Excel's defaults, the
# caller's arguments, or None). The failure mode is drift between two lists,
# which is what these tests are for.
_SEEDING_COLLECTIONS = [
    (
        "pictures",
        lambda sheet: sheet.pictures.add(this_dir.parent / "sample_picture.png"),
    ),
    ("tables", lambda sheet: sheet.tables.add(source=sheet["A1:B2"])),
]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize("attribute,add", _SEEDING_COLLECTIONS)
def test_add_seeds_every_payload_key(attribute, add):
    """A locally created object must carry every key the client sends.

    Otherwise its getters raise KeyError until the next round-trip refreshes
    the payload -- the bug that hit Sheets.add(), Tables.add() and
    Pictures.add() in turn.
    """
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    from_payload = set(sheet.impl.api[attribute][0])
    add(sheet)
    seeded = set(sheet.impl.api[attribute][-1])
    missing = from_payload - seeded
    assert not missing, (
        f"{attribute}.add() doesn't seed {sorted(missing)}, so those getters "
        f"raise KeyError on a newly added object"
    )


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheets_add_seeds_every_payload_key():
    # Same check for sheets, whose payload entry is keyed differently (its
    # values are nested collections rather than scalars).
    book = xw.Book(json=json.loads(json.dumps(data)))
    from_payload = set(book.sheets[0].impl.api)
    sheet = book.sheets.add(name="seedcheck")
    missing = from_payload - set(sheet.impl.api)
    # used_range_address is client-only metadata, not something add() can know
    missing -= {"used_range_address"}
    assert not missing, (
        f"Sheets.add() doesn't seed {sorted(missing)}, so those getters raise "
        f"KeyError on a newly added sheet"
    )


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_copy_seeds_every_payload_key():
    # Sheet.copy() builds its entry from the source sheet's, so it inherits
    # the keys -- but pin it, since it's a separate code path.
    book = xw.Book(json=json.loads(json.dumps(data)))
    from_payload = set(book.sheets[0].impl.api)
    copied = book.sheets[0].copy()
    assert from_payload - set(copied.impl.api) == set()


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_copy_owns_its_nested_metadata():
    book = xw.Book(json=json.loads(json.dumps(data)))
    source = book.sheets[0]
    copied = source.copy()

    for key in ["values", "pictures", "shapes", "charts", "notes", "tables"]:
        assert copied.impl.api[key] is not source.impl.api[key]

    copied.charts[0].name = "copied-chart"
    assert source.charts[0].name == "mychart1"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_charts_collection(book):
    charts = book.sheets[0].charts
    assert len(charts) == 1
    assert charts[0].name == "mychart1"
    assert "mychart1" in charts


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_properties(book):
    chart = book.sheets[0].charts[0]
    # the payload carries Office.js' enum value; the getter maps it back
    assert chart.chart_type == "line"
    assert chart.left == 10
    assert chart.top == 20
    assert chart.width == 300
    assert chart.height == 200
    assert chart.parent.name == book.sheets[0].name


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "attribute,func,value,expected",
    [
        ("name", "setChartName", "renamed", "renamed"),
        ("chart_type", "setChartType", "pie", "Pie"),
        ("left", "setChartPosition", 111, 111),
        ("top", "setChartPosition", 222, 222),
        ("width", "setChartPosition", 333, 333),
        ("height", "setChartPosition", 444, 444),
    ],
)
def test_chart_setters(attribute, func, value, expected):
    book = xw.Book(json=json.loads(json.dumps(data)))
    chart = book.sheets[0].charts[0]
    setattr(chart, attribute, value)
    action = book.json()["actions"][-1]
    assert action["func"] == func
    assert expected in action["args"]
    # written through, so a read-after-write in the same script is correct
    assert getattr(chart, attribute) == value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_type_invalid(book):
    with pytest.raises(ValueError, match="Invalid chart type"):
        book.sheets[0].charts[0].chart_type = "nonsense"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_charts_add_defers_until_source_data():
    # Office.js' charts.add() needs a type and source data together, which the
    # documented xlwings flow doesn't have at add() time -- so the chart is
    # created on the first set_source_data() instead.
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    before = len(sheet.charts)
    chart = sheet.charts.add()
    assert book.json()["actions"] == []
    assert len(sheet.charts) == before

    chart.set_source_data(sheet["A1:B2"])
    assert len(sheet.charts) == before + 1
    action = book.json()["actions"][-1]
    assert action["func"] == "addChart"
    assert action["args"][1] == "ColumnClustered"
    assert action["args"][3] == "$A$1:$B$2"

    # ...and the documented follow-up calls work on the created chart
    chart.chart_type = "line"
    assert book.json()["actions"][-1]["func"] == "setChartType"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_charts_add_assigns_unique_default_names():
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    first = sheet.charts.add()
    second = sheet.charts.add()

    first.set_source_data(sheet["A1:B2"])
    second.set_source_data(sheet["A1:B2"])

    add_actions = [
        action for action in book.json()["actions"] if action["func"] == "addChart"
    ]
    names = [action["args"][0] for action in add_actions]
    assert len(names) == len(set(names)) == 2


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_set_source_data_on_existing_chart():
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    sheet.charts[0].set_source_data(sheet["A1:B2"])
    action = book.json()["actions"][-1]
    assert action["func"] == "setChartSourceData"
    assert action["args"] == [0, sheet.name, "$A$1:$B$2"]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_delete():
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    sheet.charts[0].delete()
    assert len(sheet.charts) == 0
    assert book.json()["actions"][-1]["func"] == "deleteChart"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_delete_pending_chart():
    # A chart that never got source data was never created in Excel, so
    # deleting it queues nothing.
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0].charts.add().delete()
    assert book.json()["actions"] == []


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_to_png(book):
    # Queues an action that writes the file, like Range.to_png(); it lands
    # when the script returns or on the next await book.flush().
    book.sheets[0].charts[0].to_png("out.png")
    action = book.json()["actions"][-1]
    assert action["func"] == "chartToPng"
    assert action["args"] == [0, "out.png"]
    assert action["sheet_position"] == 0


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_to_png_defaults_the_path(book):
    # The public method resolves the default before reaching the impl.
    chart = book.sheets[0].charts[0]
    chart.to_png()
    assert book.json()["actions"][-1]["args"][1].endswith(f"{chart.name}.png")


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_to_png_requires_source_data_for_pending_chart(book):
    chart = book.sheets[0].charts.add()

    with pytest.raises(xw.XlwingsError, match="set_source_data"):
        chart.to_png("out.png")

    assert book.json()["actions"] == []


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_chart_to_pdf_and_get_png_not_supported(book):
    chart = book.sheets[0].charts[0]
    with pytest.raises(NotImplementedError, match="has no PDF export"):
        chart.to_pdf("out.pdf")
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(chart.get_png())


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_range_note(book):
    # The payload says which cells have a note, not what they say, so this is
    # answerable synchronously without carrying every note's text.
    assert book.sheets[0]["A1"].note is not None
    # a cell without one reports None, as on the other engines
    assert book.sheets[0]["B2"].note is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_note_text_sync_points_at_async(book):
    with pytest.raises(NotImplementedError, match=r"get_text\(\)"):
        book.sheets[0]["A1"].note.text


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_note_get_text_not_supported_off_lite(book):
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(book.sheets[0]["A1"].note.get_text())


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_note_text_setter():
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0]["A1"].note.text = "updated"
    action = book.json()["actions"][-1]
    assert action["func"] == "setNoteText"
    assert action["args"] == ["$A$1", "updated"]
    assert action["sheet_position"] == 0


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_note_delete():
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0]["A1"].note.delete()
    action = book.json()["actions"][-1]
    assert action["func"] == "deleteNote"
    assert action["args"] == ["$A$1"]
    # gone locally too, so Range.note reports None right away
    assert book.sheets[0]["A1"].note is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_note_missing_from_payload():
    # Clients that predate the payload field report no note rather than
    # raising, which is what None means anyway.
    payload = json.loads(json.dumps(data))
    del payload["sheets"][0]["notes"]
    book = xw.Book(json=payload)
    assert book.sheets[0]["A1"].note is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shapes_collection(book):
    shapes = book.sheets[0].shapes
    assert len(shapes) == 2
    assert [shape.name for shape in shapes] == ["myshape1", "mypic1"]
    # index is impl-only on Shape, as it is on Table and Picture
    assert shapes["myshape1"].impl.index == 1
    assert "myshape1" in shapes
    assert "nope" not in shapes
    # a picture is a shape too, so it appears in both collections
    assert shapes["mypic1"].type == "picture"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_properties(book):
    shape = book.sheets[0].shapes[0]
    assert shape.name == "myshape1"
    # mapped to xlwings' vocabulary, as on the desktop engines
    assert shape.type == "auto_shape"
    assert shape.left == 11
    assert shape.top == 22
    assert shape.width == 33
    assert shape.height == 44
    assert shape.parent.name == book.sheets[0].name


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "attribute,func,value",
    [
        ("name", "setShapeName", "renamed"),
        ("left", "setShapeLeft", 123),
        ("top", "setShapeTop", 234),
        ("width", "setShapeWidth", 345),
        ("height", "setShapeHeight", 456),
    ],
)
def test_shape_setters(attribute, func, value):
    book = xw.Book(json=json.loads(json.dumps(data)))
    shape = book.sheets[0].shapes[0]
    setattr(shape, attribute, value)
    action = book.json()["actions"][-1]
    assert action["func"] == func
    assert action["args"] == [0, value]
    assert action["sheet_position"] == 0
    # written through, so a read-after-write in the same script is correct
    assert getattr(shape, attribute) == value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_text_setter_and_getter(book):
    shape = book.sheets[0].shapes[0]
    shape.text = "hello"
    action = book.json()["actions"][-1]
    assert action["func"] == "setShapeText"
    assert action["args"] == [0, "hello"]
    # shape text is unbounded, so it's fetched on demand rather than sent
    # with every request; the sync property points at the async method
    with pytest.raises(NotImplementedError, match=r"get_text\(\)"):
        shape.text


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize("reported", ["hello world", None])
def test_shape_get_text(book, reported):
    # None is what a shape without text reports, as on the desktop engines.
    shape = book.sheets[0].shapes[0]

    async def fake(self, key):
        assert key == "text"
        return reported

    with mock.patch.object(xw.pro._xlremote.Shape, "_get_shape_data", fake):
        assert asyncio.run(shape.get_text()) == reported


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_get_text_not_supported_off_lite(book):
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(book.sheets[0].shapes[0].get_text())


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "method,relative,scale,expected",
    [
        ("scale_height", False, "scale_from_top_left", "CurrentSize"),
        ("scale_width", True, "scale_from_middle", "OriginalSize"),
    ],
)
def test_shape_scale(method, relative, scale, expected):
    book = xw.Book(json=json.loads(json.dumps(data)))
    shape = book.sheets[0].shapes[0]
    getattr(shape, method)(1.5, relative, scale)
    action = book.json()["actions"][-1]
    assert action["func"] == "scaleShape"
    assert action["args"][1] == 1.5
    assert action["args"][2] == expected
    assert action["args"][4] == method.split("_")[1]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_scale_updates_the_local_geometry():
    # Scaling has to be visible to a read-after-write in the same script, the
    # way the width/height setters are.
    book = xw.Book(json=json.loads(json.dumps(data)))
    shape = book.sheets[0].shapes[0]
    width, height = shape.width, shape.height

    shape.scale_width(1.5, False, "scale_from_top_left")
    shape.scale_height(2.0, False, "scale_from_top_left")
    assert shape.width == width * 1.5
    assert shape.height == height * 2.0


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_scale_from_original_size_leaves_geometry_alone():
    # The original size isn't something this engine knows, so that case waits
    # for the next round-trip rather than guessing.
    book = xw.Book(json=json.loads(json.dumps(data)))
    shape = book.sheets[0].shapes[0]
    width = shape.width
    shape.scale_width(1.5, True, "scale_from_top_left")
    assert shape.width == width


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_scale_invalid_anchor(book):
    with pytest.raises(ValueError, match="Invalid scale"):
        book.sheets[0].shapes[0].scale_height(1.5, False, "nonsense")


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_delete():
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    sheet.shapes[0].delete()
    assert len(sheet.shapes) == 1
    action = book.json()["actions"][-1]
    assert action["func"] == "deleteShape"
    assert action["args"] == [0]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "reported,expected",
    [
        ("Image", "picture"),
        ("GeometricShape", "auto_shape"),
        ("Group", "group"),
        ("Line", "line"),
        # no xlwings equivalent, so it passes through
        ("Unsupported", "Unsupported"),
    ],
)
def test_shape_type_uses_xlwings_names(reported, expected):
    # Office.js' ShapeType is coarser than the desktop engines', but each of
    # its members maps onto an xlwings name, so the property is portable.
    payload = json.loads(json.dumps(data))
    payload["sheets"][0]["shapes"][0]["type"] = reported
    book = xw.Book(json=payload)
    assert book.sheets[0].shapes[0].type == expected


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_unsupported_members(book):
    with pytest.raises(NotImplementedError, match="no way to activate"):
        book.sheets[0].shapes[0].activate()


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "item,start,length",
    [
        (slice(3, 7), 3, 4),
        (slice(None, 5), 0, 5),
        (slice(2, None), 2, None),
        (2, 2, 1),
    ],
)
def test_shape_characters_slicing(item, start, length):
    # The slice is carried as start/length, which is what getSubstring takes.
    book = xw.Book(json=json.loads(json.dumps(data)))
    characters = book.sheets[0].shapes[0].characters[item]
    characters.font.bold = True
    action = book.json()["actions"][-1]
    assert action["func"] == "setShapeFontProperty"
    assert action["args"] == [0, start, length, "bold", True]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_font_setter_uses_shape_action():
    # A shape's font can't use the range font action, which addresses cells.
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0].shapes[0].font.italic = True
    action = book.json()["actions"][-1]
    assert action["func"] == "setShapeFontProperty"
    # no slice: the whole text range
    assert action["args"] == [0, None, None, "italic", True]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_font_and_characters_getters(book):
    shape = book.sheets[0].shapes[0]

    async def fake(self, key, start=None, length=None):
        return {
            "characters_text": "slice",
            "font": {
                "bold": True,
                "italic": False,
                "size": 12.0,
                "color": "#00ff00",
                "name": "Arial",
            },
        }[key]

    with mock.patch.object(xw.pro._xlremote.Shape, "_get_shape_data", fake):
        assert asyncio.run(shape.font.get_bold()) is True
        assert asyncio.run(shape.characters[1:4].get_text()) == "slice"
        assert asyncio.run(shape.characters[1:4].font.get_color()) == (0, 255, 0)


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_font_getter_on_shape_without_text(book):
    # getShapeData reports null for a shape with no text; the attributes come
    # back as None rather than raising.
    shape = book.sheets[0].shapes[0]

    async def fake(self, key, start=None, length=None):
        return None

    with mock.patch.object(xw.pro._xlremote.Shape, "_get_shape_data", fake):
        assert asyncio.run(shape.font.get_bold()) is None
        assert asyncio.run(shape.font.get_color()) is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_shape_characters_sync_text_points_at_async(book):
    with pytest.raises(NotImplementedError, match=r"get_text\(\)"):
        book.sheets[0].shapes[0].characters.text


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_column_width_passes_points_through(book):
    # Points on this engine, which is what the host measures in -- no
    # conversion, so a value written is the value read back.
    book.sheets[0]["A1"].column_width = 110.5
    action = book.json()["actions"][-1]
    assert action["func"] == "setColumnWidth"
    assert action["args"] == [110.5]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize("value", [-1, True, "wide", None])
def test_range_column_width_rejects_invalid(book, value):
    with pytest.raises(ValueError, match="column_width"):
        book.sheets[0]["A1"].column_width = value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_column_width_allows_zero_and_large_values(book):
    # 0 hides the column, and points aren't capped at the 255 that bounded the
    # old character unit.
    for value in [0, 400]:
        book.sheets[0]["A1"].column_width = value
        assert book.json()["actions"][-1]["args"] == [value]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_to_png(book):
    # Already worked via the rangeToPng action, unlike the other exports.
    book.sheets[0]["A1"].to_png("out.png")
    action = book.json()["actions"][-1]
    assert action["func"] == "rangeToPng"
    assert action["args"] == ["out.png"]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "call,match",
    [
        (lambda rng: rng.copy_picture(), "no clipboard access"),
        (lambda rng: rng.paste(), "no clipboard access"),
        (lambda rng: rng.to_pdf(), "no PDF export"),
    ],
)
def test_range_unsupported_exports(book, call, match):
    with pytest.raises(NotImplementedError, match=match):
        call(book.sheets[0]["A1"])


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_characters_not_supported(book):
    # Office.js has no character-range object for cells; only shapes have one.
    with pytest.raises(NotImplementedError, match="address a range of characters"):
        book.sheets[0]["A1"].characters


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_left_top_lock_aspect_ratio(book):
    assert book.sheets[0].pictures[0].left == 50
    assert book.sheets[0].pictures[0].top == 60
    assert book.sheets[0].pictures[0].lock_aspect_ratio is True
    # the second picture is the inverse, so a getter reading the wrong
    # entry can't pass by accident
    assert book.sheets[0].pictures[1].left == 70
    assert book.sheets[0].pictures[1].top == 80
    assert book.sheets[0].pictures[1].lock_aspect_ratio is False


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "attribute,func,value",
    [
        ("left", "setPictureLeft", 123),
        ("top", "setPictureTop", 456),
        ("lock_aspect_ratio", "setPictureLockAspectRatio", False),
        ("width", "setPictureWidth", 321),
        ("height", "setPictureHeight", 654),
    ],
)
def test_pictures_geometry_setters(attribute, func, value):
    book = xw.Book(json=json.loads(json.dumps(data)))
    picture = book.sheets[0].pictures[0]
    setattr(picture, attribute, value)
    action = book.json()["actions"][-1]
    assert action["func"] == func
    assert action["args"] == [0, value]
    assert action["sheet_position"] == 0
    # written through, so a read-after-write in the same script is correct
    assert getattr(picture, attribute) == value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_pictures_add_seeds_geometry():
    # Pictures.add() seeds the new picture's api dict, so the getters work
    # before the next round-trip rather than raising KeyError.
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    picture = sheet.pictures.add(
        this_dir.parent / "sample_picture.png", name="new", left=5, top=15
    )
    assert picture.left == 5
    assert picture.top == 15
    assert picture.lock_aspect_ratio is None


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_add_and_delete(book):
    sheet = book.sheets[0]
    sheet.pictures.add(this_dir.parent / "sample_picture.png", name="new")
    assert len(sheet.pictures) == 3
    assert sheet.pictures[2].name == "new"
    # assert sheet.pictures[2].impl.index == 3  # TODO: implement
    sheet.pictures["new"].delete()
    assert len(sheet.pictures) == 2


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_iter(book):
    sheet = book.sheets[0]
    pic_names = []
    for pic in sheet.pictures:
        pic_names.append(pic.name)
    assert pic_names == ["mypic1", "mypic2"]


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_contains(book):
    sheet = book.sheets[0]
    assert "mypic1" in sheet.pictures
    assert 1 in sheet.pictures
    assert "mypic2" in sheet.pictures
    assert 2 in sheet.pictures
    assert "no" not in sheet.pictures
    assert 3 not in sheet.pictures


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_empty_pictures(book):
    assert not book.sheets[1].pictures


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_picture_exists(book):
    with pytest.raises(xw.ShapeAlreadyExists):
        book.sheets[0].pictures.add(
            this_dir.parent / "sample_picture.png", name="mypic1"
        )


# Named Ranges
def test_named_range_book_scope(book):
    sheet1 = book.sheets[0]
    assert sheet1["one"].address == "$A$1"


def test_named_range_sheet_scope(book):
    sheet1 = book.sheets[0]
    assert sheet1["two"].address == "$C$7:$D$8"


@pytest.mark.skipif(engine == "excel", reason="unhandled engine error")
def test_named_range_missing(book):
    sheet1 = book.sheets[0]
    with pytest.raises(xw.NoSuchObjectError):
        sheet1["doesnt_exist"].value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_named_range_book_change_value(book):
    sheet1 = book.sheets[0]
    assert sheet1["one"].value == "a"
    sheet1["one"].value = 1000
    assert book.json()["actions"][0]["values"] == [[1000]]
    assert book.json()["actions"][0]["sheet_position"] == 0
    assert book.json()["actions"][0]["start_row"] == 0
    assert book.json()["actions"][0]["start_column"] == 0


# Names collection
def test_names_len(book):
    assert len(book.names) == 4


def test_names_index_vs_name(book):
    assert book.names[0].name == "one"
    assert book.names["one"].name == "one"


@pytest.mark.skipif(engine == "calamine", reason="doesn't support local scope yet")
def test_name_local_scope1(book):
    assert book.names[1].name == "'Sheet 1'!two"
    assert book.names[2].name == "Sheet2!two"


@pytest.mark.skipif(engine == "calamine", reason="doesn't support local scope yet")
def test_name_local_scope2(book):
    assert book.sheets["Sheet 1"].names[0].name == "'Sheet 1'!two"
    assert book.sheets["Sheet2"].names[0].name == "Sheet2!two"


def test_name_refers_to(book):
    assert book.names[0].refers_to == "='Sheet 1'!$A$1"


def test_name_refers_to_range(book):
    assert book.names[0].refers_to_range == book.sheets[0]["A1"]
    assert book.names[1].refers_to_range == book.sheets[0]["C7:D8"]
    assert book.names[3].refers_to_range == book.sheets[1]["A1:A2"]


def test_name_contains(book):
    assert "one" in book.names


def test_names_iter(book):
    for ix, name in enumerate(book.names):
        if ix == 0:
            assert name.refers_to_range == book.sheets[0]["A1"]
        elif ix == 1:
            assert name.refers_to_range == book.sheets[0]["C7:D8"]
        elif ix == 3:
            assert name.refers_to_range == book.sheets[1]["A1:A2"]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_range_get_name(book):
    assert book.sheets[0]["A1"].name.name == "one"
    assert book.sheets[0]["C7:D8"].name.name == "'Sheet 1'!two"
    assert book.sheets[1]["A1:A2"].name.name == "two"
    assert book.sheets[0]["X1"].name is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_set_name(book):
    book.sheets[0]["A1:C3"].name = "mytestrange"
    assert json.dumps(book.json()["actions"][0]["func"]) == '"setRangeName"'
    assert json.dumps(book.json()["actions"][0]["args"][0]) == '"mytestrange"'


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_book_names_add(book):
    book.names.add("test1", "=Sheet1!$A$1:$B$3")
    assert book.json()["actions"][0]["func"] == "namesAdd"
    assert book.json()["actions"][0]["args"] == ["test1", "=Sheet1!$A$1:$B$3"]
    assert book.json()["actions"][0]["sheet_position"] is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_names_add(book):
    book.sheets[0].names.add("test1", "=Sheet1!$A$1:$B$3")
    assert book.json()["actions"][0]["func"] == "namesAdd"
    assert book.json()["actions"][0]["args"] == ["test1", "=Sheet1!$A$1:$B$3"]
    assert book.json()["actions"][0]["sheet_position"] == 0


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_names_add_seeds_scope_so_the_name_is_usable():
    # A name added mid-script has to be usable straight away: delete() and the
    # refers_to setter both read the scope keys, which add() didn't seed.
    book = xw.Book(json=json.loads(json.dumps(data)))

    book_scoped = book.names.add("BookScoped", f"={book.sheets[0].name}!$A$1")
    book_scoped.refers_to = f"={book.sheets[0].name}!$B$2"
    assert book.json()["actions"][-1]["args"][1:3] == [True, None]
    book_scoped.delete()
    assert book.json()["actions"][-1]["func"] == "nameDelete"

    sheet = book.sheets[0]
    sheet_scoped = sheet.names.add("SheetScoped", f"={sheet.name}!$C$3")
    sheet_scoped.refers_to = f"={sheet.name}!$D$4"
    # sheet-scoped names carry their scope index, which nameDelete needs too
    assert book.json()["actions"][-1]["args"][1:3] == [False, sheet.index - 1]
    sheet_scoped.delete()
    assert book.json()["actions"][-1]["func"] == "nameDelete"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_name_refers_to_setter():
    book = xw.Book(json=json.loads(json.dumps(data)))
    name = book.names[0]
    assert name.api["book_scope"] is True
    name.refers_to = "=Sheet2!$C$3"
    action = book.json()["actions"][-1]
    assert action["func"] == "setNameRefersTo"
    assert action["args"] == [name.api["name"], True, None, "=Sheet2!$C$3"]
    # refers_to is computed from sheet_index/address, so the setter updates
    # those -- check it round-trips through the getter and refers_to_range
    assert name.refers_to == "=Sheet2!$C$3"
    assert name.refers_to_range.sheet.name == "Sheet2"
    assert name.refers_to_range.address == "$C$3"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_name_refers_to_setter_sheet_scope():
    book = xw.Book(json=json.loads(json.dumps(data)))
    name = next(n for n in book.names if not n.api["book_scope"])
    name.refers_to = "=Sheet2!$D$4"
    action = book.json()["actions"][-1]
    assert action["func"] == "setNameRefersTo"
    # sheet-scoped names carry their scope index, as nameDelete does
    assert action["args"][1] is False
    assert action["args"][2] == name.api["scope_sheet_index"]
    assert action["args"][3] == "=Sheet2!$D$4"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_name_refers_to_setter_unknown_sheet(book):
    with pytest.raises(ValueError, match="doesn't exist"):
        book.names[0].refers_to = "=NoSuchSheet!$A$1"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_name_name_setter_not_supported(book):
    # Excel.NamedItem.name is read-only in Office.js.
    with pytest.raises(NotImplementedError, match="is read-only"):
        book.names[0].name = "newname"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_name_delete(book):
    book.names[0].delete()
    assert book.json()["actions"][0]["func"] == "nameDelete"
    assert book.json()["actions"][0]["args"] == [
        "one",
        "='Sheet 1'!$A$1",
        "one",
        0,
        True,
        None,
    ]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_delete(book):
    book.sheets[0]["A1:B2"].delete("up")
    assert book.json()["actions"][0]["func"] == "rangeDelete"
    assert book.json()["actions"][0]["args"] == ["up"]


# Tables
@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_names(book):
    assert book.sheets[0].tables[0].name == "Table1"
    assert book.sheets[0].tables[1].name == "Table2"


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_range(book):
    assert book.sheets[0].tables[0].range == book.sheets[0]["A10:B11"]
    assert book.sheets[0].tables[1].range == book.sheets[0]["A15:C17"]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_header_row_range(book):
    assert book.sheets[0].tables[0].header_row_range == book.sheets[0]["A10:B10"]
    assert book.sheets[0].tables[1].header_row_range is None


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_totals_row_range(book):
    assert book.sheets[0].tables[0].totals_row_range is None
    assert book.sheets[0].tables[1].totals_row_range == book.sheets[0]["A17:C17"]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_show_headers(book):
    assert book.sheets[0].tables[0].show_headers is True
    assert book.sheets[0].tables[1].show_headers is False


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_show_totals(book):
    assert book.sheets[0].tables[0].show_totals is False
    assert book.sheets[0].tables[1].show_totals is True


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_table_style(book):
    assert book.sheets[0].tables[0].table_style == "TableStyleMedium2"
    assert book.sheets[0].tables[1].table_style == "TableStyleLight1"


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_data_body_range(book):
    assert book.sheets[0].tables[0].data_body_range == book.sheets[0]["A11:B11"]
    assert book.sheets[0].tables[1].data_body_range == book.sheets[0]["A15:C16"]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_parent(book):
    assert book.sheets[0].tables[0].parent == book.sheets[0]
    assert book.sheets[0].tables[1].parent == book.sheets[0]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_show_autofilter(book):
    assert book.sheets[0].tables[0].show_autofilter is True
    assert book.sheets[0].tables[1].show_autofilter is False


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_get_values(book):
    assert book.sheets[0].tables[0].range.value == [["Column1", "Column2"], [1.1, 2.2]]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_range_with_table_name(book):
    assert book.sheets[0]["Table1"].value == [1.1, 2.2]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_tables_add(book):
    sheet1 = book.sheets[0]
    sheet1.tables.add(sheet1["A1:B2"], name="Table1")
    assert book.json() == {
        "actions": [
            {
                "func": "addTable",
                "args": ["$A$1:$B$2", True, "TableStyleMedium2", "Table1"],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }


@pytest.mark.skipif(not pd, reason="requires pandas")
@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_tables_update(book):
    sheet1 = book.sheets[0]
    sheet1.tables[0].update(pd.DataFrame({"A": [1, 2], "B": [3, 4]}))
    assert book.json() == {
        "actions": [
            {
                "func": "resizeTable",
                "args": [0, "$A$10:$C$12"],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
            {
                "func": "setValues",
                "args": [None],
                "values": [[" ", "A", "B"]],
                "sheet_position": 0,
                "start_row": 9,
                "start_column": 0,
                "row_count": 1,
                "column_count": 3,
            },
            {
                "func": "setValues",
                "args": [None],
                "values": [[0, 1, 3], [1, 2, 4]],
                "sheet_position": 0,
                "start_row": 10,
                "start_column": 0,
                "row_count": 2,
                "column_count": 3,
            },
        ]
    }


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_tables_resize(book):
    sheet1 = book.sheets[0]
    sheet1.tables[0].resize(sheet1["A10:C12"])
    assert book.json() == {
        "actions": [
            {
                "func": "resizeTable",
                "args": [0, "$A$10:$C$12"],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_set_name(book):
    sheet1 = book.sheets[0]
    sheet1.tables[0].name = "NewName"
    assert book.json() == {
        "actions": [
            {
                "func": "setTableName",
                "args": [0, "NewName"],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_set_show_autofilter(book):
    sheet1 = book.sheets[0]
    sheet1.tables[0].show_autofilter = False
    assert book.json() == {
        "actions": [
            {
                "func": "showAutofilterTable",
                "args": [0, False],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_set_show_headers(book):
    sheet1 = book.sheets[0]
    sheet1.tables[0].show_headers = False
    assert book.json() == {
        "actions": [
            {
                "func": "showHeadersTable",
                "args": [0, False],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_set_show_totals(book):
    sheet1 = book.sheets[0]
    sheet1.tables[0].show_totals = True
    assert book.json() == {
        "actions": [
            {
                "func": "showTotalsTable",
                "args": [0, True],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "attribute",
    ["cut_copy_mode", "enable_events", "interactive", "status_bar"],
)
def test_app_unsupported_read_write(book, attribute):
    # Excel.Application has no equivalent, so both accessors raise with the
    # reason rather than a bare NotImplementedError.
    with pytest.raises(NotImplementedError, match="not supported on this engine"):
        getattr(book.app, attribute)
    with pytest.raises(NotImplementedError, match="not supported on this engine"):
        setattr(book.app, attribute, True)


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize("attribute", ["path", "startup_path", "version"])
def test_app_unsupported_read_only(book, attribute):
    with pytest.raises(NotImplementedError, match="not supported on this engine"):
        getattr(book.app, attribute)
    # read-only in the public API, so assigning raises AttributeError here too
    with pytest.raises(AttributeError):
        setattr(book.app, attribute, "x")


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_app_quit_not_supported(book):
    with pytest.raises(NotImplementedError, match="can't close the Excel"):
        book.app.quit()


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "axis,expected",
    [
        ("rows", ["rows"]),
        ("r", ["rows"]),
        ("columns", ["columns"]),
        ("c", ["columns"]),
        (None, ["rows", "columns"]),
    ],
)
def test_sheet_autofit(axis, expected):
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0].autofit(axis)
    actions = book.json()["actions"]
    assert [a["func"] for a in actions] == ["setSheetAutofit"] * len(expected)
    assert [a["args"][0] for a in actions] == expected
    assert all(a["sheet_position"] == 0 for a in actions)


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_autofit_invalid_axis(book):
    with pytest.raises(ValueError, match="Invalid axis"):
        book.sheets[0].autofit("diagonal")


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_select():
    # Office.js has no separate select; selecting a sheet activates it.
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[1].select()
    actions = book.json()["actions"]
    assert actions[-1]["func"] == "activateSheet"
    assert book.sheets.active.name == book.sheets[1].name


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_to_html_not_supported(book):
    with pytest.raises(NotImplementedError, match="has no HTML export"):
        book.sheets[0].to_html()


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_copy():
    # The public copy() finds the new sheet by diffing names before/after, so
    # the local api list has to gain it synchronously.
    book = xw.Book(json=json.loads(json.dumps(data)))
    names_before = [sheet.name for sheet in book.sheets]
    copied = book.sheets[0].copy()
    assert copied.name == f"{names_before[0]} (2)"
    assert [sheet.name for sheet in book.sheets] == names_before + [copied.name]
    action = book.json()["actions"][-1]
    assert action["func"] == "copySheet"
    assert action["args"] == ["After", len(names_before) - 1, copied.name]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_copy_before_and_name():
    book = xw.Book(json=json.loads(json.dumps(data)))
    copied = book.sheets[0].copy(before=book.sheets[0], name="mycopy")
    assert copied.name == "mycopy"
    assert [sheet.name for sheet in book.sheets][0] == "mycopy"
    funcs = [a["func"] for a in book.json()["actions"]]
    # copied, then renamed to the requested name by the public method
    assert funcs[-2:] == ["copySheet", "setSheetName"]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_copy_increments_suffix():
    payload = json.loads(json.dumps(data))
    first = payload["sheets"][0]["name"]
    payload["sheets"].append(dict(payload["sheets"][0], name=f"{first} (2)"))
    book = xw.Book(json=payload)
    copied = book.sheets[0].copy()
    assert copied.name == f"{first} (3)"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_sheet_copy_to_other_book_not_supported():
    book = xw.Book(json=json.loads(json.dumps(data)))
    other = xw.Book(json=json.loads(json.dumps(data)))
    with pytest.raises(NotImplementedError, match="different book"):
        book.sheets[0].copy(after=other.sheets[0])


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_book_save():
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.save()
    actions = book.json()["actions"]
    assert actions[-1]["func"] == "save"
    assert actions[-1]["args"] == []


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_book_save_path_not_supported(book):
    # Office.js has no SaveAs, so a path must raise rather than silently
    # saving in place somewhere else.
    with pytest.raises(NotImplementedError, match="can only save the book in place"):
        book.save("/tmp/somewhere.xlsx")


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_book_save_password_not_supported(book):
    with pytest.raises(NotImplementedError, match="password"):
        book.save(password="secret")


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_book_to_pdf_not_supported(book):
    with pytest.raises(NotImplementedError, match="has no PDF export"):
        book.to_pdf()


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_app_calculation_get():
    book = xw.Book(json=json.loads(json.dumps(data)))
    assert book.app.calculation == "automatic"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "value,js_value",
    [
        ("automatic", "Automatic"),
        ("manual", "Manual"),
        ("semiautomatic", "AutomaticExceptTables"),
    ],
)
def test_app_calculation_set(value, js_value):
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.app.calculation = value
    actions = book.app.impl.books.active.json()["actions"]
    assert actions[-1]["func"] == "setCalculation"
    assert actions[-1]["args"] == [js_value]
    # written through, so a read-after-write in the same script is correct
    assert book.app.calculation == value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_app_calculation_set_invalid():
    book = xw.Book(json=json.loads(json.dumps(data)))
    with pytest.raises(ValueError, match="Invalid calculation mode"):
        book.app.calculation = "nonsense"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_app_calculation_get_missing_from_payload():
    # Clients that predate the payload field get a pointed error rather than a
    # KeyError.
    payload = json.loads(json.dumps(data))
    del payload["book"]["calculation"]
    book = xw.Book(json=payload)
    with pytest.raises(NotImplementedError, match="newer version"):
        book.app.calculation


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_app_screen_updating():
    # Office.js can only suspend until its next sync, so the setter queues an
    # action either way and the getter raises.
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.app.screen_updating = False
    actions = book.app.impl.books.active.json()["actions"]
    assert actions[-1]["func"] == "setScreenUpdating"
    assert actions[-1]["args"] == [False]

    book.app.screen_updating = True
    actions = book.app.impl.books.active.json()["actions"]
    assert actions[-1]["args"] == [True]

    with pytest.raises(NotImplementedError, match="suspend screen updating"):
        book.app.screen_updating


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_app_calculate():
    # App queues onto books.active (the convention alert() and selection use),
    # and the App is process-global here, so other tests creating books move
    # books.active. Assert on the App's own active book rather than on a
    # specific one.
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.app.calculate()
    actions = book.app.impl.books.active.json()["actions"]
    assert actions[-1]["func"] == "calculate"
    assert actions[-1]["args"] == []
    assert actions[-1]["sheet_position"] is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_insert_row_range_not_supported(book):
    # Office.js has no InsertRowRange equivalent. Returning None would be
    # indistinguishable from the documented "table isn't empty" answer, so
    # this raises rather than answering wrongly.
    with pytest.raises(NotImplementedError, match="has no equivalent"):
        book.sheets[0].tables[0].insert_row_range


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_display_name_aliases_name():
    # Office.js' Excel.Table only has `name`, so display_name aliases it --
    # keeping scripts that use display_name portable across backends.
    book = xw.Book(json=json.loads(json.dumps(data)))
    table = book.sheets[0].tables[0]
    # not hardcoded: an earlier test renames this table, and Table.name writes
    # through to the module-level `data` dict that the fixture book wraps
    assert table.display_name == table.name

    table.display_name = "myname"
    assert table.display_name == "myname"
    assert table.name == "myname"
    # emits the same action as setting name
    assert book.json()["actions"][0]["func"] == "setTableName"
    assert book.json()["actions"][0]["args"] == [0, "myname"]


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_table_show_table_style_flags(book):
    table1 = book.sheets[0].tables[0]
    table2 = book.sheets[0].tables[1]
    assert table1.show_table_style_first_column is True
    assert table1.show_table_style_last_column is False
    assert table1.show_table_style_row_stripes is True
    assert table1.show_table_style_column_stripes is False
    # Table2 is the inverse of Table1, so a getter reading the wrong field
    # can't pass by accident.
    assert table2.show_table_style_first_column is False
    assert table2.show_table_style_last_column is True
    assert table2.show_table_style_row_stripes is False
    assert table2.show_table_style_column_stripes is True


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "attribute,func,value",
    [
        ("show_table_style_first_column", "showTableStyleFirstColumn", False),
        ("show_table_style_last_column", "showTableStyleLastColumn", True),
        ("show_table_style_row_stripes", "showTableStyleRowStripes", False),
        ("show_table_style_column_stripes", "showTableStyleColumnStripes", True),
    ],
)
def test_table_set_show_table_style_flags(book, attribute, func, value):
    table = book.sheets[0].tables[0]
    setattr(table, attribute, value)
    assert book.json() == {
        "actions": [
            {
                "func": func,
                "args": [0, value],
                "values": None,
                "sheet_position": 0,
                "start_row": None,
                "start_column": None,
                "row_count": None,
                "column_count": None,
            },
        ]
    }
    # written through, so a read-after-write in the same script is correct
    assert getattr(table, attribute) is value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_tables_add_seeds_the_requested_values():
    # A table added mid-script has to read back as what the caller asked for,
    # not as placeholders: name was hardcoded to "" here, so table.name and
    # display_name came back empty until the next round-trip.
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    table = sheet.tables.add(
        source=sheet["A1:B2"], name="MyTable", table_style_name="TableStyleLight1"
    )
    assert table.name == "MyTable"
    assert table.display_name == "MyTable"
    assert table.range.address == "$A$1:$B$2"
    assert table.table_style == "TableStyleLight1"
    assert table.show_headers is True


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_tables_add_without_a_name_leaves_it_empty():
    # Excel assigns the name in that case, so there's nothing to seed until
    # the payload refreshes.
    book = xw.Book(json=json.loads(json.dumps(data)))
    sheet = book.sheets[0]
    assert sheet.tables.add(source=sheet["A1:B2"]).name == ""


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_table_add_seeds_show_table_style_flags():
    # Tables.add() seeds the flags with Excel's defaults, so the getters work
    # before the next round-trip rather than raising KeyError.
    book = xw.Book(json=json.loads(json.dumps(data)))
    table = book.sheets[0].tables.add(source=book.sheets[0]["A1:B2"])
    assert table.show_table_style_first_column is False
    assert table.show_table_style_last_column is False
    assert table.show_table_style_row_stripes is True
    assert table.show_table_style_column_stripes is False
    assert table.show_autofilter is True


# Lazy loading: these methods are only supported in xlwings Lite
# and should raise NotImplementedError on all other platforms/engines.


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "getter",
    [
        "get_formula_array",
        "get_number_format",
        "get_color",
        "get_wrap_text",
        "get_column_width",
        "get_row_height",
        "get_left",
        "get_top",
        "get_width",
        "get_height",
        "get_current_region",
        "get_merge_area",
        "get_merge_cells",
        "get_table",
        "get_hyperlink",
    ],
)
def test_range_async_getters_not_supported(book, getter):
    # Lite-only, like get_value()/get_formula().
    with pytest.raises(NotImplementedError):
        asyncio.run(getattr(book.sheets[0].range("A1"), getter)())


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "prop,hint",
    [
        ("color", "get_color()"),
        ("column_width", "get_column_width()"),
        ("row_height", "get_row_height()"),
        ("wrap_text", "get_wrap_text()"),
        ("formula_array", "get_formula_array()"),
        ("number_format", "get_number_format()"),
        ("left", "get_left()"),
        ("top", "get_top()"),
        ("width", "get_width()"),
        ("height", "get_height()"),
        ("current_region", "get_current_region()"),
        ("merge_area", "get_merge_area()"),
        ("merge_cells", "get_merge_cells()"),
        ("table", "get_table()"),
    ],
)
def test_range_sync_getters_point_at_async(book, prop, hint):
    # The sync properties raise, but name the async method to use instead.
    with pytest.raises(NotImplementedError, match=re.escape(hint)):
        getattr(book.sheets[0].range("A1"), prop)


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "getter,reported",
    [
        # COM reports None for a range whose cells don't agree; Office.js does
        # the same, so these pass it through rather than inventing a default.
        ("get_column_width", None),
        ("get_row_height", None),
        ("get_wrap_text", None),
        ("get_number_format", None),
        ("get_formula_array", None),
        ("get_color", None),
        # ...and the ordinary uniform case still comes back as-is
        ("get_row_height", 15.0),
        ("get_wrap_text", True),
        ("get_number_format", "General"),
        ("get_formula_array", "{=SUM(A1:B2)}"),
    ],
)
def test_range_group_a_getters_pass_through(book, getter, reported):
    rng = book.sheets[0].range("A1:B2")

    async def fake(self, key, method=None):
        return reported

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        assert asyncio.run(getattr(rng, getter)()) == reported


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
@pytest.mark.parametrize(
    "kwargs,expected",
    [
        ({}, "$A$1:$C$3"),
        ({"row_absolute": False}, "$A1:$C3"),
        ({"column_absolute": False}, "A$1:C$3"),
        ({"row_absolute": False, "column_absolute": False}, "A1:C3"),
        # the fixture's sheet is "Sheet 1", so Excel quotes the prefix
        ({"include_sheetname": True}, "'Sheet 1'!$A$1:$C$3"),
        ({"external": True}, "'[engines.xlsm]Sheet 1'!$A$1:$C$3"),
    ],
)
def test_range_get_address(book, kwargs, expected):
    # Purely local: the engine already knows the coordinates.
    assert book.sheets[0].range((1, 1), (3, 3)).get_address(**kwargs) == expected


@pytest.mark.skipif(engine == "calamine", reason="unsupported by calamine")
def test_range_get_address_single_cell(book):
    rng = book.sheets[0].range((1, 1))
    assert rng.get_address() == "$A$1"
    assert rng.get_address(False, False) == "A1"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_get_address_matches_address(book):
    # The address property is the get_address() default.
    for arg1, arg2 in [((1, 1), None), ((1, 1), (3, 3)), ((2, 3), (5, 7))]:
        rng = book.sheets[0].range(arg1, arg2) if arg2 else book.sheets[0].range(arg1)
        assert rng.address == rng.get_address()


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_get_address_quotes_names_with_spaces():
    # Excel quotes the prefix when the book or sheet name contains a space.
    payload = json.loads(json.dumps(data))
    payload["book"]["name"] = "My Book.xlsx"
    book = xw.Book(json=payload)
    rng = book.sheets[0].range((1, 1))
    assert rng.get_address(external=True) == "'[My Book.xlsx]Sheet 1'!$A$1"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_hyperlink_sync_points_at_async(book):
    # The public hyperlink property reads .formula first, so it raises that
    # property's message; both point at an async method to use instead.
    with pytest.raises(NotImplementedError, match=r"get_formula\(\)"):
        book.sheets[0].range("A1").hyperlink
    # The impl property names get_hyperlink() for anyone reaching it directly.
    with pytest.raises(NotImplementedError, match=r"get_hyperlink\(\)"):
        book.sheets[0].range("A1").impl.hyperlink


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "formula,hyperlink,expected",
    [
        # set pragmatically: comes from the cell's hyperlink object
        ("plain text", "http://xlwings.org", "http://xlwings.org"),
        # a HYPERLINK() formula keeps the target in the formula string
        ('=HYPERLINK("http://xlwings.org","xlwings")', None, "http://xlwings.org"),
    ],
)
def test_range_get_hyperlink(book, formula, hyperlink, expected):
    rng = book.sheets[0].range("A1")
    fetched = {"formulas": [[formula]], "hyperlink": hyperlink}

    async def fake(self, key, method=None):
        return fetched[key]

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        assert asyncio.run(rng.get_hyperlink()) == expected


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "formula,hyperlink",
    [
        ("plain text", None),  # no hyperlink at all
        ("=SUM(A1:B2)", None),  # a formula, but not a HYPERLINK() one
    ],
)
def test_range_get_hyperlink_raises_without_one(book, formula, hyperlink):
    # The desktop engines raise rather than returning None, so this does too.
    rng = book.sheets[0].range("A1")
    fetched = {"formulas": [[formula]], "hyperlink": hyperlink}

    async def fake(self, key, method=None):
        return fetched[key]

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        with pytest.raises(Exception, match="doesn't seem to contain a hyperlink"):
            asyncio.run(rng.get_hyperlink())


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_scalar_getters_are_not_matrices(book):
    # number_format and formula_array are single strings on every engine, so
    # they must not go through the value-shaping stage.
    rng = book.sheets[0].range("A1:B2")

    async def fake(self, key, method=None):
        return "General" if key == "number_format" else "{=SUM(A1:B2)}"

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        assert asyncio.run(rng.get_number_format()) == "General"
        assert asyncio.run(rng.get_formula_array()) == "{=SUM(A1:B2)}"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_matrix_getters_apply_ndim(book):
    # The matrix-valued getters reuse the stage .value goes through, so the
    # shape rules and options(ndim=...) match reading values.
    rng = book.sheets[0].range("A1:B2")

    async def fake(self):
        return [["#,##0", "0.00"], ["General", "General"]]

    with mock.patch.object(type(rng.impl), "get_formula", fake, create=True):
        assert asyncio.run(rng.get_formula()) == [
            ["#,##0", "0.00"],
            ["General", "General"],
        ]

    single = book.sheets[0].range("A1")

    async def fake_single(self):
        return [["#,##0"]]

    with mock.patch.object(type(single.impl), "get_formula", fake_single, create=True):
        # a single cell squeezes to a scalar...
        assert asyncio.run(single.get_formula()) == "#,##0"
        # ...unless ndim asks otherwise
        assert asyncio.run(single.options(ndim=2).get_formula()) == [["#,##0"]]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_rows_and_columns(book):
    # rows/columns are built from the range itself in main.py and never touch
    # the impl, so they already work on this engine.
    rng = book.sheets[0].range("A1:B3")
    assert len(rng.rows) == 3
    assert len(rng.columns) == 2
    assert rng.rows[0].address == "$A$1:$B$1"
    assert rng.columns[1].address == "$B$1:$B$3"


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_group_b_getters_build_objects(book):
    # These fetch an address (or a name) and build the object locally.
    rng = book.sheets[0].range("A1")
    # not hardcoded: an earlier test renames this table, and Table.name writes
    # through to the module-level `data` dict the fixture book wraps
    table_name = book.sheets[0].tables[0].name
    fetched = {
        "current_region": "$A$1:$C$5",
        "merge_area": "$A$1:$B$1",
        "merge_cells": True,
        "table": table_name,
    }

    async def fake(self, key, method=None):
        return fetched[key]

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        assert asyncio.run(rng.get_current_region()).address == "$A$1:$C$5"
        assert asyncio.run(rng.get_merge_area()).address == "$A$1:$B$1"
        assert asyncio.run(rng.get_merge_cells()) is True
        assert asyncio.run(rng.get_table()).name == table_name


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize("reported", [True, False, None])
def test_range_merge_cells_is_tristate(book, reported):
    # Mirrors COM's Range.MergeCells: True when the whole range is merged,
    # False when none of it is, None when it's only partly merged.
    rng = book.sheets[0].range("A1:C1")

    async def fake(self, key, method=None):
        return reported

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        assert asyncio.run(rng.get_merge_cells()) is reported


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_merge_area_and_table_fall_back(book):
    # An unmerged cell reports no merged area; xlwings returns the cell itself.
    # A range outside any table reports no table; xlwings returns None.
    rng = book.sheets[0].range("A1")

    async def fake(self, key, method=None):
        return None

    with mock.patch.object(type(rng.impl), "_get_range_data", fake):
        assert asyncio.run(rng.get_merge_area()).address == "$A$1"
        assert asyncio.run(rng.get_table()) is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "attribute,value",
    [
        ("bold", True),
        ("italic", True),
        ("size", 13),
        ("color", "#ff0000"),
        ("name", "Calibri"),
    ],
)
def test_font_setters(book, attribute, value):
    setattr(book.sheets[0].range("A1").font, attribute, value)
    action = book.json()["actions"][-1]
    assert action["func"] == "setFontProperty"
    assert action["args"] == [attribute, value]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "value,expected",
    [
        # the forms the public API documents, matching the desktop engines
        ((255, 255, 255), "#ffffff"),
        ((255, 0, 0), "#ff0000"),
        ("#FFA500", "#FFA500"),
        ("FFA500", "#FFA500"),  # a missing "#" is normalized
        (255, "#ff0000"),  # Excel colour constant (little-endian int)
    ],
)
def test_range_color_setter_accepts_every_form(value, expected):
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0].range("A1").color = value
    action = book.json()["actions"][-1]
    assert action["func"] == "setRangeColor"
    assert action["args"] == [expected]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_color_setter_none_clears():
    # Documented: setting None removes the background.
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0].range("A1").color = None
    assert book.json()["actions"][-1]["args"] == [None]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize("value", [(1, 2), (1, 2, 3, 4), object()])
def test_range_color_setter_rejects_invalid(book, value):
    with pytest.raises(ValueError, match="Color must be an RGB tuple"):
        book.sheets[0].range("A1").color = value


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "value,expected",
    [((255, 0, 0), "#ff0000"), ("#00ff00", "#00ff00")],
)
def test_font_color_setter_accepts_tuple_and_hex(value, expected):
    book = xw.Book(json=json.loads(json.dumps(data)))
    book.sheets[0].range("A1").font.color = value
    action = book.json()["actions"][-1]
    assert action["func"] == "setFontProperty"
    assert action["args"] == ["color", expected]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_font_color_setter_rejects_invalid(book):
    with pytest.raises(ValueError, match="Color must be an RGB tuple"):
        book.sheets[0].range("A1").font.color = (1, 2)


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "getter,expected",
    [
        ("get_bold", True),
        ("get_italic", False),
        ("get_size", 13.0),
        ("get_color", (255, 0, 0)),
        ("get_name", "Calibri"),
    ],
)
def test_font_async_getters(book, getter, expected):
    # All five come from one Office.js object, so they're fetched together.
    font = book.sheets[0].range("A1").font

    async def fake(self, key, method=None):
        assert key == "font"
        return {
            "bold": True,
            "italic": False,
            "size": 13.0,
            "color": "#ff0000",
            "name": "Calibri",
        }

    with mock.patch.object(xw.pro._xlremote.Range, "_get_range_data", fake):
        assert asyncio.run(getattr(font, getter)()) == expected


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_font_color_none_when_unset(book):
    font = book.sheets[0].range("A1").font

    async def fake(self, key, method=None):
        return {
            "bold": None,
            "italic": None,
            "size": None,
            "color": None,
            "name": None,
        }

    with mock.patch.object(xw.pro._xlremote.Range, "_get_range_data", fake):
        # None is also what a range whose cells disagree reports
        assert asyncio.run(font.get_color()) is None
        assert asyncio.run(font.get_bold()) is None


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.parametrize(
    "prop,hint",
    [
        ("bold", "get_bold()"),
        ("italic", "get_italic()"),
        ("size", "get_size()"),
        ("color", "get_color()"),
        ("name", "get_name()"),
    ],
)
def test_font_sync_getters_point_at_async(book, prop, hint):
    with pytest.raises(NotImplementedError, match=re.escape(hint)):
        getattr(book.sheets[0].range("A1").font, prop)


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_font_on_non_range_parent_raises(book):
    # Shape.font and Characters.font would land here; Office.js only exposes
    # font attributes through a Range on this engine.
    from xlwings.pro._xlremote import Font

    font = Font(parent=object(), api={})
    with pytest.raises(NotImplementedError, match="only supported on a Range"):
        font.bold = True
    with pytest.raises(NotImplementedError, match="only supported on a Range"):
        asyncio.run(font.get_bold())


def test_get_value_not_supported(book):
    with pytest.raises(NotImplementedError):
        asyncio.run(book.sheets[0].range("A1").get_value())


def test_sheet_load_not_supported(book):
    with pytest.raises(NotImplementedError):
        asyncio.run(book.sheets[0].load())


def test_book_load_not_supported(book):
    with pytest.raises(NotImplementedError):
        asyncio.run(book.load())


def test_books_get_active_not_supported(book):
    with pytest.raises(NotImplementedError):
        asyncio.run(book.app.books.get_active())


def test_sheets_get_active_not_supported(book):
    with pytest.raises(NotImplementedError):
        asyncio.run(book.sheets.get_active())


def test_app_get_selection_not_supported(book):
    with pytest.raises(NotImplementedError):
        asyncio.run(book.app.get_selection())


# Pipeline.async_call()


def test_pipeline_async_call_mixed_stages():
    """Pipeline.async_call() handles both sync and async stages in order."""
    from xlwings.conversion.framework import Pipeline

    log = []

    class SyncStage:
        def __call__(self, ctx):
            log.append("sync")
            ctx["value"] += 1

    class AsyncStage:
        async def __call__(self, ctx):
            log.append("async")
            ctx["value"] += 10

    pipeline = Pipeline([SyncStage(), AsyncStage(), SyncStage()])
    ctx = {"value": 0}
    asyncio.run(pipeline.async_call(ctx))
    assert ctx["value"] == 12
    assert log == ["sync", "async", "sync"]
