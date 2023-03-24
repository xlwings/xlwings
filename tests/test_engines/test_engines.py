import datetime as dt
import json
import os
from pathlib import Path

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
# "xlsx", "xlsb", or "xls"
file_extension = os.environ.get("XLWINGS_FILE_EXTENSION") or "xlsx"


data = {
    "client": "Microsoft Office Scripts",
    "version": xw.__version__,
    "book": {
        "name": f"engines.{file_extension}",
        "active_sheet_index": 0,
        "selection": "B3:B4",
    },
    "names": [
        {"name": "one", "sheet_index": 0, "address": "A1", "book_scope": True},
        {
            "name": "Sheet1!two",
            "sheet_index": 0,
            "address": "C7:D8",
            "book_scope": False,
        },
        {"name": "two", "sheet_index": 1, "address": "A1:A2", "book_scope": True},
    ],
    "sheets": [
        {
            "name": "Sheet1",
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
            "pictures": [
                {
                    "name": "mypic1",
                    "height": 10,
                    "width": 20,
                },
                {
                    "name": "mypic2",
                    "height": 30,
                    "width": 40,
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
                },
            ],
        },
        {"name": "Sheet2", "values": [["aa", "bb"], [11.1, 22.2]], "pictures": []},
        {
            "name": "Sheet3",
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


@pytest.mark.skipif(
    file_extension != "xlsx" and engine == "calamine", reason="datetime unsupported"
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
        '["2021-10-01T00:00:00", "2021-12-31T23:35:00"]]'
    )


# sheets
def test_sheet_access(book):
    assert book.sheets[0] == book.sheets["Sheet1"]
    assert book.sheets[1] == book.sheets["Sheet2"]
    assert book.sheets[0].name == "Sheet1"
    assert book.sheets[1].name == "Sheet2"


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_sheet_active(book):
    assert book.sheets.active == book.sheets[0]


def test_sheets_iteration(book):
    for ix, sheet in enumerate(book.sheets):
        assert sheet.name == "Sheet1" if ix == 0 else "Sheet2"


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
    sheet2 = book.sheets[1]
    assert sheet1["one"].address == "$A$1"
    assert sheet2["two"].address == "$A$1:$A$2"


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
    assert len(book.names) == 3


def test_names_index_vs_name(book):
    assert book.names[0].name == "one"
    assert book.names["one"].name == "one"


@pytest.mark.skipif(engine == "calamine", reason="doesn't support local scope yet")
def test_name_local_scope1(book):
    assert book.names[1].name == "Sheet1!two"


@pytest.mark.skipif(engine == "calamine", reason="doesn't support local scope yet")
def test_name_local_scope2(book):
    assert book.sheets["Sheet1"].names[0].name == "Sheet1!two"


def test_name_refers_to(book):
    assert book.names[0].refers_to == "=Sheet1!$A$1"


def test_name_refers_to_range(book):
    assert book.names[0].refers_to_range == book.sheets[0]["A1"]
    assert book.names[1].refers_to_range == book.sheets[0]["C7:D8"]
    assert book.names[2].refers_to_range == book.sheets[1]["A1:A2"]


def test_name_contains(book):
    assert "one" in book.names


def test_names_iter(book):
    for ix, name in enumerate(book.names):
        if ix == 0:
            assert name.refers_to_range == book.sheets[0]["A1"]
        elif ix == 1:
            assert name.refers_to_range == book.sheets[0]["C7:D8"]
        elif ix == 2:
            assert name.refers_to_range == book.sheets[1]["A1:A2"]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
def test_range_get_name(book):
    assert book.sheets[0]["A1"].name == book.names[0]
    assert book.sheets[0]["C7:D8"].name == book.names[1]
    assert book.sheets[1]["A1:A2"].name == book.names[2]
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
def test_sheet_name_delete(book):
    book.names[0].delete()
    assert book.json()["actions"][0]["func"] == "nameDelete"
    assert book.json()["actions"][0]["args"] == ["one", "=Sheet1!$A$1"]


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
