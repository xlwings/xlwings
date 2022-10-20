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
engine = os.environ.get("XLWINGS_ENGINE") or "calamine"
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
                [1.0, 2.0, 3.0, "2021-01-01T00:00:00.000Z"],
                [4.0, 5.0, 6.0, ""],
                ["", "", "", ""],
            ],
            "pictures": [
                {
                    "name": "pic1",
                    "height": 10,
                    "width": 20,
                },
                {
                    "name": "pic2",
                    "height": 30,
                    "width": 40,
                },
            ],
        },
        {"name": "Sheet2", "values": [["aa", "bb"], [11.0, 22.0]], "pictures": []},
        {
            "name": "Sheet3",
            "values": [
                ["", "string"],
                [-1.0, 1.0],
                [True, False],
                ["2021-10-01T00:00:00.000Z", "2021-12-31T23:35:00.000Z"],
            ],
            "pictures": [],
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


# range.value
def test_range_index(book):
    sheet = book.sheets[0]
    assert sheet.range((1, 1)).value == "a"
    assert sheet.range((1, 1), (3, 1)).value == ["a", 1.0, 4.0]
    assert sheet.range((1, 3), (3, 3)).value == ["c", 3.0, 6.0]
    assert sheet.range((1, 1), (3, 3)).value == [
        ["a", "b", "c"],
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
    ]
    assert sheet.range((2, 2), (3, 3)).value == [[2.0, 3.0], [5.0, 6.0]]


def test_range_a1(book):
    sheet = book.sheets[0]
    assert sheet.range("A1").value == "a"
    assert sheet.range("A1:A3").value == ["a", 1.0, 4.0]
    assert sheet.range("C1:C3").value == ["c", 3.0, 6.0]
    assert sheet.range("A1:C3").value == [
        ["a", "b", "c"],
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
    ]
    assert sheet.range("B2:C3").value == [[2.0, 3.0], [5.0, 6.0]]


def test_range_shortcut_address(book):
    sheet = book.sheets[0]
    assert sheet["A1"].value == "a"
    assert sheet["A1:A3"].value == ["a", 1.0, 4.0]
    assert sheet["C1:C3"].value == ["c", 3.0, 6.0]
    assert sheet["A1:C3"].value == [["a", "b", "c"], [1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
    assert sheet["B2:C3"].value == [[2.0, 3.0], [5.0, 6.0]]


def test_range_shortcut_index(book):
    sheet = book.sheets[0]
    assert sheet[0, 0].value == "a"
    assert sheet[0:3, 0].value == ["a", 1.0, 4.0]
    assert sheet[0:3, 2].value == ["c", 3.0, 6.0]
    assert sheet[0:3, 0:3].value == [["a", "b", "c"], [1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
    assert sheet[1:3, 1:3].value == [[2.0, 3.0], [5.0, 6.0]]


def test_range_from_range(book):
    sheet = book.sheets[0]
    assert sheet.range(sheet.range((1, 1)), sheet.range((3, 1))).value == [
        "a",
        1.0,
        4.0,
    ]
    assert sheet.range(sheet.range("C1"), sheet.range("C3")).value == ["c", 3.0, 6.0]
    assert sheet.range(sheet.range("A1"), sheet.range("C3")).value == [
        ["a", "b", "c"],
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
    ]
    assert sheet.range(sheet.range("B2"), sheet.range("C3")).value == [
        [2.0, 3.0],
        [5.0, 6.0],
    ]


def test_range_round_indexing(book):
    sheet = book.sheets[0]
    assert sheet["B2:C3"](1, 1).value == 2.0
    assert sheet["B2:C3"](1, 1).address == "$B$2"
    assert sheet["B2:C3"](2, 1).value == 5.0
    assert sheet["B2:C3"](2, 1).address == "$B$3"


def test_range_square_indexing_2d(book):
    sheet = book.sheets[0]
    assert sheet["B2:C3"][0, 0].value == 2.0
    assert sheet["B2:C3"][0, 0].address == "$B$2"
    assert sheet["B2:C3"][1, 0].value == 5.0
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
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
    ]
    assert sheet1["B1"].expand().address == "$B$1:$C$3"
    assert sheet1["B1"].expand().value == [["b", "c"], [2.0, 3.0], [5.0, 6.0]]
    assert sheet1["C3"].expand().address == "$C$3"
    assert sheet1["C3"].expand().value == 6.0

    # Edge case (no more rows/cols after expanded range
    sheet2 = book.sheets[1]
    assert sheet2["A1"].expand().value == [["aa", "bb"], [11.0, 22.0]]
    assert sheet2["A1"].expand().address == "$A$1:$B$2"


def test_completely_outside_usedrange(book):
    sheet = book.sheets[0]
    assert sheet["D5"].value is None
    assert sheet["D5:D6"].value == [None, None]
    assert sheet["D5:E7"].value == [[None, None], [None, None], [None, None]]


def test_partly_outside_usedrange(book):
    sheet = book.sheets[0]
    assert sheet["A4:A5"].value == [None, None]
    assert sheet["A3:A5"].value == [4, None, None]
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
        sheet["B2:C3"].options(np.array).value, np.array([[2.0, 3.0], [5.0, 6.0]])
    )


@pytest.mark.skipif(not pd, reason="requires pandas")
def test_pandas_df(book):
    sheet = book.sheets[0]
    pd.testing.assert_frame_equal(
        sheet["A1:C3"].options(pd.DataFrame, index=False).value,
        pd.DataFrame(data=[[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]], columns=["a", "b", "c"]),
    )


@pytest.mark.skipif(
    file_extension != "xlsx" and engine == "calamine", reason="datetime unsupported"
)
def test_read_basic_types(book):
    sheet = book.sheets[2]
    assert sheet["A1:B4"].value == [
        [None, "string"],
        [-1.0, 1.0],
        [True, False],
        [dt.datetime(2021, 10, 1, 0, 0), dt.datetime(2021, 12, 31, 23, 35)],
    ]


def test_read_basic_types_no_datetime(book):
    sheet = book.sheets[2]
    assert sheet["A1:B3"].value == [
        [None, "string"],
        [-1.0, 1.0],
        [True, False],
    ]


@pytest.mark.skipif(engine != "remote", reason="requires remote engine")
@pytest.mark.skipif(not tz, reason="requires dateutil")
def test_write_basic_types(book):
    sheet = book.sheets[0]
    sheet["Z10"].value = [
        [None, "string"],
        [-1.0, 1.0],
        [True, False],
        [
            dt.date(2021, 10, 1),
            dt.datetime(2021, 12, 31, 23, 35, tzinfo=tz.gettz("Europe/Paris")),
        ],
    ]
    assert (
        json.dumps(book.json()["actions"][0]["values"])
        == '[["", "string"], [-1.0, 1.0], [true, false], '
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
    assert book.sheets[0].pictures[0].name == "pic1"
    assert book.sheets[0].pictures[1].name == "pic2"
    assert book.sheets[0].pictures(1).name == "pic1"
    assert book.sheets[0].pictures(2).name == "pic2"


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
    assert sheet.pictures[2].impl.index == 3
    sheet.pictures["new"].delete()
    assert len(sheet.pictures) == 2


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_iter(book):
    sheet = book.sheets[0]
    pic_names = []
    for pic in sheet.pictures:
        pic_names.append(pic.name)
    assert pic_names == ["pic1", "pic2"]


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_pictures_contains(book):
    sheet = book.sheets[0]
    assert "pic1" in sheet.pictures
    assert 1 in sheet.pictures
    assert "pic2" in sheet.pictures
    assert 2 in sheet.pictures
    assert "no" not in sheet.pictures
    assert 3 not in sheet.pictures


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_empty_pictures(book):
    assert not book.sheets[1].pictures


@pytest.mark.skipif(engine == "calamine", reason="calamine engine")
def test_picture_exists(book):
    with pytest.raises(xw.ShapeAlreadyExists):
        book.sheets[0].pictures.add(this_dir.parent / "sample_picture.png", name="pic1")


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
    book.impl._json = {"actions": []}
    sheet1 = book.sheets[0]
    assert sheet1["one"].value == "a"
    sheet1["one"].value = 1000
    assert book.json()["actions"][0]["values"] == [[1000]]
    assert book.json()["actions"][0]["sheet_position"] == 0
    assert book.json()["actions"][0]["start_row"] == 0
    assert book.json()["actions"][0]["start_column"] == 0


# Names collection
@pytest.mark.skipif(engine == "remote", reason="TODO: remote")
def test_names_len(book):
    assert len(book.names) == 3


@pytest.mark.skipif(engine == "remote", reason="TODO: remote")
def test_names_index_vs_name(book):
    assert book.names[0].name == "one"
    assert book.names["one"].name == "one"


@pytest.mark.skipif(engine != "excel", reason="TODO: calamine, remote")
def test_name_local_scope(book):
    assert book.names[1].name == "Sheet1!two"


@pytest.mark.skipif(engine == "remote", reason="TODO: remote")
def test_name_refers_to(book):
    assert book.names[0].refers_to == "=Sheet1!$A$1"


@pytest.mark.skipif(engine == "remote", reason="TODO: remote")
def test_name_refers_to_range(book):
    assert book.names[0].refers_to_range == book.sheets[0]["A1"]
    assert book.names[1].refers_to_range == book.sheets[0]["C7:D8"]
    assert book.names[2].refers_to_range == book.sheets[1]["A1:A2"]


@pytest.mark.skipif(engine == "remote", reason="TODO: remote")
def test_name_contains(book):
    assert "one" in book.names


@pytest.mark.skipif(engine == "remote", reason="TODO: remote")
def test_names_iter(book):
    for ix, name in enumerate(book.names):
        if ix == 0:
            assert name.refers_to_range == book.sheets[0]["A1"]
        elif ix == 1:
            assert name.refers_to_range == book.sheets[0]["C7:D8"]
        elif ix == 2:
            assert name.refers_to_range == book.sheets[1]["A1:A2"]
