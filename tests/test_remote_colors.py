"""Bulk direct-fill reads through the public API and the Pyodide bridge."""

import asyncio
import copy
import sys
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings import base_classes
from xlwings.pro import _xlremote as R


def make_book():
    impl = R.App(R.Apps(), add_book=False).books.open(
        {
            "client": "Office.js",
            "version": xw.__version__,
            "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
            "names": [],
            "sheets": [{"name": "S", "values": [[]], "pictures": [], "tables": []}],
        },
        lazy=True,
    )
    return xw.Book(impl=impl)


def install_bridge(monkeypatch, colors=None, error=None):
    calls = []

    async def get_range_data(sheet, address, keys):
        calls.append((sheet, address, keys))
        if error:
            raise error
        return SimpleNamespace(to_py=lambda: {"colors": colors})

    js = ModuleType("js")
    js.xlwings = SimpleNamespace(getRangeData=get_range_data)
    ffi = ModuleType("pyodide.ffi")
    ffi.to_js = lambda value: value
    monkeypatch.setitem(sys.modules, "js", js)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    monkeypatch.setattr(sys, "platform", "emscripten")
    return calls


@pytest.mark.parametrize(
    "address, colors, expected",
    [
        ("B3", [[None]], [[None]]),
        ("B3", [["#FFFFFF"]], [[(255, 255, 255)]]),
        ("B3:C3", [["#ff0000", None]], [[(255, 0, 0), None]]),
        ("B3:B4", [["#000000"], [None]], [[(0, 0, 0)], [None]]),
        (
            "B3:C4",
            [["#ff0000", None], ["#FFFFFF", "#123456"]],
            [[(255, 0, 0), None], [(255, 255, 255), (18, 52, 86)]],
        ),
    ],
)
@pytest.mark.parametrize("options", [{}, {"ndim": 1}, {"ndim": 2, "transpose": True}])
def test_colors_keep_shape_and_do_not_load_values(
    monkeypatch, address, colors, expected, options
):
    book = make_book()
    before = copy.deepcopy(book.impl.json())
    calls = install_bridge(monkeypatch, colors)

    result = asyncio.run(book.sheets[0][address].options(**options).get_colors())

    assert result == expected
    assert calls == [("S", book.sheets[0][address].address, ["colors"])]
    assert book.impl.json() == before


def test_colors_do_not_flush_pending_writes(monkeypatch):
    book = make_book()
    rng = book.sheets[0]["A1"]
    rng.color = "#ff0000"
    before = copy.deepcopy(book.impl.json())
    install_bridge(monkeypatch, [[None]])

    assert asyncio.run(rng.get_colors()) == [[None]]
    assert book.impl.json() == before


def test_colors_propagate_host_error(monkeypatch):
    book = make_book()
    error = RuntimeError("range_too_large")
    install_bridge(monkeypatch, error=error)
    with pytest.raises(RuntimeError) as caught:
        asyncio.run(book.sheets[0]["A1"].get_colors())
    assert caught.value is error


def test_colors_require_lite_on_remote():
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(make_book().sheets[0]["A1"].get_colors())


def test_colors_require_lite_on_other_engines():
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(xw.Range(impl=base_classes.Range()).get_colors())


def test_set_colors_queues_one_shape_preserving_action():
    book = make_book()
    rng = book.sheets[0]["B3:D4"]

    assert rng.set_colors([[(0, 128, 0), ..., None], ["#FFFFFF", 255, ...]]) is None

    (action,) = book.impl.json()["actions"]
    assert action["func"] == "setRangeColors"
    assert action["args"] == [
        [["#008000", "keep", None], ["#ffffff", "#ff0000", "keep"]]
    ]
    assert action["values"] is None
    assert (action["start_row"], action["start_column"]) == (2, 1)
    assert (action["row_count"], action["column_count"]) == (2, 3)


def test_set_colors_single_cell_and_all_unchanged():
    book = make_book()
    rng = book.sheets[0]["A1"]
    rng.set_colors([[...]])
    assert book.impl.json()["actions"] == []
    rng.set_colors([[None]])
    assert book.impl.json()["actions"][0]["args"] == [[[None]]]


def test_set_colors_preserves_remote_null_selection_no_op():
    book = make_book()
    rng = xw.Range(impl=R.Range(sheet=book.sheets[0].impl, arg1=None))
    assert rng.shape == (1, 1)
    assert rng.set_colors([]) is None
    assert rng.set_colors([[(255, 0, 0)]]) is None
    assert book.impl.json()["actions"] == []


def test_set_colors_accepts_rgb_lists_and_excel_packed_integers():
    book = make_book()
    book.sheets[0]["A1:C1"].set_colors([[[0, 128, 0], 0x0000FF, 0xFF0000]])
    assert book.impl.json()["actions"][0]["args"] == [
        [["#008000", "#ff0000", "#0000ff"]]
    ]


@pytest.mark.parametrize(
    "colors",
    [
        (255, 0, 0),
        [["#ff0000"]],
        [["#ff0000", None], [None, None]],
        [["red", None]],
        [["#ff0000", "red"]],
        [[(256, 0, 0), None]],
        [[(True, 0, 0), None]],
        [[True, None]],
        [[-1, None]],
    ],
)
def test_set_colors_rejects_invalid_input_without_queuing(colors):
    book = make_book()
    with pytest.raises(ValueError):
        book.sheets[0]["A1:B1"].set_colors(colors)
    assert book.impl.json()["actions"] == []


def test_set_colors_rejects_large_range_without_queuing():
    book = make_book()
    with pytest.raises(ValueError, match="10,000"):
        book.sheets[0]["A1:A10001"].set_colors([])
    assert book.impl.json()["actions"] == []
