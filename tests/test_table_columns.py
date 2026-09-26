"""Table column API validation and remote action payloads."""

import asyncio
import copy
import sys
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings.pro import _xlremote as remote


@pytest.fixture
def book():
    payload = {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "S",
                "values": [[None] * 5 for _ in range(5)],
                "pictures": [],
                "tables": [
                    {
                        "name": "Sales",
                        "range_address": "B2:C4",
                        "row_count": 3,
                        "column_count": 2,
                        "columns": ["Item", "Amount"],
                        "header_row_range_address": "B2:C2",
                        "data_body_range_address": "B3:C4",
                        "total_row_range_address": None,
                        "show_headers": True,
                        "show_totals": False,
                    }
                ],
            }
        ],
    }
    impl = remote.App(remote.Apps(), add_book=False).books.open(payload)
    return xw.Book(impl=impl)


def test_lookup_and_one_based_indexes(book):
    columns = book.sheets[0].tables["Sales"].columns
    assert len(columns) == 2
    assert columns[0].name == "Item"
    assert columns[1].index == 2
    assert columns["Amount"].index == 2
    with pytest.raises(KeyError):
        _ = columns["Missing"]


def test_add_and_delete_queue_native_actions(book):
    columns = book.sheets[0].tables["Sales"].columns
    added = columns.add("Margin", index=2)
    assert added.index == 2
    assert [column.name for column in columns] == ["Item", "Margin", "Amount"]
    action = book.impl.json()["actions"][-1]
    assert action["func"] == "addTableColumn"
    assert action["args"] == [0, 1, "Margin"]
    columns["Margin"].delete()
    action = book.impl.json()["actions"][-1]
    assert action["func"] == "deleteTableColumn"
    assert action["args"] == [0, "Margin"]
    assert [column.name for column in columns] == ["Item", "Amount"]
    assert columns.add("Last").index == 3


@pytest.mark.parametrize("index", [0, 4, -1, 1.5, True, "2"])
def test_invalid_index_does_not_queue(book, index):
    columns = book.sheets[0].tables[0].columns
    before = len(book.impl.json()["actions"])
    with pytest.raises((TypeError, IndexError)):
        columns.add("New", index=index)
    assert len(book.impl.json()["actions"]) == before


@pytest.mark.parametrize("name", ["", "  ", 1, "amount"])
def test_invalid_name_does_not_queue(book, name):
    columns = book.sheets[0].tables[0].columns
    before = len(book.impl.json()["actions"])
    with pytest.raises((TypeError, ValueError)):
        columns.add(name)
    assert len(book.impl.json()["actions"]) == before


def test_remote_ranges_require_async_getters(book):
    column = book.sheets[0].tables[0].columns[0]
    with pytest.raises(NotImplementedError, match="get_range"):
        _ = column.range
    with pytest.raises(NotImplementedError, match="get_data_body_range"):
        _ = column.data_body_range


@pytest.mark.parametrize("scope", ["book", "sheet"])
def test_new_remote_table_requires_column_metadata_refresh(book, monkeypatch, scope):
    sheet = book.sheets[0]
    table = sheet.tables.add(sheet["A1:B2"], name="Fresh")
    with pytest.raises(xw.XlwingsError, match="flush.*load"):
        len(table.columns)

    payload = copy.deepcopy(book.impl.api)
    payload["sheets"][0]["tables"][-1]["columns"] = ["Item", "Amount"]
    payload["sheets"][0]["tables"][-1].pop("_columns_pending")
    monkeypatch.setattr(sys, "platform", "emscripten")
    js = ModuleType("js")
    js.Object = SimpleNamespace(fromEntries=lambda entries: dict(entries))
    js.xlwings = SimpleNamespace(
        getBookData=lambda options: _async_result(
            SimpleNamespace(to_py=lambda: payload)
        )
    )
    monkeypatch.setitem(sys.modules, "js", js)
    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.to_js = lambda value, **kwargs: value
    pyodide.ffi = ffi
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)

    asyncio.run((book if scope == "book" else sheet).load(values=False))
    assert [column.name for column in table.columns] == ["Item", "Amount"]
    assert "_columns_pending" not in table.impl.api


async def _async_result(value):
    return value


def test_column_get_count_refreshes_drifted_metadata(book, monkeypatch):
    table = book.sheets[0].tables["Sales"]
    columns = table.columns
    monkeypatch.setattr(sys, "platform", "emscripten")
    js = ModuleType("js")
    js.xlwings = SimpleNamespace(getTableColumnCount=lambda *args: _async_result(3))
    monkeypatch.setitem(sys.modules, "js", js)
    refreshes = []

    async def refresh(values=None):
        refreshes.append(values)
        table.impl.api["columns"] = ["Item", "Amount", "Tax"]
        table.impl.api["column_count"] = 3

    monkeypatch.setattr(book.impl, "load", refresh)
    assert asyncio.run(columns.get_count()) == 3
    assert refreshes == [False]
    assert len(columns) == 3
    assert [column.name for column in columns] == ["Item", "Amount", "Tax"]

    columns.add("Discount")
    assert asyncio.run(columns.get_count()) == 3
    assert refreshes == [False]
    assert len(columns) == 4


def test_remote_table_rejects_header_guess_before_queue(book):
    sheet = book.sheets[0]
    with pytest.raises(TypeError, match="has_headers must be True or False"):
        sheet.tables.add(sheet["A1:B2"], name="Guess", has_headers="guess")
    assert book.impl.json()["actions"] == []


def test_remote_table_accepts_false_header_option(book):
    sheet = book.sheets[0]
    table = sheet.tables.add(sheet["A1:B2"], name="NoHeaders", has_headers=False)
    assert table.show_headers is True
    assert book.impl.json()["actions"][-1]["args"][1] is False


def test_failed_flush_refreshes_optimistic_column_metadata(book, monkeypatch):
    table = book.sheets[0].tables["Sales"]
    table.columns.add("Tax")
    assert len(table.columns) == 3
    monkeypatch.setattr(sys, "platform", "emscripten")
    js = ModuleType("js")
    js.Object = SimpleNamespace(fromEntries=lambda entries: dict(entries))

    async def fail_actions(payload):
        raise RuntimeError("partial dispatch")

    js.xlwings = SimpleNamespace(runActions=fail_actions)
    monkeypatch.setitem(sys.modules, "js", js)
    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.to_js = lambda value, **kwargs: value
    pyodide.ffi = ffi
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    refreshes = []

    async def refresh(values=None):
        refreshes.append(values)
        table.impl.api["columns"] = ["Item", "Amount"]
        table.impl.api["column_count"] = 2

    monkeypatch.setattr(book.impl, "load", refresh)
    with pytest.raises(RuntimeError, match="partial dispatch"):
        asyncio.run(book.flush())
    assert refreshes == [False]
    assert len(table.columns) == 2
    assert book.impl.json()["actions"] == []
