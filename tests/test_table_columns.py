"""Table column API validation and remote action payloads."""

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
