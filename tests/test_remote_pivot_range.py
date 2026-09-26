"""PivotTable ranges are fetched from Excel in xlwings Lite."""

import asyncio
import sys
from types import ModuleType

import pytest

import xlwings as xw
from xlwings.pro import _xlremote as remote


def book():
    payload = {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "Pivot Sheet",
                "values": [[]],
                "pictures": [],
                "tables": [],
                "pivot_tables": [
                    {
                        "id": "pivot-1",
                        "name": "Sales",
                        "field_names": ["Region", "Sales"],
                        "rows": ["Region"],
                        "columns": [],
                        "filters": [],
                        "values": [{"name": "Sum of Sales", "source_field": "Sales"}],
                        "layout": "Compact",
                        "show_row_grand_totals": True,
                        "show_column_grand_totals": True,
                    }
                ],
            }
        ],
    }
    impl = remote.App(remote.Apps(), add_book=False).books.open(payload, lazy=True)
    return xw.Book(impl=impl)


class PivotClient:
    def __init__(self, addresses):
        self.addresses = addresses
        self.calls = []

    async def getPivotTableRangeAddress(self, sheet, index, pivot_id, pivot_name, kind):
        self.calls.append((sheet, index, pivot_id, pivot_name, kind))
        return self.addresses[kind]


def install_js(monkeypatch, addresses):
    client = PivotClient(addresses)
    js = ModuleType("js")
    js.xlwings = client
    monkeypatch.setitem(sys.modules, "js", js)
    monkeypatch.setattr(sys, "platform", "emscripten")
    return client


def test_pivot_range_properties_direct_to_async_getters():
    pivot = book().sheets[0].pivot_tables[0]
    with pytest.raises(NotImplementedError, match="get_range"):
        _ = pivot.range
    with pytest.raises(NotImplementedError, match="get_data_body_range"):
        _ = pivot.data_body_range


def test_pivot_range_reads_require_xlwings_lite():
    pivot = book().sheets[0].pivot_tables[0]
    with pytest.raises(NotImplementedError, match="require xlwings Lite"):
        asyncio.run(pivot.get_range())
    with pytest.raises(NotImplementedError, match="require xlwings Lite"):
        asyncio.run(pivot.get_data_body_range())


def test_pivot_range_reads_current_addresses(monkeypatch):
    client = install_js(monkeypatch, {"report": "A3:C8", "data_body": "B4:C8"})
    pivot = book().sheets[0].pivot_tables[0]

    report = asyncio.run(pivot.get_range())
    body = asyncio.run(pivot.get_data_body_range())

    assert isinstance(report, xw.Range)
    assert report.address == "$A$3:$C$8"
    assert report.sheet.name == "Pivot Sheet"
    assert body.address == "$B$4:$C$8"
    assert client.calls == [
        ("Pivot Sheet", 0, "pivot-1", "Sales", "report"),
        ("Pivot Sheet", 0, "pivot-1", "Sales", "data_body"),
    ]


def test_empty_data_body_returns_none(monkeypatch):
    client = install_js(monkeypatch, {"data_body": None})
    pivot = book().sheets[0].pivot_tables[0]

    assert asyncio.run(pivot.get_data_body_range()) is None
    assert client.calls == [("Pivot Sheet", 0, "pivot-1", "Sales", "data_body")]


def test_missing_report_address_is_an_error(monkeypatch):
    install_js(monkeypatch, {"report": None})
    pivot = book().sheets[0].pivot_tables[0]
    with pytest.raises(RuntimeError, match="no pivot table report range"):
        asyncio.run(pivot.get_range())


def test_deleted_pivot_is_rejected_before_js_call(monkeypatch):
    client = install_js(monkeypatch, {"report": "A1"})
    sheet = book().sheets[0]
    pivot = sheet.pivot_tables[0]
    pivot.delete()
    with pytest.raises(KeyError, match="deleted"):
        asyncio.run(pivot.get_range())
    assert client.calls == []
