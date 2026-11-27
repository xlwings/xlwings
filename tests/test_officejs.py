"""
Tests for the Office.js session and proxy recording (`book.officejs()`) with a
fake transport. The JS executor is tested in the xlwings Lite repo.
"""

import json
import sys
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings.pro import _xlremote as R
from xlwings.pro._officejs import (
    OfficeJsError,
    OfficeJsSession,
    OfficeJsWriteWithheld,
)


def _book_json():
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "S1",
                "values": [[None] * 3] * 3,
                "pictures": [],
                "tables": [],
            }
        ],
    }


@pytest.fixture
def anyio_backend():
    return "asyncio"


@pytest.fixture
def client(monkeypatch):
    """Fake `js.xlwings`: logs action dispatches and Office.js requests in order.

    `client.responses` maps a request type to a function returning the response
    fields for a request.
    """
    monkeypatch.setattr(sys, "platform", "emscripten")
    state = SimpleNamespace(log=[], requests=[], responses={})

    async def run_actions(payload):
        state.log.append(("dispatch", [a["func"] for a in payload["actions"]]))

    async def officejs(text):
        request = json.loads(text)
        state.requests.append(request)
        state.log.append(("officejs", request["type"]))
        handler = state.responses.get(request["type"])
        fields = handler(request) if handler else {}
        return json.dumps({"requestId": request["requestId"], "ok": True, **fields})

    js = ModuleType("js")
    js.Object = SimpleNamespace(fromEntries=lambda entries: dict(entries))
    js.xlwings = SimpleNamespace(runActions=run_actions, officejs=officejs)
    monkeypatch.setitem(sys.modules, "js", js)
    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.to_js = lambda value, **kwargs: value
    pyodide.ffi = ffi
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    return state


@pytest.fixture
def book(client):
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    return xw.Book(impl=impl)


def ops(request):
    return [
        (op["kind"], op["target"], op["name"], op.get("args", op.get("value")))
        for op in request["ops"]
    ]


@pytest.mark.anyio
async def test_flushes_pending_actions_on_entry(book, client):
    book.sheets[0]["A1"].value = "x"
    async with book.officejs():
        pass
    assert client.log == [
        ("dispatch", ["setValues"]),
        ("officejs", "officejs.open"),
        ("officejs", "officejs.close"),
    ]


@pytest.mark.anyio
async def test_lazy_recording_and_memoized_identity(book, client):
    async with book.officejs() as ctx:
        rng = ctx.workbook.worksheets.getItem("S1").getRange("A1")
        assert rng.format is rng.format
        _unused = ctx.workbook.names  # never reaches JS
        rng.format.fill.color = "yellow"
        rng.load("values")
        await ctx.sync()
        sync = client.requests[-1]
        assert ops(sync) == [
            ("get", 0, "workbook", None),
            ("get", 1, "worksheets", None),
            ("call", 2, "getItem", ["S1"]),
            ("call", 3, "getRange", ["A1"]),
            ("get", 4, "format", None),
            ("get", 5, "fill", None),
            ("set", 6, "color", "yellow"),
            ("call", 4, "load", ["values"]),
        ]
        assert sync["return"] == [4]


@pytest.mark.anyio
async def test_loaded_data_nested_and_unloaded(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {"2": {"address": "S1!A1", "format": {"fill": {"color": "#FF0"}}}}
    }
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        rng.load("address,format/fill/color")
        await ctx.sync()
        assert rng.address == "S1!A1"
        assert rng.format.fill.color == "#FF0"
        with pytest.raises(OfficeJsError, match="isn't loaded"):
            str(rng.values)


@pytest.mark.anyio
async def test_collection_items_bound_to_handles(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {
            "2": {"items": [{"name": "S1"}, {"name": "S2"}], "$items": ["2:0", "2:1"]}
        }
    }
    async with book.officejs() as ctx:
        sheets = ctx.workbook.worksheets
        sheets.load("items/name")
        await ctx.sync()
        assert [s.name for s in sheets.items] == ["S1", "S2"]
        sheets.items[1].delete()
        client.responses.pop("officejs.sync")
        await ctx.sync()
        assert ops(client.requests[-1]) == [("call", "2:1", "delete", [])]


@pytest.mark.anyio
async def test_client_result_value_and_proxy_args(book, client):
    async with book.officejs() as ctx:
        ws = ctx.workbook.worksheets
        count = ws.getCount()
        client.responses["officejs.sync"] = lambda r: {
            "results": {str(count._handle): {"value": 3}}
        }
        await ctx.sync()
        assert count.value == 3
        ws.getItem("S1").copy("After", ws.getItem("S1"))
        client.responses.pop("officejs.sync")
        await ctx.sync()
        assert ops(client.requests[-1])[-1][3] == ["After", {"$h": 5}]


@pytest.mark.anyio
async def test_invalid_values_raise_before_sending(book, client):
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        with pytest.raises(ValueError, match="NaN"):
            rng.values = [[float("nan")]]
        with pytest.raises(TypeError):
            rng.values = object()


@pytest.mark.anyio
async def test_failed_sync_discards_on_exit(book, client):
    def fail(request):
        return {
            "ok": False,
            "error": {
                "code": "InvalidArgument",
                "message": "bad",
                "opId": request["ops"][-1]["id"],
            },
        }

    client.responses["officejs.sync"] = fail
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().values = [[1]]
        with pytest.raises(
            OfficeJsError, match=r"InvalidArgument: bad \(at .*values\)"
        ):
            await ctx.sync()
        with pytest.raises(OfficeJsError, match="failed earlier"):
            await ctx.sync()
    assert client.requests[-1]["type"] == "officejs.close"
    assert client.requests[-1]["discard"] is True


@pytest.mark.anyio
async def test_withheld_write_ends_session_with_dedicated_exception(book, client):
    def withhold(request):
        return {"ok": False, "error": {"code": "withheld", "message": "Dry run"}}

    client.responses["officejs.sync"] = withhold
    with pytest.raises(OfficeJsWriteWithheld, match="withheld: Dry run"):
        async with book.officejs() as ctx:
            ctx.workbook.getSelectedRange().values = [[1]]
            await ctx.sync()
    assert client.requests[-1]["type"] == "officejs.close"
    assert client.requests[-1]["discard"] is True

    # A write left for close is withheld there; the add-in discards it.
    client.responses["officejs.close"] = withhold
    with pytest.raises(OfficeJsWriteWithheld):
        async with book.officejs() as ctx:
            ctx.workbook.getSelectedRange().values = [[1]]
    assert book.impl._officejs_session is None


@pytest.mark.anyio
async def test_exception_discards_and_unsynced_ops_go_with_close(book, client):
    with pytest.raises(RuntimeError):
        async with book.officejs() as ctx:
            ctx.workbook.getSelectedRange().values = [[1]]
            raise RuntimeError
    assert client.requests[-1] == {
        **client.requests[-1],
        "type": "officejs.close",
        "discard": True,
    }
    assert "ops" not in client.requests[-1]

    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().values = [[1]]
    assert client.requests[-1]["type"] == "officejs.close"
    assert len(client.requests[-1]["ops"]) == 3


@pytest.mark.anyio
async def test_loaded_data_readable_after_exit_but_nothing_new(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {
            "2": {"address": "S1!A1", "format": {"fill": {"color": "#FF0"}}},
            "4": {"items": [{"name": "S1"}], "$items": ["4:0"]},
        }
    }
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        rng.load("address,format/fill/color")
        sheets = ctx.workbook.worksheets
        sheets.load("items/name")
        await ctx.sync()
    assert rng.address == "S1!A1"
    assert rng.format.fill.color == "#FF0"
    assert [s.name for s in sheets.items] == ["S1"]
    for use in (
        lambda: rng.values,
        lambda: rng.format.font,
        lambda: setattr(rng, "values", [[1]]),
        lambda: rng.getCell(0, 0),
        lambda: rng.load("values"),
        lambda: sheets.items[0].delete(),
    ):
        with pytest.raises(xw.XlwingsError, match="closed session"):
            use()
    with pytest.raises(xw.XlwingsError, match="closed session"):
        await ctx.sync()


@pytest.mark.anyio
async def test_plain_object_values_are_dicts(book, client):
    rule = {"formula1": "=10", "formula2": None, "operator": "GreaterThan"}
    client.responses["officejs.sync"] = lambda r: {
        "results": {
            "4": {
                "type": "CellValue",
                "cellValue": {"rule": rule, "format": {"fill": {"color": "#F00"}}},
                "$plain": ["cellValue/rule"],
            },
            "3": {
                "items": [{"cellValue": {"rule": rule}, "$plain": ["cellValue/rule"]}],
                "$items": ["3:0"],
            },
        }
    }
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        cf = rng.conditionalFormats.add("CellValue")
        cf.load("type,cellValue")
        formats = rng.conditionalFormats
        formats.load("items/cellValue")
        await ctx.sync()
        assert cf.cellValue.rule == rule
        assert cf.cellValue.rule["operator"] == "GreaterThan"
        # Navigation properties stay proxies, even inside a plain-object parent
        assert cf.cellValue.format.fill.color == "#F00"
        assert formats.items[0].cellValue.rule == rule
        assert "$plain" not in formats.items[0]._data


@pytest.mark.anyio
async def test_curated_calls_blocked_inside_session(book, client, reloads):
    sheet = book.sheets[0]
    async with book.officejs():
        with pytest.raises(xw.XlwingsError, match="Office.js session is open"):
            sheet["A1"].value = 1
        with pytest.raises(xw.XlwingsError, match="Office.js session is open"):
            await sheet["A1"].get_value()
        # Snapshot access keeps working
        assert sheet.name == "S1"
        assert sheet["A1"].address == "$A$1"
    sheet["A1"].value = 1
    assert book.impl.json()["actions"][-1]["func"] == "setValues"


@pytest.mark.anyio
async def test_api_inside_and_outside_session(book, client):
    sheet = book.sheets[0]
    with pytest.raises(xw.XlwingsError, match="book.officejs"):
        book.api
    with pytest.raises(xw.XlwingsError, match="book.officejs"):
        sheet["A1"].api
    book.impl.api["sheets"][0]["id"] = "{S1-ID}"
    async with book.officejs() as ctx:
        assert book.api is ctx.workbook
        assert sheet.api is book.sheets["S1"].api
        rng = sheet["A1:B2"].api
        assert rng is sheet["A1:B2"].api
        rng.format.fill.color = "yellow"
        await ctx.sync()
        assert ops(client.requests[-1])[:4] == [
            ("get", 0, "workbook", None),
            ("get", 1, "worksheets", None),
            ("call", 2, "getItem", ["{S1-ID}"]),
            ("call", 3, "getRange", ["$A$1:$B$2"]),
        ]
    with pytest.raises(xw.XlwingsError, match="book.officejs"):
        sheet.api


def test_api_unchanged_outside_pyodide(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    book = xw.Book(impl=impl)
    assert book.api is impl.api
    assert book.sheets[0].api is impl.api["sheets"][0]


def _wrote(request):
    return {"wrote": any(op["kind"] == "set" for op in request.get("ops", []))}


@pytest.fixture
def reloads(client, monkeypatch):
    """Count snapshot reloads, serving `sheets` as the reloaded sheet list."""
    state = SimpleNamespace(count=0, sheets=None)

    async def get_book_data(opts):
        state.count += 1
        data = _book_json()
        if state.sheets is not None:
            data["sheets"] = state.sheets
        return SimpleNamespace(to_py=lambda: data)

    sys.modules["js"].xlwings.getBookData = get_book_data
    return state


@pytest.mark.anyio
async def test_reload_only_after_writes(book, client, reloads):
    client.responses["officejs.sync"] = _wrote
    client.responses["officejs.close"] = _wrote
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().load("values")
        await ctx.sync()
    assert reloads.count == 0
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().values = [[1]]
        await ctx.sync()
    assert reloads.count == 1
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().values = [[1]]
    assert reloads.count == 2


@pytest.mark.anyio
async def test_reload_after_failed_write_batch(book, client, reloads):
    client.responses["officejs.sync"] = lambda r: {
        "ok": False,
        "error": {"code": "ItemNotFound", "message": "gone", "wrote": True},
    }
    with pytest.raises(OfficeJsError):
        async with book.officejs() as ctx:
            ctx.workbook.getSelectedRange().values = [[1]]
            await ctx.sync()
    assert reloads.count == 1


@pytest.mark.anyio
async def test_reload_matches_sheets_by_id(book, client, reloads):
    impl = book.impl
    impl.api["sheets"][0]["id"] = "{1}"
    sheet = book.sheets["S1"]
    other = book.sheets.add("S2")
    other.impl.api["id"] = "{2}"
    impl._json["actions"].clear()
    client.responses["officejs.close"] = lambda r: {"wrote": True}
    reloads.sheets = [
        {"name": "Renamed", "id": "{1}", "values": [[1]], "pictures": [], "tables": []}
    ]
    async with book.officejs() as ctx:
        ctx.workbook.worksheets.getItem("{1}").name = "Renamed"
        ctx.workbook.worksheets.getItem("{2}").delete()
    assert sheet.name == "Renamed"
    assert sheet.index == 1
    assert sheet["A1"].value == 1
    with pytest.raises(xw.XlwingsError, match="'S2' was deleted"):
        other["A1"].value = 1


def test_requires_xlwings_lite(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    with pytest.raises(NotImplementedError):
        xw.Book(impl=impl).officejs()
    assert isinstance(OfficeJsSession(impl), OfficeJsSession)


@pytest.mark.anyio
async def test_reload_keeps_copied_sheet_apart_from_its_source(book, client, reloads):
    book.impl.api["sheets"][0]["id"] = "{1}"
    source = book.sheets["S1"]
    copied = source.copy(after=source)
    assert "id" not in copied.impl.api
    client.responses["officejs.close"] = lambda r: {"wrote": True}
    reloads.sheets = [
        {"name": "S1", "id": "{1}", "values": [[1]], "pictures": [], "tables": []},
        {"name": "S1 (2)", "id": "{2}", "values": [[2]], "pictures": [], "tables": []},
    ]
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().values = [[1]]
    assert (source.name, source.index) == ("S1", 1)
    assert (copied.name, copied.index) == ("S1 (2)", 2)
    copied["A1"].value = 3
    assert book.impl.json()["actions"][-1]["sheet_position"] == 1


@pytest.mark.anyio
async def test_blocked_action_leaves_no_phantom_sheet(book, client, reloads):
    sheet = book.sheets["S1"]
    async with book.officejs():
        with pytest.raises(xw.XlwingsError, match="Office.js session is open"):
            book.sheets.add("New", before=sheet)
    # Nothing was written, but the blocked add had changed the snapshot
    assert reloads.count == 1
    assert book.sheet_names == ["S1"]
    sheet["A1"].value = 1
    assert book.impl.json()["actions"][-1]["sheet_position"] == 0


@pytest.mark.anyio
async def test_is_null_object_after_sync(book, client):
    async with book.officejs() as ctx:
        ws = ctx.workbook.worksheets.getItemOrNullObject("Nope")
        with pytest.raises(OfficeJsError, match="available after 'await ctx.sync"):
            bool(ws.isNullObject)
        client.responses["officejs.sync"] = lambda r: {
            "results": {str(ws._handle): {"isNullObject": True}}
        }
        await ctx.sync()
        assert ws.isNullObject is True
