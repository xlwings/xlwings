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
from xlwings.pro._officejs import OfficeJsError, OfficeJsSession


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
async def test_proxies_invalid_after_exit(book, client):
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
    with pytest.raises(xw.XlwingsError, match="closed session"):
        rng.values = [[1]]


def test_requires_xlwings_lite(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    with pytest.raises(NotImplementedError):
        xw.Book(impl=impl).officejs()
    assert isinstance(OfficeJsSession(impl), OfficeJsSession)
