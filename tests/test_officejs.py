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


@pytest.mark.anyio
async def test_set_after_load_records_and_keeps_loaded_value(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {"2": {"values": [[1]]}} if r["return"] else {}
    }
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        rng.load("values")
        await ctx.sync()
        rng.values = [[2]]
        await ctx.sync()
        assert ops(client.requests[-1]) == [("set", 2, "values", [[2]])]
        assert client.requests[-1]["return"] == []
        # Loaded values are as of the sync that followed their load()
        assert rng.values == [[1]]
        client.responses["officejs.sync"] = lambda r: {
            "results": {"2": {"values": [[2]]}}
        }
        rng.load("values")
        await ctx.sync()
        assert rng.values == [[2]]


@pytest.mark.anyio
async def test_proxies_have_no_value(book, client):
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        for use in (
            str,
            bool,
            len,
            list,
            int,
            float,
            lambda value: value == 1,
            lambda value: value < 1,
            lambda value: value + 1,
            lambda value: 1 - value,
            lambda value: value[0],
            lambda value: 1 in value,
        ):
            with pytest.raises(
                OfficeJsError,
                match=r'values isn\'t loaded; call .*getSelectedRange\(\)\.load\("values"\)',
            ):
                use(rng.values)
        with pytest.raises(OfficeJsError, match="is an Office.js object, not a value"):
            str(rng)
        assert repr(rng) == "<Proxy context.workbook.getSelectedRange()>"
        assert (
            repr(rng.values)
            == "<Proxy context.workbook.getSelectedRange().values, not loaded>"
        )
        # Proxies can be dict keys and set members
        assert len({rng, rng.format, rng.format}) == 2


@pytest.mark.anyio
async def test_messages_for_method_results_and_loaded_objects(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {"1": {"worksheets": {"items": [{"name": "S1"}]}}}
    }
    async with book.officejs() as ctx:
        count = ctx.workbook.worksheets.getCount()
        # A ClientResult has no load()
        with pytest.raises(OfficeJsError, match=r"value isn't available yet.*sync"):
            int(count.value)
        with pytest.raises(OfficeJsError, match=r"not a value.*\.value for the result"):
            int(count)
        ctx.workbook.load("worksheets")
        await ctx.sync()
        # Loading a navigation property doesn't turn it into a value
        with pytest.raises(OfficeJsError, match=r"not a value.*\.items for a collec"):
            len(ctx.workbook.worksheets)
        assert repr(ctx.workbook.worksheets) == "<Proxy context.workbook.worksheets>"


@pytest.mark.anyio
async def test_session_misuse_raises(book, client):
    unopened = book.officejs()
    with pytest.raises(xw.XlwingsError, match="async with book.officejs"):
        unopened.workbook
    with pytest.raises(xw.XlwingsError, match="async with book.officejs"):
        await unopened.sync()
    assert client.requests == []
    async with book.officejs() as ctx:
        with pytest.raises(xw.XlwingsError, match="already open"):
            async with book.officejs():
                pass
        rng = ctx.workbook.getSelectedRange()
        with pytest.raises(TypeError, match="positional arguments only"):
            rng.getCell(row=0, column=0)
        with pytest.raises(TypeError, match="isn't callable"):
            rng()
        with pytest.raises(TypeError, match="keys .* must be strings"):
            rng.format.autofitColumns({1: "x"})
        with pytest.raises(ValueError, match="inf"):
            rng.getCell(float("inf"), 0)
        # The outer session is unaffected and nothing invalid was recorded
        await ctx.sync()
        assert ops(client.requests[-1]) == [
            ("get", 0, "workbook", None),
            ("call", 1, "getSelectedRange", []),
        ]


@pytest.mark.anyio
async def test_api_falls_back_to_sheet_name_without_id(book, client):
    sheet = book.sheets[0]
    assert "id" not in sheet.impl.api
    async with book.officejs() as ctx:
        assert sheet["A1"].api is sheet["A1"].api
        sheet["A1"].api.clear()
        await ctx.sync()
        assert ops(client.requests[-1])[2:4] == [
            ("call", 2, "getItem", ["S1"]),
            ("call", 3, "getRange", ["$A$1"]),
        ]


def _fail(message):
    return lambda r: {
        "ok": False,
        "error": {"code": "GeneralException", "message": message},
    }


@pytest.mark.anyio
async def test_failing_cleanup_doesnt_mask_the_original_exception(
    book, client, reloads
):
    client.responses["officejs.close"] = _fail("close failed")
    with pytest.raises(RuntimeError, match="original"):
        async with book.officejs():
            raise RuntimeError("original")
    assert book.impl._officejs_session is None
    # Without an exception in the block, a failing close is what the script sees
    with pytest.raises(OfficeJsError, match="close failed"):
        async with book.officejs():
            pass
    assert book.impl._officejs_session is None

    async def broken_reload(opts):
        raise OSError("reload failed")

    sys.modules["js"].xlwings.getBookData = broken_reload
    client.responses.pop("officejs.close")
    client.responses["officejs.sync"] = _wrote
    with pytest.raises(RuntimeError, match="original"):
        async with book.officejs() as ctx:
            ctx.workbook.getSelectedRange().values = [[1]]
            await ctx.sync()
            raise RuntimeError("original")
    with pytest.raises(OSError, match="reload failed"):
        async with book.officejs() as ctx:
            ctx.workbook.getSelectedRange().values = [[1]]
            await ctx.sync()
    assert book.impl._officejs_session is None


@pytest.mark.anyio
async def test_lazy_book_values_need_a_fresh_load_after_a_write(client, reloads):
    data = _book_json()
    data["sheets"][0]["values"] = [[1]]
    impl = R.App(R.Apps(), add_book=False).books.open(data, lazy=True)
    book = xw.Book(impl=impl)
    sheet = book.sheets[0]
    R._mark_sheet_values_loaded(sheet.impl.api)
    assert sheet["A1"].value == 1
    client.responses["officejs.close"] = _wrote
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().load("values")
    assert sheet["A1"].value == 1
    async with book.officejs() as ctx:
        ctx.workbook.getSelectedRange().values = [[2]]
    assert reloads.count == 1
    with pytest.raises(xw.XlwingsError, match="haven't been loaded"):
        sheet["A1"].value


@pytest.mark.anyio
async def test_refused_open_leaves_no_session(book, client):
    client.responses["officejs.open"] = lambda r: {
        "ok": False,
        "error": {"code": "NoActiveRun", "message": "No run"},
    }
    with pytest.raises(OfficeJsError, match="NoActiveRun"):
        async with book.officejs():
            pass
    assert book.impl._officejs_session is None
    assert [r["type"] for r in client.requests] == ["officejs.open"]
    book.sheets[0]["A1"].value = 1


@pytest.mark.anyio
async def test_options_dicts_and_numpy_scalars(book, client):
    np = pytest.importorskip("numpy")
    async with book.officejs() as ctx:
        sheets = ctx.workbook.worksheets
        sheets.getItem("S1").copy(
            {"position": "After", "relativeTo": sheets.getItem("S1")}
        )
        rng = ctx.workbook.getSelectedRange()
        rng.values = [[np.bool_(True), np.int64(1), np.float64(1.5)], (None, "", 2)]
        with pytest.raises(TypeError, match=r"numpy\.ndarray .*\.tolist\(\)"):
            rng.values = np.array([[1]])
        for date in (np.datetime64("2026-11-27", "ns"), np.datetime64("2026-11-27")):
            with pytest.raises(TypeError, match=r"numpy\.datetime64"):
                rng.values = [[date]]
        with pytest.raises(TypeError, match=r"numpy\.timedelta64"):
            rng.values = [[np.timedelta64(1, "D")]]
        with pytest.raises(ValueError, match='use "" instead'):
            rng.values = [[np.float64("nan")]]
        await ctx.sync()
        assert ops(client.requests[-1])[4:] == [
            ("call", 3, "copy", [{"position": "After", "relativeTo": {"$h": 4}}]),
            ("call", 1, "getSelectedRange", []),
            ("set", 6, "values", [[True, 1, 1.5], [None, "", 2]]),
        ]


@pytest.mark.anyio
async def test_api_messages_name_the_xlwings_object(book, client):
    book.impl.api["sheets"][0]["id"] = "{S1-ID}"
    client.responses["officejs.sync"] = lambda r: {
        "ok": False,
        "error": {"code": "InvalidArgument", "message": "bad", "opId": 4},
    }
    async with book.officejs() as ctx:
        rng = book.sheets[0]["A1"].api
        with pytest.raises(
            OfficeJsError,
            match=r'book\.sheets\["S1"\]\["\$A\$1"\]\.api\.load\("values"\)',
        ):
            str(rng.values)
        assert repr(book.sheets[0].api) == '<Proxy book.sheets["S1"].api>'
        with pytest.raises(OfficeJsError) as info:
            await ctx.sync()
        assert info.value.op == 'book.sheets["S1"]["$A$1"].api'
        # The request still addresses the sheet by ID
        assert ops(client.requests[-1])[2][3] == ["{S1-ID}"]


@pytest.mark.anyio
async def test_reloaded_collection_keeps_earlier_items_on_their_handles(book, client):
    loads = iter(
        [
            {"items": [{"name": "S1"}], "$items": ["2:0~1"]},
            {"items": [{"name": "New"}, {"name": "S1"}], "$items": ["2:0~2", "2:1~2"]},
        ]
    )
    client.responses["officejs.sync"] = lambda r: (
        {"results": {"2": next(loads)}} if r["return"] else {}
    )
    async with book.officejs() as ctx:
        sheets = ctx.workbook.worksheets
        sheets.load("items/name")
        await ctx.sync()
        first = sheets.items[0]
        sheets.load("items/name")
        await ctx.sync()
        assert [s.name for s in sheets.items] == ["New", "S1"]
        first.delete()
        await ctx.sync()
        assert first.name == "S1"
        assert ops(client.requests[-1]) == [("call", "2:0~1", "delete", [])]


@pytest.mark.anyio
async def test_latest_load_wins_over_an_earlier_nested_load(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {
            str(h): (
                {"color": "#000"} if h == 4 else {"format": {"fill": {"color": "#FF0"}}}
            )
            for h in r["return"]
        }
    }
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        fill = rng.format.fill
        fill.load("color")
        await ctx.sync()
        assert fill.color == "#000"
        fill.color = "#FF0"
        rng.load("format/fill/color")
        await ctx.sync()
        assert rng.format.fill.color == "#FF0"
        assert fill.color == "#FF0"


@pytest.mark.anyio
async def test_items_of_a_collection_loaded_through_its_parent(book, client):
    # Office.js's toJSON() has a collection loaded through its parent as a list
    client.responses["officejs.sync"] = lambda r: {
        "results": {
            "1": {
                "worksheets": [{"name": "S1"}, {"name": "S2"}],
                "$itemsAt": {"worksheets": ["1.worksheets:0~1", "1.worksheets:1~1"]},
            }
        }
        if r["return"]
        else {}
    }
    async with book.officejs() as ctx:
        ctx.workbook.load("worksheets/items/name")
        await ctx.sync()
        sheets = ctx.workbook.worksheets
        assert [s.name for s in sheets.items] == ["S1", "S2"]
        assert repr(sheets.items[1]) == "<Proxy context.workbook.worksheets.items[1]>"
        with pytest.raises(OfficeJsError, match=r"not a value.*\.items for a collec"):
            len(sheets)
        sheets.items[1].delete()
        await ctx.sync()
        assert ops(client.requests[-1]) == [("call", "1.worksheets:1~1", "delete", [])]
    # Loaded data stays readable after the session
    assert [s.name for s in ctx.workbook.worksheets.items] == ["S1", "S2"]


@pytest.mark.anyio
async def test_context_load_loads_its_argument(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {str(h): {"values": [[1]]} for h in r["return"]}
    }
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        assert ctx.context.load(rng, "values") is rng
        await ctx.sync()
        assert ops(client.requests[-1])[-1] == (
            "call",
            0,
            "load",
            [{"$h": 2}, "values"],
        )
        assert rng.values == [[1]]


@pytest.mark.anyio
async def test_numpy_bool_arrays_need_tolist(book, client):
    np = pytest.importorskip("numpy")
    async with book.officejs() as ctx:
        rng = ctx.workbook.getSelectedRange()
        for array in (np.array([True, False]), np.array([[True]])):
            with pytest.raises(TypeError, match=r"\.tolist\(\)"):
                rng.values = array


@pytest.mark.anyio
async def test_concurrent_sessions_on_one_book_are_refused(book, client):
    import asyncio

    async def open_session():
        async with book.officejs():
            await asyncio.sleep(0.01)

    book.sheets[0]["A1"].value = 1  # entry awaits a flush
    results = await asyncio.gather(
        open_session(), open_session(), return_exceptions=True
    )
    assert results[0] is None
    assert isinstance(results[1], xw.XlwingsError)
    assert "already open" in str(results[1])
    assert [r["type"] for r in client.requests] == ["officejs.open", "officejs.close"]


@pytest.mark.anyio
async def test_failed_entry_releases_the_book(book, client):
    async def broken_run_actions(payload):
        raise RuntimeError("flush failed")

    sys.modules["js"].xlwings.runActions = broken_run_actions
    sys.modules["js"].xlwings.getBookData = None
    book.sheets[0]["A1"].value = 1
    with pytest.raises(RuntimeError, match="flush failed"):
        async with book.officejs():
            pass
    assert book.impl._officejs_session is None


def test_xlwings_server_wasm_keeps_the_data_dict_api(book, client):
    del sys.modules["js"].xlwings.officejs
    sheet = book.sheets[0]
    assert book.api is book.impl.api
    assert sheet.api is book.impl.api["sheets"][0]
    assert sheet["A1"].api == [(None,)]
    with pytest.raises(NotImplementedError, match="xlwings Lite"):
        book.officejs()


def test_macro_arguments_use_the_data_dict(book, client):
    book.macro("callback")(book.sheets[0]["A1"])
    action = book.impl.json()["actions"][-1]
    assert action["func"] == "runMacro"
    assert action["args"][1:] == [[(None,)]]


def test_office_js_errors_are_public():
    from xlwings.pro import _officejs

    assert xw.OfficeJsError is _officejs.OfficeJsError is OfficeJsError
    assert issubclass(xw.OfficeJsWriteWithheld, xw.OfficeJsError)
    assert issubclass(xw.OfficeJsError, xw.XlwingsError)


@pytest.mark.anyio
async def test_session_forwards_office_js_members_to_the_context(book, client):
    client.responses["officejs.sync"] = lambda r: {
        "results": {str(h): {"values": [[1]]} for h in r["return"]}
    }
    async with book.officejs() as ctx:
        assert ctx.workbook is ctx.context.workbook
        assert ctx.application is ctx.context.application
        ctx.application.calculate("Full")
        rng = ctx.workbook.getSelectedRange()
        assert ctx.load(rng, "values") is rng
        # sync() is the session's own, not RequestContext.sync recorded as an op
        await ctx.sync()
        assert ops(client.requests[-1]) == [
            ("get", 0, "application", None),
            ("call", 1, "calculate", ["Full"]),
            ("get", 0, "workbook", None),
            ("call", 3, "getSelectedRange", []),
            ("call", 0, "load", [{"$h": 4}, "values"]),
        ]
        assert rng.values == [[1]]
        with pytest.raises(AttributeError):
            ctx._nope
    with pytest.raises(xw.XlwingsError, match="closed session"):
        ctx.runtime
