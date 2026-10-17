"""
Tests that async reads on the remote (Office.js) backend see the script's own
queued writes.

Writes are queued until the script returns, while reads go straight to Excel, so
every client read flushes pending actions first. The Pyodide bridge modules
aren't available in the test environment, so the fixture below fakes the client
and records the order in which actions are dispatched and reads are made.
"""

import asyncio
import sys
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings.pro import _xlremote as R


def _book_json():
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "S1",
                "values": [[None] * 5] * 5,
                "pictures": [],
                "tables": [],
            }
        ],
    }


@pytest.fixture
def book():
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    return xw.Book(impl=impl)


@pytest.fixture
def client(monkeypatch):
    """Fake the client and log its calls in order.

    Each entry is `("dispatch", [action funcs])` or `("read", func, sheet_name)`.
    Set `client.fail_dispatch` to make runActions raise.
    """
    monkeypatch.setattr(sys, "platform", "emscripten")
    log = []
    state = SimpleNamespace(log=log, fail_dispatch=False)

    async def run_actions(payload):
        log.append(("dispatch", [action["func"] for action in payload["actions"]]))
        if state.fail_dispatch:
            raise RuntimeError("dispatch failed")

    async def get_range_data(sheet_name, address, keys):
        log.append(("read", "getRangeData", sheet_name))
        return SimpleNamespace(to_py=lambda: {keys[0]: [[""]]})

    async def get_range_values(sheet_name, address):
        log.append(("read", "getRangeValues", sheet_name))
        return SimpleNamespace(to_py=lambda: [["x"]])

    async def get_book_data(opts):
        log.append(("read", "getBookData", None))
        return SimpleNamespace(to_py=lambda: _book_json())

    js = ModuleType("js")
    js.Object = SimpleNamespace(fromEntries=lambda entries: dict(entries))
    js.xlwings = SimpleNamespace(
        runActions=run_actions,
        getRangeData=get_range_data,
        getRangeValues=get_range_values,
        getBookData=get_book_data,
    )
    monkeypatch.setitem(sys.modules, "js", js)

    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.to_js = lambda value, **kwargs: value
    pyodide.ffi = ffi
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    return state


@pytest.fixture
def anyio_backend():
    return "asyncio"


def pending(book):
    return book.impl.json()["actions"]


@pytest.mark.anyio
async def test_read_flushes_queued_writes_first(book, client):
    rng = book.sheets[0]["A1"]
    rng.value = "Hello"
    await rng.get_formula()
    assert client.log == [
        ("dispatch", ["setValues"]),
        ("read", "getRangeData", "S1"),
    ]
    assert pending(book) == []


@pytest.mark.anyio
async def test_get_value_flushes_queued_writes_first(book, client):
    rng = book.sheets[0]["A1"]
    rng.value = "Hello"
    await rng.get_value()
    assert client.log == [
        ("dispatch", ["setValues"]),
        ("read", "getRangeValues", "S1"),
    ]


@pytest.mark.anyio
async def test_read_without_queued_actions_does_not_flush(book, client):
    await book.sheets[0]["A1"].get_formula()
    assert client.log == [("read", "getRangeData", "S1")]


@pytest.mark.anyio
async def test_consecutive_reads_flush_once(book, client):
    sheet = book.sheets[0]
    sheet["A1"].value = 1
    sheet["A2"].value = 2
    await sheet["A1"].get_formula()
    await sheet["A2"].get_value()
    assert client.log == [
        ("dispatch", ["setValues", "setValues"]),
        ("read", "getRangeData", "S1"),
        ("read", "getRangeValues", "S1"),
    ]


@pytest.mark.anyio
async def test_read_on_a_new_sheet_dispatches_its_creation_first(book, client):
    sheet = book.sheets.add("New")
    sheet["A1:B2"].formula = [["=1", None], [None, "=2"]]
    await sheet["A1:B2"].options(ndim=2).get_formula()
    assert client.log == [
        ("dispatch", ["addSheet", "setFormula"]),
        ("read", "getRangeData", "New"),
    ]


@pytest.mark.anyio
async def test_failed_flush_reloads_metadata_without_flushing_again(book, client):
    client.fail_dispatch = True
    rng = book.sheets[0]["A1"]
    rng.value = "Hello"
    with pytest.raises(RuntimeError, match="dispatch failed"):
        await rng.get_formula()
    # The metadata reload after the failed dispatch reads directly, and the
    # failed read never reaches the client.
    assert client.log == [
        ("dispatch", ["setValues"]),
        ("read", "getBookData", None),
    ]
    assert pending(book) == []


def _delay_dispatch(client):
    """Make runActions take a while and log when it's done."""
    js = sys.modules["js"]
    run_actions = js.xlwings.runActions

    async def slow_run_actions(payload):
        await run_actions(payload)
        await asyncio.sleep(0.01)
        client.log.append(("dispatched",))

    js.xlwings.runActions = slow_run_actions


@pytest.mark.anyio
async def test_overlapping_reads_wait_for_the_dispatch_in_flight(book, client):
    _delay_dispatch(client)
    sheet = book.sheets.add("New")
    sheet["A1"].value = 1
    await asyncio.gather(sheet["A1"].get_value(), sheet["A1"].get_value())
    assert client.log == [
        ("dispatch", ["addSheet", "setValues"]),
        ("dispatched",),
        ("read", "getRangeValues", "New"),
        ("read", "getRangeValues", "New"),
    ]


@pytest.mark.anyio
async def test_batch_queued_during_a_dispatch_is_sent_after_it(book, client):
    _delay_dispatch(client)
    sheet = book.sheets[0]

    async def write_then_read(value):
        sheet["A1"].value = value
        await sheet["A1"].get_value()

    first = asyncio.ensure_future(write_then_read(1))
    await asyncio.sleep(0)  # first batch is now in flight
    await asyncio.gather(first, write_then_read(2))
    # The second batch waits for the first instead of overlapping it, and each
    # read comes after its own batch.
    assert [entry for entry in client.log if entry[0] != "read"] == [
        ("dispatch", ["setValues"]),
        ("dispatched",),
        ("dispatch", ["setValues"]),
        ("dispatched",),
    ]
    reads = [i for i, entry in enumerate(client.log) if entry[0] == "read"]
    done = [i for i, entry in enumerate(client.log) if entry == ("dispatched",)]
    assert len(reads) == 2
    assert reads[0] > done[0]
    assert reads[1] > done[1]
