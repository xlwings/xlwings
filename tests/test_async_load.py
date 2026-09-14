"""
Tests for the async (lazy) book load semantics in the remote backend:

* an async book is marked lazy (``Book._lazy``); sync ``.value`` reads raise
  until values are loaded
* ``Book.load()`` / ``Sheet.load()`` default to metadata-only on a lazy book and
  to a full (values) load on a regular book; ``values=True`` forces a value load
* a metadata-only load doesn't clobber values already loaded

``js`` / ``pyodide`` aren't available in the test environment and the load
methods gate on ``sys.platform == "emscripten"``, so the fixtures below fake just
enough of that surface to run the real branching logic.
"""

import copy
import sys
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro import _xlremote as R


def _book_json(values=None):
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "S",
                "values": values if values is not None else [[]],
                "pictures": [],
                "tables": [],
            }
        ],
    }


def _make_book(lazy, values=None):
    """Create an impl Book through a real *remote* App.

    Building via R.App(...).books.open(...) makes `book.app.engine` the remote
    engine, so reads go through the remote backend rather than launching the
    native Excel engine (which xw.App(impl=None) would do if `book.app` were
    None).
    """
    app = R.App(R.Apps(), add_book=False)
    return app.books.open(_book_json(values), lazy=lazy)


# --- raw_value guard (no js needed) ---


def test_lazy_book_value_read_raises():
    book = _make_book(lazy=True)
    rng = R.Range(sheet=book.sheets.active, arg1=(1, 1))
    with pytest.raises(XlwingsError, match="haven't been loaded"):
        rng.raw_value


def test_eager_book_value_read_works():
    book = _make_book(lazy=False, values=[[1, 2], [3, 4]])
    rng = R.Range(sheet=book.sheets.active, arg1=(1, 1))
    assert rng.raw_value == [(1,)]


def test_lazy_book_value_read_works_after_sheet_marked_loaded():
    book = _make_book(lazy=True, values=[[1, 2], [3, 4]])
    sheet = book.sheets.active
    rng = R.Range(sheet=sheet, arg1=(1, 1))
    with pytest.raises(XlwingsError):
        rng.raw_value
    R._mark_sheet_values_loaded(sheet.api)
    assert rng.raw_value == [(1,)]


def test_loaded_state_survives_sheet_rename():
    # Regression: loaded state must not be keyed on the mutable sheet name.
    book = _make_book(lazy=True, values=[[1, 2], [3, 4]])
    sheet = book.sheets.active
    R._mark_sheet_values_loaded(sheet.api)
    rng = R.Range(sheet=sheet, arg1=(1, 1))
    assert rng.raw_value == [(1,)]
    sheet.api["name"] = "Renamed"  # rename without emitting a JS action
    assert rng.raw_value == [(1,)]  # still loaded, must not raise


# --- load() branching (js/emscripten faked) ---


@pytest.fixture
def fake_emscripten(monkeypatch):
    """Fake sys.platform + js/pyodide so load() runs its real branching logic.

    getBookData records the options it was called with and returns book data
    whose sheet values are non-empty only when NOT lazy (mirroring the server).
    """
    monkeypatch.setattr(sys, "platform", "emscripten")

    calls = []

    async def get_book_data(opts=None):
        opts = dict(opts) if opts else {}
        calls.append(opts)
        lazy = opts.get("lazy", False)
        values = [[]] if lazy else [[10, 20], [30, 40]]
        return SimpleNamespace(to_py=lambda: _book_json(values))

    js = ModuleType("js")
    js.xlwings = SimpleNamespace(getBookData=get_book_data)
    # to_js(..., dict_converter=js.Object.fromEntries) references this; our fake
    # to_js ignores it, but the attribute must exist to be passed as an arg.
    js.Object = SimpleNamespace(fromEntries=lambda x: x)

    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    # to_js just needs to pass the plain dict through for our fake getBookData.
    ffi.to_js = lambda obj, dict_converter=None: obj
    pyodide.ffi = ffi

    monkeypatch.setitem(sys.modules, "js", js)
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    return calls


@pytest.fixture
def anyio_backend():
    return "asyncio"


@pytest.mark.anyio
async def test_lazy_book_load_defaults_to_metadata_only(fake_emscripten):
    calls = fake_emscripten
    book = _make_book(lazy=True)
    await book.load()
    # Requested a lazy (metadata-only) fetch, and no sheet is marked as loaded.
    assert calls[-1].get("lazy") is True
    assert not R._sheet_values_loaded(book.sheets.active.api)
    # Sync reads still raise.
    rng = R.Range(sheet=book.sheets.active, arg1=(1, 1))
    with pytest.raises(XlwingsError):
        rng.raw_value


@pytest.mark.anyio
async def test_lazy_book_load_values_true_loads_values(fake_emscripten):
    calls = fake_emscripten
    book = _make_book(lazy=True)
    await book.load(values=True)
    assert calls[-1].get("lazy") is False
    assert R._sheet_values_loaded(book.sheets.active.api)
    rng = R.Range(sheet=book.sheets.active, arg1=(1, 1))
    assert rng.raw_value == [(10,)]


@pytest.mark.anyio
async def test_eager_book_load_defaults_to_full(fake_emscripten):
    calls = fake_emscripten
    book = _make_book(lazy=False)
    await book.load()
    assert calls[-1].get("lazy") is False
    assert R._sheet_values_loaded(book.sheets.active.api)


@pytest.mark.anyio
async def test_metadata_only_load_does_not_clobber_loaded_values(fake_emscripten):
    book = _make_book(lazy=True)
    await book.load(values=True)  # values now present
    rng = R.Range(sheet=book.sheets.active, arg1=(1, 1))
    assert rng.raw_value == [(10,)]
    await book.load()  # metadata-only refresh must not wipe values
    assert rng.raw_value == [(10,)]


@pytest.mark.anyio
async def test_sheet_load_values_true_marks_only_that_sheet(fake_emscripten):
    book = _make_book(lazy=True)
    sheet = book.sheets.active
    await sheet.load(values=True)
    assert R._sheet_values_loaded(sheet.api)


# --- through the public xw.Book / xw.Sheet wrappers ---


def test_public_wrapper_lazy_value_read_raises():
    book = xw.Book(impl=_make_book(lazy=True))
    with pytest.raises(XlwingsError, match="haven't been loaded"):
        book.sheets[0]["A1"].value


@pytest.mark.anyio
async def test_public_wrapper_forwards_values_kwarg(fake_emscripten):
    book = xw.Book(impl=_make_book(lazy=True))
    await book.load(values=True)
    # Values are now readable through the public wrapper.
    assert book.sheets[0]["A1"].value is not None


def _pivot_metadata(name="Pivot", id="pivot-1"):
    return {
        "id": id,
        "name": name,
        "field_names": ["Region", "Sales"],
        "rows": ["Region"],
        "columns": [],
        "filters": [],
        "values": [
            {
                "id": "value-1",
                "name": "Sum of Sales",
                "source_field": "Sales",
                "function": "Sum",
                "number_format": "General",
            },
            {
                "id": "value-2",
                "name": "Count of Sales",
                "source_field": "Sales",
                "function": "Count",
                "number_format": "General",
            },
        ],
        "layout": "Compact",
        "show_row_grand_totals": True,
        "show_column_grand_totals": True,
    }


@pytest.mark.anyio
@pytest.mark.parametrize("scope", ["book", "sheet"])
async def test_pivot_aliases_survive_metadata_reload(
    fake_emscripten, monkeypatch, scope
):
    payload = _book_json()
    payload["sheets"][0]["pivot_tables"] = [_pivot_metadata()]
    app = R.App(R.Apps(), add_book=False)
    book = xw.Book(impl=app.books.open(copy.deepcopy(payload)))
    try:
        sheet = book.sheets[0]
        collection = sheet.pivot_tables
        pt = collection[0]
        values = pt.values
        first, removed = values[0], values[1]

        async def get_book_data(opts=None):
            return SimpleNamespace(to_py=lambda: copy.deepcopy(payload))

        monkeypatch.setattr(sys.modules["js"].xlwings, "getBookData", get_book_data)
        reload = book.load if scope == "book" else sheet.load
        await reload()
        pt.refresh()
        first.number_format = "0.00"
        assert values[0].number_format == "0.00"

        # IDs preserve wrappers across external renames and reordering. An
        # object with a reused caption but a different ID is a replacement.
        incoming = payload["sheets"][0]["pivot_tables"][0]
        incoming["name"] = "Renamed"
        incoming["values"][0]["name"] = "Revenue"
        incoming["values"][1]["id"] = "replacement-value"
        incoming["values"].reverse()
        payload["sheets"][0]["pivot_tables"].insert(
            0, _pivot_metadata("Other", "pivot-2")
        )
        await reload()
        assert pt.name == "Renamed"
        assert first.name == "Revenue"
        assert collection[1].impl.api is pt.impl.api
        assert values[1].impl.api is first.impl.api
        first.number_format = "0.000"
        assert book.json()["actions"][-1]["args"] == [1, 1, "number_format", "0.000"]
        with pytest.raises(KeyError, match="removed"):
            removed.remove()
        payload["sheets"][0]["pivot_tables"].pop()
        await reload()
        with pytest.raises(KeyError, match="deleted"):
            pt.refresh()
    finally:
        book.close()


@pytest.mark.anyio
async def test_new_unnamed_values_survive_first_reload(fake_emscripten, monkeypatch):
    payload = _book_json()
    payload["sheets"][0]["pivot_tables"] = []
    app = R.App(R.Apps(), add_book=False)
    book = xw.Book(impl=app.books.open(copy.deepcopy(payload)))
    try:
        sheet = book.sheets[0]
        pt = sheet.pivot_tables.add(sheet["A1:B3"], sheet["D1"])
        first = pt.values.add("Sales", function="sum")
        second = pt.values.add("Sales", function="count")
        # Mimic Excel's metadata after this script's queued actions were flushed.
        payload["sheets"][0]["pivot_tables"] = [_pivot_metadata(pt.name)]

        async def get_book_data(opts=None):
            return SimpleNamespace(to_py=lambda: copy.deepcopy(payload))

        monkeypatch.setattr(sys.modules["js"].xlwings, "getBookData", get_book_data)
        await book.load()
        assert first.name == "Sum of Sales"
        assert second.name == "Count of Sales"
        first.remove()
        second.number_format = "0.0"
        assert book.json()["actions"][-1]["args"] == [0, 0, "number_format", "0.0"]
        with pytest.raises(KeyError):
            first.remove()
    finally:
        book.close()
