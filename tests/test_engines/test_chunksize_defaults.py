"""Headless tests for the automatic ``chunksize`` defaults.

Large range reads/writes are chunked automatically when the user hasn't passed an
explicit ``chunksize``: the engine's ``max_cells_per_read`` / ``max_cells_per_write``
budget applies only when the range is larger than the budget. These tests drive the
real remote engine (``xw.Book(json=...)``) with tiny, test-only budgets so chunk
boundaries can be asserted without allocating millions of cells, and the Calamine
engine for its whole-sheet shortcut. No Excel is needed.

Run with ``pytest tests/test_engines/test_chunksize_defaults.py``.
"""

import asyncio
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings import XlwingsError, base_classes, conversion
from xlwings.conversion import standard
from xlwings.pro import _xlcalamine, _xlremote

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

this_dir = Path(__file__).resolve().parent


# --- helpers / fixtures ---------------------------------------------------------


def _payload(nrows=6, ncols=3):
    values = [[r * ncols + c for c in range(ncols)] for r in range(nrows)]
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [{"name": "S", "values": values, "pictures": [], "tables": []}],
    }


def _values(nrows=6, ncols=3):
    return _payload(nrows, ncols)["sheets"][0]["values"]


@pytest.fixture
def book():
    book = xw.Book(json=_payload())
    yield book
    book.close()


@pytest.fixture
def sheet(book):
    return book.sheets[0]


def _actions(book):
    return [
        {
            k: v
            for k, v in a.items()
            if k in ("func", "start_row", "start_column", "row_count", "column_count")
        }
        | {"values": a.get("values")}
        for a in book.json()["actions"]
    ]


def _set_values_actions(book):
    return [a for a in _actions(book) if a["func"] == "setValues"]


@pytest.fixture
def read_budget(monkeypatch):
    def apply(budget):
        monkeypatch.setattr(
            _xlremote.Range, "max_cells_per_read", property(lambda self: budget)
        )

    return apply


@pytest.fixture
def write_budget(monkeypatch):
    def apply(budget):
        monkeypatch.setattr(
            _xlremote.Range, "max_cells_per_write", property(lambda self: budget)
        )

    return apply


@pytest.fixture
def raw_reads(monkeypatch):
    """Spy on the remote engine's raw reads: records (address, options) per read."""
    calls = []
    original = _xlremote.Range.raw_value

    def getter(self):
        calls.append((self.address, dict(self.options)))
        return original.fget(self)

    monkeypatch.setattr(_xlremote.Range, "raw_value", property(getter, original.fset))
    return calls


@pytest.fixture
def com_style_reads(monkeypatch):
    """Like raw_reads, but a single cell comes back as a scalar (Windows/macOS)."""
    calls = []
    original = _xlremote.Range.raw_value

    def getter(self):
        calls.append(self.address)
        value = original.fget(self)
        if self.shape == (1, 1):
            return value[0][0]
        return value

    monkeypatch.setattr(_xlremote.Range, "raw_value", property(getter, original.fset))
    return calls


class _FakeJsNull:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __bool__(self):
        return False


@pytest.fixture
def fake_pyodide(monkeypatch):
    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.JsNull = _FakeJsNull
    ffi.jsnull = _FakeJsNull()
    pyodide.ffi = ffi
    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    return ffi.jsnull


class _FakeJsProxy:
    def __init__(self, data):
        self._data = data

    def to_py(self):
        return self._data


@pytest.fixture
def live_js(monkeypatch, book):
    """Fake ``js.xlwings`` for the async (xlwings Lite) read path.

    ``getRangeValues`` serves the book's own values and records the addresses
    requested. ``responses`` lets a test override what a given call returns (or
    raise): a value, ``None``, a JsNull, or an exception instance.
    """
    state = SimpleNamespace(addresses=[], responses={})

    async def get_range_values(sheet_name, address):
        state.addresses.append(address)
        ix = len(state.addresses) - 1
        if ix in state.responses:
            response = state.responses[ix]
            if isinstance(response, Exception):
                raise response
            return response
        raw = book.sheets[sheet_name].range(address).impl.raw_value
        return _FakeJsProxy([list(row) for row in raw])

    async def get_expanded_address(sheet_name, address, mode):
        return "A1:C6"

    js = ModuleType("js")
    js.xlwings = ModuleType("js.xlwings")
    js.xlwings.getRangeValues = get_range_values
    js.xlwings.getExpandedAddress = get_expanded_address
    monkeypatch.setitem(sys.modules, "js", js)
    monkeypatch.setattr(sys, "platform", "emscripten")
    return state


def _run(coro):
    return asyncio.run(coro)


# --- policy -----------------------------------------------------------------------


def test_default_budgets_are_independent():
    assert base_classes.DEFAULT_MAX_CELLS_PER_READ == 4_000_000
    assert base_classes.DEFAULT_MAX_CELLS_PER_WRITE == 100_000
    rng = base_classes.Range()
    assert rng.max_cells_per_read == 4_000_000
    assert rng.max_cells_per_write == 100_000
    # Remote inherits both, Calamine opts out of implicit read chunking only
    assert _xlremote.Range.max_cells_per_read is base_classes.Range.max_cells_per_read
    assert _xlremote.Range.max_cells_per_write is base_classes.Range.max_cells_per_write
    assert _xlcalamine.Range.__dict__["max_cells_per_read"].fget(None) is None


def test_budget_chunksize_arithmetic():
    assert standard._budget_chunksize(4_000_000, 25_000, 200) == 20_000
    assert standard._budget_chunksize(4_000_000, 20_000, 200) is None  # exactly
    assert standard._budget_chunksize(4_000_000, 19_999, 200) is None
    assert standard._budget_chunksize(100_000, 1_001, 200) == 500
    assert standard._budget_chunksize(2, 3, 5) == 1  # never 0 rows
    assert standard._budget_chunksize(None, 10**9, 10**9) is None


class _ShapeForbidden:
    """Stand-in range whose shape must not be resolved."""

    def __init__(self, read=None, write=None):
        self.impl = SimpleNamespace(max_cells_per_read=read, max_cells_per_write=write)

    @property
    def shape(self):
        raise AssertionError("shape must not be resolved")


@pytest.mark.parametrize("chunksize", [7, None, 0])
def test_explicit_chunksize_skips_shape_lookup(chunksize):
    rng = _ShapeForbidden(read=1, write=1)
    options = {"chunksize": chunksize}
    assert standard._resolve_read_chunksize(options, rng) == chunksize
    assert (
        standard._resolve_write_chunksize(options, rng, None, scalar=True) == chunksize
    )


def test_none_budget_skips_shape_lookup():
    rng = _ShapeForbidden(read=None, write=None)
    assert standard._resolve_read_chunksize({}, rng) is None
    assert standard._resolve_write_chunksize({}, rng, None, scalar=True) is None


def test_write_resolver_sizes_non_scalar_from_value():
    rng = _ShapeForbidden(write=4)
    value = [[1, 2], [3, 4], [5, 6]]  # 6 cells > 4 -> 2 rows per chunk
    assert standard._resolve_write_chunksize({}, rng, value, scalar=False) == 2
    assert standard._resolve_write_chunksize({}, rng, value[:2], scalar=False) is None


# --- synchronous reads ------------------------------------------------------------


def test_read_at_or_below_budget_is_direct(sheet, read_budget, raw_reads):
    read_budget(18)  # 6 x 3
    assert sheet["A1:C6"].value == _values()
    assert [a for a, _ in raw_reads] == ["$A$1:$C$6"]


def test_read_above_budget_chunks_rows(sheet, read_budget, raw_reads):
    read_budget(7)  # 7 // 3 -> 2 rows per chunk
    assert sheet["A1:C6"].value == _values()
    assert [a for a, _ in raw_reads] == ["$A$1:$C$2", "$A$3:$C$4", "$A$5:$C$6"]


def test_read_final_chunk_is_the_remainder(sheet, read_budget, raw_reads):
    read_budget(12)  # 4 rows per chunk -> 4 + 2
    assert sheet["A1:C6"].value == _values()
    assert [a for a, _ in raw_reads] == ["$A$1:$C$4", "$A$5:$C$6"]


def test_read_offset_anchor_chunks_absolute_addresses(sheet, read_budget, raw_reads):
    read_budget(3)  # 2 cols -> 1 row per chunk
    assert sheet["B2:C4"].value == [[4, 5], [7, 8], [10, 11]]
    assert [a for a, _ in raw_reads] == ["$B$2:$C$2", "$B$3:$C$3", "$B$4:$C$4"]


def test_explicit_chunksize_wins_over_budget(sheet, read_budget, raw_reads):
    read_budget(None)
    assert sheet["A1:C6"].options(chunksize=4).value == _values()
    assert [a for a, _ in raw_reads] == ["$A$1:$C$4", "$A$5:$C$6"]


@pytest.mark.parametrize("chunksize", [None, 0])
def test_explicit_opt_out_disables_budget(sheet, read_budget, raw_reads, chunksize):
    read_budget(1)
    assert sheet["A1:C6"].options(chunksize=chunksize).value == _values()
    assert [a for a, _ in raw_reads] == ["$A$1:$C$6"]


def test_engine_override_none_restores_direct_read(sheet, read_budget, raw_reads):
    read_budget(None)
    assert sheet["A1:C6"].value == _values()
    assert [a for a, _ in raw_reads] == ["$A$1:$C$6"]


def test_single_cell_read_stays_direct_on_scalar_engine(sheet, com_style_reads):
    # Windows COM returns a scalar for one cell; the default must not route it
    # through the chunk loop (which indexes raw_value[0]).
    assert sheet["B2"].value == 4
    assert com_style_reads == ["$B$2"]


def test_scalar_remainder_chunk_on_scalar_engine(sheet, read_budget, com_style_reads):
    read_budget(2)  # one column -> 2 rows per chunk -> 2 + 2 + 1
    assert sheet["A1:A5"].value == [0, 3, 6, 9, 12]
    assert com_style_reads == ["$A$1:$A$2", "$A$3:$A$4", "$A$5"]


def test_chunksize_one_on_a_column_with_scalar_engine(sheet, com_style_reads):
    assert sheet["A1:A3"].options(chunksize=1).value == [0, 3, 6]
    assert com_style_reads == ["$A$1", "$A$2", "$A$3"]


def test_single_row_read(sheet, read_budget, raw_reads):
    read_budget(2)  # 3 cols -> max(1, 2 // 3) = 1 row per chunk
    assert sheet["A1:C1"].value == [0, 1, 2]
    assert [a for a, _ in raw_reads] == ["$A$1:$C$1"]


def test_ndim_1_with_automatic_chunking(sheet, read_budget, raw_reads):
    read_budget(2)
    assert sheet["A1:A6"].options(ndim=1).value == [0, 3, 6, 9, 12, 15]
    assert len(raw_reads) == 3


def test_err_to_str_reaches_every_chunk(sheet, read_budget, raw_reads):
    read_budget(7)
    sheet["A1:C6"].options(err_to_str=True).value
    assert len(raw_reads) == 3
    assert all(options.get("err_to_str") is True for _, options in raw_reads)

    raw_reads.clear()
    sheet["A1:C6"].options(err_to_str=True, chunksize=3).value
    assert len(raw_reads) == 2
    assert all(options.get("err_to_str") is True for _, options in raw_reads)


def test_no_range_read_leaves_value_untouched(monkeypatch):
    # Pre-materialized UDF arguments go through conversion.read(None, ...)
    def boom(*args, **kwargs):
        raise AssertionError("resolver must not run without a range")

    monkeypatch.setattr(standard, "_resolve_read_chunksize", boom)
    value = [[1, 2], [3, 4]]
    assert conversion.read(None, value, {}, engine_name="remote") == value
    assert conversion.read(None, [[5]], {"ndim": 2}, engine_name="remote") == [[5]]


@pytest.mark.skipif(np is None or pd is None, reason="numpy/pandas not installed")
def test_chunked_reads_equal_unchunked(sheet, read_budget):
    expected_list = sheet["A1:C6"].options(chunksize=None).value
    expected_np = sheet["A1:C6"].options(np.array, chunksize=None).value
    expected_df = sheet["A1:C6"].options(pd.DataFrame, chunksize=None).value
    read_budget(5)
    assert sheet["A1:C6"].value == expected_list
    np.testing.assert_array_equal(sheet["A1:C6"].options(np.array).value, expected_np)
    pd.testing.assert_frame_equal(
        sheet["A1:C6"].options(pd.DataFrame).value, expected_df
    )


# --- Calamine ---------------------------------------------------------------------


@pytest.fixture
def calamine_book():
    book = xw.Book(this_dir.parent / "cell_errors.xlsx", mode="r")
    yield book
    book.close()


@pytest.fixture
def calamine_range_reads(monkeypatch):
    calls = []
    original = _xlcalamine.xlwingslib.get_range_values

    def wrapper(*args):
        calls.append(args)
        return original(*args)

    monkeypatch.setattr(_xlcalamine.xlwingslib, "get_range_values", wrapper)
    return calls


def test_calamine_whole_sheet_shortcut_is_preserved(
    calamine_book, calamine_range_reads
):
    sheet = calamine_book.sheets[0]
    values = sheet.cells.value
    assert isinstance(values, list) and len(values) < 100  # used portion only
    assert calamine_range_reads == []


def test_calamine_ordinary_read_is_direct(calamine_book, calamine_range_reads):
    sheet = calamine_book.sheets[0]
    sheet["A1:A5"].value
    assert len(calamine_range_reads) == 1


def test_calamine_explicit_chunks_keep_err_to_str(calamine_book, calamine_range_reads):
    sheet = calamine_book.sheets[0]
    direct = sheet["A1:A5"].options(err_to_str=True, chunksize=None).value
    assert any(isinstance(v, str) and v.startswith("#") for v in direct)
    calamine_range_reads.clear()

    chunked = sheet["A1:A5"].options(err_to_str=True, chunksize=2).value
    assert chunked == direct
    assert len(calamine_range_reads) == 3
    assert all(call[-1] is True for call in calamine_range_reads)


# --- asynchronous (xlwings Lite) reads --------------------------------------------


def test_async_read_below_budget_is_direct(sheet, read_budget, live_js):
    read_budget(18)
    assert _run(sheet["A1:C6"].get_value()) == _values()
    assert live_js.addresses == ["$A$1:$C$6"]


def test_async_read_above_budget_chunks(sheet, read_budget, live_js):
    read_budget(7)
    assert _run(sheet["A1:C6"].get_value()) == _values()
    assert live_js.addresses == ["$A$1:$C$2", "$A$3:$C$4", "$A$5:$C$6"]


def test_async_expansion_happens_before_budget_decision(sheet, read_budget, live_js):
    read_budget(7)
    assert _run(sheet["A1"].options(expand="table").get_value()) == _values()
    assert live_js.addresses == ["$A$1:$C$2", "$A$3:$C$4", "$A$5:$C$6"]


def test_async_explicit_chunksize_and_opt_out(sheet, read_budget, live_js):
    read_budget(1)
    assert _run(sheet["A1:C6"].options(chunksize=None).get_value()) == _values()
    assert live_js.addresses == ["$A$1:$C$6"]
    live_js.addresses.clear()
    read_budget(None)
    assert _run(sheet["A1:C6"].options(chunksize=4).get_value()) == _values()
    assert live_js.addresses == ["$A$1:$C$4", "$A$5:$C$6"]


def test_async_direct_read_null_raises_diagnostic(sheet, live_js):
    live_js.responses[0] = None
    with pytest.raises(XlwingsError) as excinfo:
        _run(sheet["A1:C6"].get_value())
    message = str(excinfo.value)
    assert "'S'!$A$1:$C$6" in message
    assert "chunksize" in message


def test_async_wide_range_null_diagnostic_does_not_recommend_larger_chunks(
    sheet, live_js
):
    # 1,000 columns means the default uses 4,000 rows per chunk. The old
    # diagnostic suggested 10,000 rows, exceeding even the 5M-cell limit.
    live_js.responses[0] = None
    with pytest.raises(XlwingsError) as excinfo:
        _run(sheet["A1:ALL6000"].get_value())
    assert live_js.addresses == ["$A$1:$ALL$4000"]
    message = str(excinfo.value)
    assert "smaller explicit chunksize" in message
    assert "chunksize=10_000" not in message


def test_async_later_chunk_jsnull_raises_with_chunk_address(
    sheet, read_budget, live_js, fake_pyodide
):
    read_budget(7)
    live_js.responses[1] = fake_pyodide
    with pytest.raises(XlwingsError) as excinfo:
        _run(sheet["A1:C6"].get_value())
    assert "'S'!$A$3:$C$4" in str(excinfo.value)
    assert live_js.addresses == ["$A$1:$C$2", "$A$3:$C$4"]  # stopped at the failure


def test_async_rejected_read_propagates_unchanged(sheet, read_budget, live_js):
    read_budget(7)
    live_js.responses[1] = RuntimeError("payload too large")
    with pytest.raises(RuntimeError, match="payload too large"):
        _run(sheet["A1:C6"].get_value())


# --- writes -----------------------------------------------------------------------


def test_write_at_or_below_budget_is_direct(sheet, book, write_budget):
    write_budget(18)
    sheet["A1"].value = _values()
    actions = _set_values_actions(book)
    assert len(actions) == 1
    assert actions[0]["row_count"] == 6 and actions[0]["values"] == _values()


def test_write_above_budget_chunks_from_offset_anchor(sheet, book, write_budget):
    write_budget(7)  # 2 rows per chunk
    sheet["B3"].value = _values()
    actions = _set_values_actions(book)
    assert [(a["start_row"], a["start_column"], a["row_count"]) for a in actions] == [
        (2, 1, 2),
        (4, 1, 2),
        (6, 1, 2),
    ]
    assert [a["values"] for a in actions] == [
        _values()[0:2],
        _values()[2:4],
        _values()[4:6],
    ]


def test_write_final_chunk_is_the_remainder(sheet, book, write_budget):
    write_budget(12)  # 4 rows per chunk -> 4 + 2
    sheet["A1"].value = _values()
    actions = _set_values_actions(book)
    assert [(a["start_row"], a["row_count"]) for a in actions] == [(0, 4), (4, 2)]
    assert actions[1]["values"] == _values()[4:6]


def test_write_explicit_chunksize_wins_even_without_budget(sheet, book, write_budget):
    write_budget(None)
    sheet["A1"].options(chunksize=4).value = _values()
    assert [a["row_count"] for a in _set_values_actions(book)] == [4, 2]


@pytest.mark.parametrize("chunksize", [None, 0])
def test_write_explicit_opt_out(sheet, book, write_budget, chunksize):
    write_budget(1)
    sheet["A1"].options(chunksize=chunksize).value = _values()
    assert [a["row_count"] for a in _set_values_actions(book)] == [6]


def test_write_chunksize_one(sheet, book):
    sheet["A1"].options(chunksize=1).value = _values()
    assert [a["start_row"] for a in _set_values_actions(book)] == [0, 1, 2, 3, 4, 5]


def test_write_single_row_and_single_column(sheet, book, write_budget):
    write_budget(2)
    sheet["A1"].value = [[1, 2, 3, 4, 5, 6]]  # 1 x 6: chunksize 1 -> one chunk
    assert [a["row_count"] for a in _set_values_actions(book)] == [1]
    book.impl._json = {"actions": []}
    sheet["A1"].value = [[1], [2], [3], [4], [5]]  # 5 x 1: 2 rows per chunk
    assert [(a["start_row"], a["row_count"]) for a in _set_values_actions(book)] == [
        (0, 2),
        (2, 2),
        (4, 1),
    ]


@pytest.mark.skipif(np is None or pd is None, reason="numpy/pandas not installed")
def test_write_sizes_dataframe_numpy_and_transpose_after_conversion(
    sheet, book, write_budget
):
    write_budget(9)  # 3 rows per chunk at 3 columns
    df = pd.DataFrame({"a": range(5), "b": range(5)})  # 5x2 -> 6x3 with index/header
    sheet["A1"].value = df
    assert [
        (a["start_row"], a["row_count"], a["column_count"])
        for a in _set_values_actions(book)
    ] == [
        (0, 3, 3),
        (3, 3, 3),
    ]
    book.impl._json = {"actions": []}

    sheet["A1"].value = np.arange(18).reshape(6, 3)
    assert [a["row_count"] for a in _set_values_actions(book)] == [3, 3]
    book.impl._json = {"actions": []}

    sheet["A1"].options(transpose=True).value = _values()  # -> 3 x 6, 1 row/chunk
    actions = _set_values_actions(book)
    assert [(a["row_count"], a["column_count"]) for a in actions] == [(1, 6)] * 3
    assert actions[1]["values"] == [[1, 4, 7, 10, 13, 16]]


@pytest.mark.parametrize("scalar", [5, "ab", False])
def test_scalar_fill_above_budget_chunks_target_range(
    sheet, book, write_budget, scalar
):
    write_budget(7)  # 6 x 3 target -> 2 rows per chunk
    sheet["A1:C6"].value = scalar
    actions = _set_values_actions(book)
    assert [(a["start_row"], a["row_count"]) for a in actions] == [
        (0, 2),
        (2, 2),
        (4, 2),
    ]
    assert all(a["values"] == [[scalar] * 3] * 2 for a in actions)


def test_scalar_fill_small_range_is_direct(sheet, book):
    sheet["A1:C6"].value = 5
    actions = _set_values_actions(book)
    assert len(actions) == 1 and actions[0]["values"] == [[5] * 3] * 6


def test_scalar_fill_explicit_chunksize_does_not_iterate_strings(sheet, book):
    sheet["A1:C6"].options(chunksize=4).value = "ab"
    actions = _set_values_actions(book)
    assert [a["row_count"] for a in actions] == [4, 2]
    assert actions[1]["values"] == [["ab"] * 3] * 2


def test_raw_write_bypasses_chunking(sheet, book, write_budget):
    write_budget(1)
    sheet["A1:C6"].options("raw").value = _values()
    assert [a["row_count"] for a in _set_values_actions(book)] == [6]


def test_formatter_runs_once_per_logical_write(sheet, book, write_budget):
    write_budget(7)
    calls = []
    sheet["A1"].options(
        formatter=lambda rng, value: calls.append(rng.address)
    ).value = _values()
    assert len(_set_values_actions(book)) == 3
    assert calls == ["$A$1:$C$6"]


def test_later_actions_follow_all_write_chunks(sheet, book, write_budget):
    write_budget(7)
    sheet["A1"].value = _values()
    sheet["E1"].clear_contents()
    assert [a["func"] for a in _actions(book)] == [
        "setValues",
        "setValues",
        "setValues",
        "rangeClearContents",
    ]


def test_failing_later_chunk_stops_and_propagates(
    sheet, book, write_budget, monkeypatch
):
    write_budget(7)
    original = _xlremote.Range.raw_value
    count = {"n": 0}

    def setter(self, value):
        count["n"] += 1
        if count["n"] == 2:
            raise RuntimeError("Excel rejected chunk 2")
        original.fset(self, value)

    monkeypatch.setattr(_xlremote.Range, "raw_value", property(original.fget, setter))
    with pytest.raises(RuntimeError, match="chunk 2"):
        sheet["A1"].value = _values()
    assert count["n"] == 2
    assert [a["row_count"] for a in _set_values_actions(book)] == [2]


def test_chunked_write_equals_direct_write(sheet, book, write_budget):
    sheet["A1"].options(chunksize=None).value = _values()
    direct = _set_values_actions(book)[0]["values"]
    book.impl._json = {"actions": []}
    write_budget(5)
    reassembled = []
    sheet["A1"].value = _values()
    for action in _set_values_actions(book):
        reassembled.extend(action["values"])
    assert reassembled == direct
