"""
Tests for the Pyodide >= 0.28 ``JsNull`` normalization in the remote backend.

Pyodide >= 0.28 (Python 3.14) is a breaking change: JavaScript ``null`` is
converted to the sentinel ``pyodide.ffi.jsnull`` (an instance of ``JsNull``)
instead of Python ``None``. JS ``undefined`` still becomes ``None``.

Crucially ``JsNull.__bool__`` returns ``False`` (so truthiness checks are
unaffected), but identity checks (``is None`` / ``is not None``) fail and
arithmetic / ``len()`` / ``int()`` on it raise ``TypeError``. This regressed
three reproduced paths in the remote backend:

* empty optional scalar UDF args no longer fell back to their defaults
  (``to_scalar`` / ``js_to_none``)
* a book-scoped name's ``scope_sheet_index`` arrived as ``JsNull`` -> ``sheet.names``
  did ``is not None`` then ``+ 1`` -> ``TypeError`` (``Books.open``)
* a non-cell (shape) selection's ``address`` arrived as ``JsNull`` -> ``is None``
  guard failed and ``JsNull`` flowed into ``Range()`` -> ``len()`` ``TypeError``
  (``App.get_selection``)
* empty cells in a custom function's *array* argument arrive from Office.js as JS
  ``null`` -> ``JsNull`` (book data uses ``""`` instead), so they weren't read as
  empty cells and ``JsNull`` flowed into user functions (``_clean_value_data_element``)

``pyodide`` is not installed in the test environment, so the production code's
``from pyodide.ffi import JsNull`` normally hits ``except ImportError`` and is a
no-op. The ``fake_pyodide`` fixture injects a faithful fake ``pyodide.ffi``
module so the real normalization logic runs.
"""

import datetime as dt
import sys
from types import ModuleType

import pytest

import xlwings as xw
from xlwings.pro import _xlofficejs, _xlremote
from xlwings.pro.udfs_officejs import js_to_none, to_scalar


class _FakeJsNull:
    """Mimics pyodide.ffi.JsNull: a falsy singleton sentinel for JS null."""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = object.__new__(cls)
        return cls._instance

    def __bool__(self):
        return False

    def __repr__(self):
        return "jsnull"


@pytest.fixture
def fake_pyodide(monkeypatch):
    """Inject a fake ``pyodide.ffi`` exposing ``JsNull``/``jsnull`` so the
    production ``from pyodide.ffi import JsNull`` resolves during the test.

    Yields the singleton ``jsnull`` instance for use in assertions.
    """
    jsnull = _FakeJsNull()

    pyodide = ModuleType("pyodide")
    ffi = ModuleType("pyodide.ffi")
    ffi.JsNull = _FakeJsNull
    ffi.jsnull = jsnull
    pyodide.ffi = ffi

    monkeypatch.setitem(sys.modules, "pyodide", pyodide)
    monkeypatch.setitem(sys.modules, "pyodide.ffi", ffi)
    yield jsnull


@pytest.fixture
def anyio_backend():
    return "asyncio"


# --- _FakeJsNull sanity: matches the real JsNull semantics we rely on ---


def test_fake_jsnull_is_falsy_and_singleton():
    a = _FakeJsNull()
    b = _FakeJsNull()
    assert a is b
    assert not a
    assert bool(a) is False
    # The whole point: identity vs None fails, which is what broke production.
    assert a is not None


# --- _normalize_jsnull ---


def test_normalize_scalar_jsnull(fake_pyodide):
    assert _xlremote._normalize_jsnull(fake_pyodide) is None


def test_normalize_passthrough_for_real_values(fake_pyodide):
    assert _xlremote._normalize_jsnull("a") == "a"
    assert _xlremote._normalize_jsnull(0) == 0
    assert _xlremote._normalize_jsnull(None) is None
    assert _xlremote._normalize_jsnull("") == ""


def test_normalize_nested_dict_and_list(fake_pyodide):
    jsnull = fake_pyodide
    obj = {
        "a": jsnull,
        "b": [1, jsnull, {"c": jsnull, "d": 2}],
        "e": {"f": {"g": jsnull}},
    }
    result = _xlremote._normalize_jsnull(obj)
    assert result == {
        "a": None,
        "b": [1, None, {"c": None, "d": 2}],
        "e": {"f": {"g": None}},
    }


def test_normalize_skips_values_key(fake_pyodide):
    """Cell ``values`` arrays are passed through untouched (Office.js never
    sends ``null`` there) — and notably without being recursed into."""
    jsnull = fake_pyodide
    # A JsNull *inside* values is intentionally left as-is (can't happen in
    # practice, but proves we don't walk it).
    values = [["x", jsnull], [1, 2]]
    obj = {"scope_sheet_index": jsnull, "values": values}
    result = _xlremote._normalize_jsnull(obj)
    assert result["scope_sheet_index"] is None
    # Same object, not a normalized copy:
    assert result["values"] is values
    assert result["values"][0][1] is jsnull


def test_normalize_noop_without_pyodide():
    """With no ``pyodide`` importable, the input is returned unchanged."""
    sentinel = object()
    obj = {"a": sentinel}
    # Don't install fake_pyodide here.
    assert _xlremote._normalize_jsnull(obj) is obj


# --- js_to_none / to_scalar (UDF arg path; the original streaming bug) ---


def test_js_to_none_converts_jsnull(fake_pyodide):
    assert js_to_none(fake_pyodide) is None


def test_js_to_none_passthrough(fake_pyodide):
    assert js_to_none(5) == 5
    assert js_to_none("x") == "x"
    assert js_to_none(None) is None


def test_to_scalar_unwraps_and_normalizes_jsnull(fake_pyodide):
    # An empty Excel cell arrives as [[JsNull]]; to_scalar must unwrap to None
    # so the optional-argument default kicks in downstream.
    assert to_scalar([[fake_pyodide]]) is None
    assert to_scalar([fake_pyodide]) is None
    assert to_scalar(fake_pyodide) is None
    # Normal scalar unwrapping still works.
    assert to_scalar([[7]]) == 7


# --- _is_jsnull (officejs engine; the UDF read path) ---


def test_is_jsnull_true_for_jsnull(fake_pyodide):
    assert _xlofficejs._is_jsnull(fake_pyodide) is True


def test_is_jsnull_false_for_real_values(fake_pyodide):
    assert _xlofficejs._is_jsnull(None) is False
    assert _xlofficejs._is_jsnull("") is False
    assert _xlofficejs._is_jsnull(0) is False
    assert _xlofficejs._is_jsnull("x") is False


def test_is_jsnull_false_without_pyodide():
    """With no ``pyodide`` importable, nothing is a JsNull."""
    assert _xlofficejs._is_jsnull(object()) is False


# --- officejs engine clean_value_data (UDF array-arg path; empty cells as JsNull) ---
#
# Custom functions use the *officejs* engine (xlwings.engines["officejs"], impl =
# _xlofficejs.engine), NOT _xlremote — scripts/runPython use _xlremote. The bug was
# that empty cells in a UDF range argument arrive from Office.js as JS null -> JsNull
# and weren't normalized in _xlofficejs._clean_value_data_element. Test that engine.


def test_officejs_clean_value_data_element_treats_jsnull_as_empty(fake_pyodide):
    assert (
        _xlofficejs._clean_value_data_element(
            fake_pyodide, dt.datetime, None, None, False
        )
        is None
    )
    sentinel = object()
    assert (
        _xlofficejs._clean_value_data_element(
            fake_pyodide, dt.datetime, sentinel, None, False
        )
        is sentinel
    )


def test_officejs_clean_value_data_maps_jsnull_in_2d_array(fake_pyodide):
    """An empty cell inside a UDF range argument arrives as JsNull and must be
    normalized to ``empty_as`` (the default ``None``)."""
    data = [["x", fake_pyodide], [fake_pyodide, 2]]
    result = _xlofficejs.Engine.clean_value_data(data, dt.datetime, None, None, False)
    assert result == [["x", None], [None, 2]]


@pytest.mark.anyio
async def test_custom_function_list_arg_normalizes_jsnull(fake_pyodide):
    """End-to-end through ``custom_functions_call``: a ``list``-typed UDF arg
    containing empty cells (JsNull) must reach the user function as ``None``,
    not the JsNull sentinel. This exercises the real engine resolution, so it
    guards against the fix landing in the wrong engine module."""
    from types import ModuleType

    from xlwings.pro.udfs_officejs import custom_functions_call, xlfunc

    received = {}

    def hello(values: list):
        received["values"] = values
        return "ok"

    module = ModuleType("usermod")
    module.hello = xlfunc(hello)

    data = {
        "func_name": "hello",
        "version": xw.__version__,
        "client": "Office.js",
        "runtime": "1.4",
        "args": [[[1, fake_pyodide], [fake_pyodide, 4]]],
    }
    await custom_functions_call(data, module)
    assert received["values"] == [[1, None], [None, 4]]


# --- Integration: book-scoped name with JsNull scope (sheet.names crash) ---


def _book_json_with_book_scoped_name(scope_sheet_index):
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [
            {
                "name": "myrange",
                "sheet_index": 0,
                "address": "A1",
                "scope_sheet_name": None,
                "scope_sheet_index": scope_sheet_index,
                "book_scope": True,
            }
        ],
        "sheets": [{"name": "S", "values": [[None]], "pictures": [], "tables": []}],
    }


def test_sheet_names_does_not_crash_on_jsnull_scope(fake_pyodide):
    """Reproduces the ``JsNull + 1`` crash: a book-scoped name's
    ``scope_sheet_index`` is JS null. ``xw.Book(json=...)`` -> ``Books.open()``
    must normalize it so ``sheet.names`` filters cleanly instead of raising."""
    book = xw.Book(json=_book_json_with_book_scoped_name(fake_pyodide))
    try:
        # Book-scoped name must not appear in a sheet's (sheet-scoped) names,
        # and iterating must not raise TypeError.
        sheet_names = list(book.sheets[0].names)
        assert sheet_names == []
        # The stored api value was normalized to a real None.
        assert book.api["names"][0]["scope_sheet_index"] is None
    finally:
        book.close()


def test_sheet_scoped_name_still_resolved_after_normalize(fake_pyodide):
    """Control: a real (integer) sheet scope is untouched and still resolves."""
    book = xw.Book(json=_book_json_with_book_scoped_name(0))
    try:
        # book_scope=True excludes it from sheet.names regardless, so flip it
        # to sheet-scoped (with a real scope_sheet_name, as Excel would send).
        book.api["names"][0]["book_scope"] = False
        book.api["names"][0]["scope_sheet_name"] = "S"
        sheet_names = [n.name for n in book.sheets[0].names]
        assert sheet_names == ["S!myrange"]
    finally:
        book.close()


# --- Integration: shape (non-cell) selection -> address JsNull (Range crash) ---


class _FakeJsProxy:
    """Minimal JsProxy-like wrapper whose ``.to_py()`` returns ``data``."""

    def __init__(self, data):
        self._data = data

    def to_py(self):
        return self._data


def _install_fake_js(monkeypatch, selection_result):
    """Install a fake ``js`` module so ``await js.xlwings.getSelection()``
    resolves to a JsProxy wrapping ``selection_result``."""

    async def get_selection():
        return _FakeJsProxy(selection_result)

    js = ModuleType("js")
    js.xlwings = ModuleType("js.xlwings")
    js.xlwings.getSelection = get_selection
    monkeypatch.setitem(sys.modules, "js", js)


def _book_json_minimal():
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [{"name": "S", "values": [["x"]], "pictures": [], "tables": []}],
    }


@pytest.mark.anyio
async def test_note_get_text_normalizes_jsnull(fake_pyodide, monkeypatch):
    async def get_note_text(sheet_name, address):
        assert sheet_name == "S"
        assert address == "$A$1"
        return fake_pyodide

    js = ModuleType("js")
    js.xlwings = ModuleType("js.xlwings")
    js.xlwings.getNoteText = get_note_text
    monkeypatch.setitem(sys.modules, "js", js)
    monkeypatch.setattr(sys, "platform", "emscripten")

    payload = _book_json_minimal()
    payload["sheets"][0]["notes"] = [{"address": "A1"}]
    book = xw.Book(json=payload)
    try:
        note = book.sheets[0]["A1"].note
        assert note is not None
        assert await note.get_text() is None
    finally:
        book.close()


@pytest.mark.anyio
async def test_get_selection_returns_none_for_jsnull_address(fake_pyodide, monkeypatch):
    """Reproduces the shape-selection crash: a non-cell selection yields
    ``address: null`` -> ``JsNull``. After normalization the ``is None`` guard
    must fire and return ``None`` instead of building ``Range(JsNull)`` (which
    crashed on ``len()``)."""
    monkeypatch.setattr(sys, "platform", "emscripten")
    _install_fake_js(monkeypatch, {"sheetIndex": 0, "address": fake_pyodide})

    book = xw.Book(json=_book_json_minimal())
    try:
        app = book.app
        result = await app.get_selection()
        assert result is None
    finally:
        book.close()


@pytest.mark.anyio
async def test_get_selection_returns_range_for_cell_address(fake_pyodide, monkeypatch):
    """Control: a real cell selection still resolves to a Range."""
    monkeypatch.setattr(sys, "platform", "emscripten")
    _install_fake_js(monkeypatch, {"sheetIndex": 0, "address": "B2"})

    book = xw.Book(json=_book_json_minimal())
    try:
        result = await book.app.get_selection()
        assert result is not None
        assert result.address.endswith("B$2")
    finally:
        book.close()
