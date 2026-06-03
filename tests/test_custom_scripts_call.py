"""
Tests for custom_scripts_call argument binding: positional args, defaults,
*args, missing/extra argument errors, and typehint injection.
"""

import types

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro.udfs_officejs import custom_scripts_call, script


@pytest.fixture
def anyio_backend():
    return "asyncio"


BOOK_JSON = {
    "client": "Office.js",
    "version": xw.__version__,
    "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
    "names": [],
    "sheets": [{"name": "S", "values": [[None]], "pictures": [], "tables": []}],
}


def _make_module(**funcs):
    """Create a module with the given functions as attributes."""
    mod = types.ModuleType("test_scripts")
    for name, func in funcs.items():
        mod.__dict__[name] = func
    return mod


def _get_actions(book):
    """Extract the actions list from a book's JSON response."""
    result = book.json()
    return result.get("actions", [])


# --- Happy path ---


@pytest.mark.anyio
async def test_args_passed_positionally():
    @script
    def my_script(book: xw.Book, name: str, count: int):
        book.sheets.active["A1"].value = f"{name}-{count}"

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}, args=["hello", 3]
    )
    actions = _get_actions(result)
    assert len(actions) == 1
    assert actions[0]["values"] == [["hello-3"]]
    book.close()


@pytest.mark.anyio
async def test_default_values_used_when_arg_omitted():
    @script
    def my_script(book: xw.Book, value: str, target: str = "A1"):
        book.sheets.active[target].value = value

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}, args=["test"]
    )
    actions = _get_actions(result)
    assert len(actions) == 1
    assert actions[0]["values"] == [["test"]]
    book.close()


@pytest.mark.anyio
async def test_no_args_backward_compat():
    @script
    def my_script(book: xw.Book):
        book.sheets.active["A1"].value = "done"

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}
    )
    actions = _get_actions(result)
    assert len(actions) == 1
    assert actions[0]["values"] == [["done"]]
    book.close()


@pytest.mark.anyio
async def test_var_positional_consumes_remaining():
    @script
    def my_script(book: xw.Book, *values):
        book.sheets.active["A1"].value = ",".join(str(v) for v in values)

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}, args=["a", "b", "c"]
    )
    actions = _get_actions(result)
    assert len(actions) == 1
    assert actions[0]["values"] == [["a,b,c"]]
    book.close()


# --- Error cases ---


@pytest.mark.anyio
async def test_missing_required_arg():
    @script
    def my_script(book: xw.Book, name: str):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="missing required argument"):
        await custom_scripts_call(
            mod, "my_script", typehint_to_value={xw.Book: book}, args=[]
        )
    book.close()


@pytest.mark.anyio
async def test_extra_args():
    @script
    def my_script(book: xw.Book, name: str):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="extra argument"):
        await custom_scripts_call(
            mod, "my_script", typehint_to_value={xw.Book: book}, args=["a", "b"]
        )
    book.close()


@pytest.mark.anyio
async def test_keyword_only_rejected():
    @script
    def my_script(book: xw.Book, *, mode: str):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="keyword-only"):
        await custom_scripts_call(
            mod, "my_script", typehint_to_value={xw.Book: book}, args=["fast"]
        )
    book.close()


@pytest.mark.anyio
async def test_var_keyword_rejected():
    @script
    def my_script(book: xw.Book, **kwargs):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="keyword-only"):
        await custom_scripts_call(
            mod, "my_script", typehint_to_value={xw.Book: book}, args=[]
        )
    book.close()
