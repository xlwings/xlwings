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


# --- Sync/async selection (BookAsync annotation + deprecated lazy=) ---


def test_book_annotation_defaults_to_sync():
    @script
    def my_script(book: xw.Book):
        pass

    # Plain xw.Book maps to the internal lazy=False wire key.
    assert my_script.__xlscript__["lazy"] is False


def test_lazy_true_is_deprecated_alias():
    with pytest.warns(UserWarning, match="'lazy'.*deprecated.*BookAsync"):

        @script(lazy=True)
        def my_script(book: xw.Book):
            pass

    assert my_script.__xlscript__["lazy"] is True
    # The deprecated kwarg is consumed, not leaked into the metadata twice.
    assert list(my_script.__xlscript__.keys()).count("lazy") == 1


def test_lazy_false_is_deprecated_alias():
    with pytest.warns(UserWarning):

        @script(lazy=False)
        def my_script(book: xw.Book):
            pass

    assert my_script.__xlscript__["lazy"] is False


def test_lazy_non_boolean_rejected():
    # bool("false") is True, so a stringy `lazy` must be rejected, not coerced.
    with pytest.raises(XlwingsError, match="'lazy'.*must be a boolean"):

        @script(lazy="false")
        def my_script(book: xw.Book):
            pass


# --- BookAsync annotation ---


def test_book_async_annotation_sets_lazy_true():
    @script
    async def my_script(book: xw.BookAsync):
        pass

    assert my_script.__xlscript__["lazy"] is True


def test_book_async_agrees_with_lazy_true():
    with pytest.warns(UserWarning):

        @script(lazy=True)
        async def my_script(book: xw.BookAsync):
            pass

    assert my_script.__xlscript__["lazy"] is True


def test_book_async_conflicts_with_lazy_false():
    with pytest.raises(XlwingsError, match="BookAsync"):
        with pytest.warns(UserWarning):

            @script(lazy=False)
            async def my_script(book: xw.BookAsync):
                pass


def test_book_async_return_annotation_does_not_enable_async():
    # A BookAsync *return* annotation must not enable the async API — only the
    # injected book parameter's annotation counts.
    @script
    def my_script(book: xw.Book) -> xw.BookAsync:
        return book

    assert my_script.__xlscript__["lazy"] is False


def test_book_async_unrelated_param_does_not_enable_async():
    # A BookAsync annotation on a non-book parameter alongside a sync book is
    # ambiguous (two book-typed params) and must be rejected, not silently
    # treated as async.
    with pytest.raises(XlwingsError, match="exactly one parameter"):

        @script
        def my_script(value: xw.BookAsync, book: xw.Book):
            pass


def test_multiple_book_params_rejected():
    with pytest.raises(XlwingsError, match="exactly one parameter"):

        @script
        def my_script(book1: xw.Book, book2: xw.Book):
            pass


@pytest.mark.anyio
async def test_book_async_annotated_book_is_injected():
    # The injected book is keyed under xw.Book by the caller; a BookAsync
    # annotation must still resolve to it at call time.
    @script
    async def my_script(book: xw.BookAsync):
        book.sheets.active["A1"].value = "async"

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}
    )
    actions = _get_actions(result)
    assert len(actions) == 1
    assert actions[0]["values"] == [["async"]]
    book.close()


@pytest.mark.anyio
async def test_book_async_marks_injected_book_lazy():
    # A BookAsync annotation must mark the injected book lazy, even though the
    # caller constructs it eagerly (xw.Book(json=...), as xlwings Lite does).
    # Sync `.value` reads then raise instead of silently returning None.
    @script
    async def my_script(book: xw.BookAsync):
        book.sheets.active["A1"].value  # sync read on a lazy book -> raises

    book = xw.Book(json=BOOK_JSON)
    assert book.impl._lazy is False  # eager as constructed
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="haven't been loaded"):
        await custom_scripts_call(mod, "my_script", typehint_to_value={xw.Book: book})
    assert book.impl._lazy is True  # annotation flipped it
    book.close()


@pytest.mark.anyio
async def test_plain_book_annotation_stays_eager():
    # Without BookAsync, the injected book stays eager and sync reads work.
    @script
    async def my_script(book: xw.Book):
        book.sheets.active["A1"].value  # must not raise

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(mod, "my_script", typehint_to_value={xw.Book: book})
    assert book.impl._lazy is False
    book.close()
