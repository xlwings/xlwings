"""
Tests for custom_scripts_call argument binding: positional args, defaults,
*args, missing/extra argument errors, and typehint injection.
"""

import datetime as dt
import types
from typing import Optional, Union

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


class FakeCurrentUser:
    """Stands in for xlwings Server's CurrentUser, i.e. a framework-injected value."""

    def __init__(self, name="alice"):
        self.name = name
        self.roles = []

    async def has_required_roles(self, required_roles):
        return True


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


@pytest.mark.anyio
async def test_injectable_var_keyword_still_rejected():
    # **kwargs is rejected even when annotated with an injectable typehint: the
    # exemption is for keyword-only params, not for arbitrary keyword collection.
    @script
    def my_script(book: xw.Book, **kwargs: FakeCurrentUser):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="keyword-only"):
        await custom_scripts_call(
            mod,
            "my_script",
            typehint_to_value={xw.Book: book, FakeCurrentUser: FakeCurrentUser()},
            args=[],
        )
    book.close()


# --- Keyword-only params for framework-injected values ---


@pytest.mark.anyio
@pytest.mark.parametrize(
    "signature, args",
    [
        ("book: xw.Book, *args: str, user: FakeCurrentUser", ["a", "b"]),
        ("book: xw.Book, *, user: FakeCurrentUser", []),
        ("book: xw.Book, name: str, *, user: FakeCurrentUser", ["a"]),
    ],
    ids=["after-varargs", "bare-star", "with-positional"],
)
async def test_keyword_only_injected_value(signature, args):
    """An injected value may sit after *args, where the caller can't reach it."""
    namespace = {"xw": xw, "FakeCurrentUser": FakeCurrentUser, "script": script}
    exec(
        f"""
@script
def my_script({signature}):
    book.sheets.active["A1"].value = user.name
""",
        namespace,
    )

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=namespace["my_script"])
    result = await custom_scripts_call(
        mod,
        "my_script",
        typehint_to_value={xw.Book: book, FakeCurrentUser: FakeCurrentUser()},
        args=args,
    )
    actions = _get_actions(result)
    assert len(actions) == 1
    assert actions[0]["values"] == [["alice"]]
    book.close()


@pytest.mark.anyio
async def test_keyword_only_book_is_injected_and_returned():
    # The book may itself be keyword-only: the @script decorator has to find it in
    # kwargs to return it, not just among the positional args.
    @script
    def my_script(user: FakeCurrentUser, *args: str, book: xw.Book):
        book.sheets.active["A1"].value = user.name

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod,
        "my_script",
        typehint_to_value={xw.Book: book, FakeCurrentUser: FakeCurrentUser()},
        args=[],
    )
    assert result is book
    assert _get_actions(result)[0]["values"] == [["alice"]]
    book.close()


@pytest.mark.anyio
async def test_keyword_only_book_async_stays_lazy():
    # The BookAsync -> Book normalization and the lazy flag must also apply when
    # the book is keyword-only.
    @script
    async def my_script(*args: str, book: xw.BookAsync):
        pass

    book = xw.Book(json=BOOK_JSON)
    assert book.impl._lazy is False
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(mod, "my_script", typehint_to_value={xw.Book: book})
    assert book.impl._lazy is True
    book.close()


@pytest.mark.anyio
async def test_keyword_only_arg_not_supplied_by_caller_is_rejected():
    # Only values the framework provides are exempt. A keyword-only param whose
    # type isn't in typehint_to_value could never be filled, so it still raises.
    @script
    def my_script(book: xw.Book, *args: str, user: FakeCurrentUser):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="keyword-only"):
        await custom_scripts_call(
            mod, "my_script", typehint_to_value={xw.Book: book}, args=[]
        )
    book.close()


@pytest.mark.anyio
async def test_positional_args_unaffected_by_keyword_only_injection():
    # The injected keyword-only param must not consume any caller arg: the
    # positional args keep binding in order.
    @script
    def my_script(book: xw.Book, a: str, b: str, *, user: FakeCurrentUser):
        book.sheets.active["A1"].value = f"{user.name}|{a}|{b}"

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod,
        "my_script",
        typehint_to_value={xw.Book: book, FakeCurrentUser: FakeCurrentUser()},
        args=["x", "y"],
    )
    assert _get_actions(result)[0]["values"] == [["alice|x|y"]]
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
async def test_optional_book_is_injected_and_returned():
    # `= None` makes get_type_hints() report Optional[Book]. The injection path
    # unwraps it, so the book-detection paths must agree - otherwise the script
    # runs and then raises "No xlwings.Book found".
    @script
    def my_script(book: Optional[xw.Book] = None):
        book.sheets.active["A1"].value = "optional"

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}
    )
    assert result is book
    assert _get_actions(result)[0]["values"] == [["optional"]]
    book.close()


def test_optional_book_async_is_detected_as_async():
    # An Optional[BookAsync] annotation must still opt into the async API,
    # i.e. produce lazy=True metadata rather than silently staying sync.
    @script
    async def my_script(book: Optional[xw.BookAsync] = None):
        pass

    assert my_script.__xlscript__["lazy"] is True


@pytest.mark.anyio
async def test_optional_book_async_marks_injected_book_lazy():
    @script
    async def my_script(book: Optional[xw.BookAsync] = None):
        pass

    book = xw.Book(json=BOOK_JSON)
    assert book.impl._lazy is False
    mod = _make_module(my_script=my_script)
    result = await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}
    )
    assert result is book
    assert book.impl._lazy is True
    book.close()


def test_optional_book_still_counts_toward_multiple_book_params():
    # Optional book hints must be seen by the "exactly one book param" guard too.
    with pytest.raises(XlwingsError, match="exactly one parameter"):

        @script
        def my_script(a: Optional[xw.Book] = None, b: xw.Book = None):
            pass


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


# --- Date coercion ---
# JSON has no date type, so date arguments arrive as ISO strings and are
# converted via the type hint, matching how custom functions handle dates.
# These assert on what the function *received*: writing a date to a cell
# serializes it back to a string in the action payload.


@pytest.mark.anyio
async def test_date_and_datetime_hints_are_coerced():
    got = {}

    @script
    def my_script(book: xw.Book, d: dt.date, ts: dt.datetime, name: str):
        got.update(d=d, ts=ts, name=name)

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(
        mod,
        "my_script",
        typehint_to_value={xw.Book: book},
        args=["2026-07-30", "2026-07-30T14:30", "2026-07-30"],
    )
    assert got["d"] == dt.date(2026, 7, 30)
    assert got["ts"] == dt.datetime(2026, 7, 30, 14, 30)
    assert got["name"] == "2026-07-30"  # a str hint stays a string
    book.close()


@pytest.mark.anyio
@pytest.mark.parametrize(
    "hint", [Optional[dt.date], "dt.date | None", Union[dt.date, None]]
)
async def test_optional_date_hints_are_coerced(hint):
    # An optional date still gets a date control in the UI, so it must coerce
    # through the union rather than staying a string.
    got = {}

    def my_script(book, d=None):
        got["d"] = d

    my_script.__annotations__ = {"book": xw.Book, "d": hint}
    my_script = script(my_script)

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(
        mod, "my_script", typehint_to_value={xw.Book: book}, args=["2026-07-30"]
    )
    assert got["d"] == dt.date(2026, 7, 30)
    book.close()


@pytest.mark.anyio
async def test_optional_date_default_stays_none():
    got = {}

    @script
    def my_script(book: xw.Book, d: Optional[dt.date] = None):
        got["d"] = d

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(mod, "my_script", typehint_to_value={xw.Book: book})
    assert got["d"] is None
    book.close()


@pytest.mark.anyio
async def test_date_default_is_not_reparsed():
    # Python defaults are already real objects, so they must pass through.
    got = {}

    @script
    def my_script(book: xw.Book, d: dt.date = dt.date(2001, 2, 3)):
        got["d"] = d

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(mod, "my_script", typehint_to_value={xw.Book: book})
    assert got["d"] == dt.date(2001, 2, 3)
    book.close()


@pytest.mark.anyio
async def test_malformed_date_raises():
    @script
    def my_script(book: xw.Book, d: dt.date):
        pass

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    with pytest.raises(XlwingsError, match="must be a date in ISO format"):
        await custom_scripts_call(
            mod, "my_script", typehint_to_value={xw.Book: book}, args=["not-a-date"]
        )
    book.close()


@pytest.mark.anyio
async def test_var_positional_dates_are_coerced():
    got = {}

    @script
    def my_script(book: xw.Book, *dates: dt.date):
        got["dates"] = dates

    book = xw.Book(json=BOOK_JSON)
    mod = _make_module(my_script=my_script)
    await custom_scripts_call(
        mod,
        "my_script",
        typehint_to_value={xw.Book: book},
        args=["2026-01-02", "2026-03-04"],
    )
    assert got["dates"] == (dt.date(2026, 1, 2), dt.date(2026, 3, 4))
    book.close()
