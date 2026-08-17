import asyncio
import dataclasses
import os
import re
import subprocess
import sys
from types import ModuleType

import pytest

import xlwings as xw
from xlwings.pro.caller import _CALLER_ADDRESS_RE
from xlwings.server import Caller, caller_from_address, parse_caller_address


@pytest.mark.parametrize(
    "caller_address, expected",
    [
        ("[Book1.xlsx]Sheet1!B21", ("Book1.xlsx", "Sheet1", "B21")),
        ("Sheet1!B21", ("", "Sheet1", "B21")),
        ("[Book1.xlsx]'My Sheet'!A1", ("Book1.xlsx", "My Sheet", "A1")),
        ("'My Sheet'!A1:C3", ("", "My Sheet", "A1:C3")),
        # Excel escapes a literal apostrophe in a quoted sheet name by doubling it
        ("[My Book.xlsx]'It''s'!$A$1", ("My Book.xlsx", "It's", "$A$1")),
        ("Sheet1!$B$21:$D$25", ("", "Sheet1", "$B$21:$D$25")),
    ],
)
def test_parse_caller_address(caller_address, expected):
    assert parse_caller_address(caller_address) == expected


@pytest.mark.parametrize(
    "caller_address",
    [
        "garbage",
        "",
        "Sheet1",
        "!A1",
        # An unquoted sheet name can't contain "!": Excel doesn't allow it in a sheet name,
        # and a name that needed one would have to be quoted. Rejecting it here also keeps
        # the sheet-name pattern unable to consume the "!" separator and backtrack over it.
        "!!A1",
        "Sheet!1!A1",
        "[Book1.xlsx]Sheet!1!A1",
    ],
)
def test_parse_caller_address_malformed(caller_address):
    with pytest.raises(xw.XlwingsError):
        parse_caller_address(caller_address)


@pytest.mark.parametrize(
    "address",
    [
        # Each drives a different quantifier in _CALLER_ADDRESS_RE with input that can never
        # complete a match, the shape that makes a backtracking regex blow up. These assert
        # only that pathological input is still rejected as malformed, not how fast: see
        # test_caller_address_re_has_no_overlapping_quantifiers for the actual ReDoS guard.
        "[" + "\\" * 20_000,
        "'" + "a''" * 20_000 + "!A1x",
        "S" * 20_000 + "!A1!",
        "&" * 20_000,
    ],
    ids=[
        "unclosed_book",
        "unclosed_quoted_sheet",
        "trailing_separator",
        "no_separator",
    ],
)
def test_parse_caller_address_rejects_pathological_input(address):
    with pytest.raises(xw.XlwingsError):
        parse_caller_address(address)


def test_caller_address_re_has_no_overlapping_quantifiers():
    # Guards against a future rewrite reintroducing the constructs CodeQL flagged as
    # "polynomial regular expression on uncontrolled data" - the caller address comes from
    # the Excel client, so match time must stay linear in its length.
    #
    # This is a structural assertion rather than a timing one on purpose. A wall-clock test
    # can't do this job: measured against both this pattern and the pre-fix one, match time
    # grows ~8x for 8x longer input (linear) in every case above, so no timing threshold
    # separates them - it would only ever fail spuriously on a loaded CI runner.
    pattern = _CALLER_ADDRESS_RE.pattern
    # (?:[^']|'')* lets the engine consume the same text two ways. The unrolled equivalent,
    # [^']*(?:''[^']*)*, has exactly one path through any input.
    assert "|'')*" not in pattern, "overlapping alternation in the quoted-sheet body"
    # A lazy sheet name that can also match "!" backtracks over the separator on failure.
    assert "*?" not in pattern, "lazy quantifier in the sheet name"
    assert re.search(
        r"\(\?P<sheet>\[\^[^\]]*!", pattern
    ), "the unquoted sheet name must exclude '!' so it can't consume the separator"


def test_caller_from_address():
    caller = caller_from_address("[Book1.xlsx]Sheet1!B21")
    assert isinstance(caller, Caller)
    assert caller.address == "B21"
    assert caller.row == 21
    assert caller.column == 2
    assert caller.sheet_name == "Sheet1"
    assert caller.book_name == "Book1.xlsx"


def test_caller_from_address_multi_cell():
    """Office.js documents invocation.address as a single cell, so this shouldn't occur.

    Parsed tolerantly anyway rather than rejected: row/column report the top-left cell, so
    a range (if one ever arrives, e.g. from a legacy CSE array formula) degrades to
    something sensible instead of failing the call.
    """
    caller = caller_from_address("Sheet1!B21:D25")
    assert caller.address == "B21:D25"
    assert (caller.row, caller.column) == (21, 2)


@pytest.mark.parametrize(
    "caller_address, expected",
    [
        ("Sheet1!$B$21", "B21"),
        ("Sheet1!$B21", "B21"),
        ("Sheet1!B$21", "B21"),
        ("Sheet1!$B$21:$D$25", "B21:D25"),
        ("[Book1.xlsx]Sheet1!$B$21", "B21"),
    ],
)
def test_caller_address_is_normalized_without_dollar_signs(caller_address, expected):
    """Excel may send any mix of absolute/relative markers; the output is always relative.

    address is rebuilt from the parsed row/column rather than echoed back, following the
    modern ($-less) convention used by Office Scripts.
    """
    assert caller_from_address(caller_address).address == expected


def test_caller_from_address_quoted_sheet():
    assert caller_from_address("[B.xlsx]'My Sheet'!A1").sheet_name == "My Sheet"
    assert caller_from_address("'It''s'!$A$1:$C$3").sheet_name == "It's"


def test_caller_from_address_without_workbook_prefix():
    # Older clients and most of the test suite send a bare Sheet1!B21. Report the missing
    # workbook as "" rather than inventing a plausible-looking name.
    assert caller_from_address("Sheet1!B21").book_name == ""


@pytest.mark.parametrize("caller_address", [None, ""])
def test_caller_from_address_unavailable(caller_address):
    assert caller_from_address(caller_address) is None


@pytest.mark.parametrize("caller_address", [0, [], {}, 5])
def test_caller_from_address_non_string(caller_address):
    # Not "unavailable": caller_address arrives in an untyped dict, so a non-string is
    # malformed input and must surface rather than silently become None.
    with pytest.raises(xw.XlwingsError):
        caller_from_address(caller_address)


@pytest.mark.parametrize(
    "caller_address",
    [
        "Sheet1!A0",  # row 0
        "Sheet1!XFE1",  # past column XFD: col_name() raises IndexError
        "Sheet1!A1048577",  # past the last row
        "Sheet1!B2:A1",  # reversed range
        "Sheet1!A1junk",  # trailing characters
        # A regex anchored with "$" would also match just before a trailing newline, which
        # both bypasses the trailing-character check and puts a newline into the log line.
        "Sheet1!A1\n",
        "Sheet1!A1\r\n",
        "\nSheet1!A1",
        "She\net1!A1",
        "'My\nSheet'!A1",
    ],
)
def test_caller_from_address_invalid_raises_xlwings_error(caller_address):
    # The type matters as much as the raise: xlwings' helpers leak IndexError, which isn't
    # an XlwingsError and would escape the routers' except clause as a 500.
    with pytest.raises(xw.XlwingsError):
        caller_from_address(caller_address)


@pytest.mark.parametrize("caller_address", ["Sheet1!A1", "Sheet1!XFD1048576"])
def test_caller_from_address_boundaries(caller_address):
    """Excel's first and last cell must still parse."""
    assert caller_from_address(caller_address) is not None


def _make_module(name, source):
    module = ModuleType(name)
    sys.modules[name] = module
    exec(compile(source, f"{name}.py", "exec"), module.__dict__)
    return module


_FUNC_HEADER = "from xlwings import Caller\nfrom xlwings.server import func\n"


@pytest.mark.parametrize("hint", ["Caller", "Caller | None"])
def test_streaming_function_with_caller_is_rejected(hint):
    """A streaming function can never learn its calling cell, so asking for it is a bug.

    Office.js sets either "stream" or "requiresAddress", never both. Failing here - at
    registration - beats injecting None at call time, which would silently push a
    permanently-dead branch into every user's function body.
    """
    module = _make_module(
        f"streaming_caller_{hint.replace(' | ', '_')}",
        _FUNC_HEADER + f"@func\nasync def streamed(caller: {hint}):\n    yield 1\n",
    )
    with pytest.raises(xw.XlwingsError, match="Caller type hint"):
        xw.server.custom_functions_meta(module)


def test_regular_function_may_use_a_bare_caller_hint():
    """`caller: Caller` needs no `| None`: a regular function always gets an address."""
    module = _make_module(
        "regular_bare_caller",
        _FUNC_HEADER
        + "@func\ndef regular(caller: Caller):\n    return caller.address\n",
    )
    meta = xw.server.custom_functions_meta(module)
    (function,) = meta["functions"]
    assert function["options"]["requiresAddress"] is True
    # Still framework-injected, so Excel must not see it in the signature.
    assert function["parameters"] == []


@pytest.mark.parametrize(
    "hint, caller, expected",
    [
        ("Caller", "Sheet1!B21", "B21"),
        ("Caller | None", "Sheet1!B21", "B21"),
        ("Caller | None", None, "no caller"),
    ],
)
def test_caller_injection(hint, caller, expected):
    module = _make_module(
        f"inject_{hint.replace(' | ', '_')}_{caller}",
        _FUNC_HEADER + f"@func\ndef f(caller: {hint}):\n"
        "    return caller.address if caller else 'no caller'\n",
    )
    data = {"func_name": "f", "args": [], "version": xw.__version__, "runtime": None}
    rv = asyncio.run(
        xw.server.custom_functions_call(
            data, module=module, typehint_to_value={Caller: caller_from_address(caller)}
        )
    )
    assert rv == [[expected]]


def test_bare_caller_hint_errors_when_the_address_is_unusable():
    """A bare `Caller` promises a value, so a None must not reach the function body.

    The routers degrade a malformed address to None so it can't 500; that None then has to
    surface as a clear error rather than an AttributeError on `caller.address`.
    """
    module = _make_module(
        "inject_bare_none",
        _FUNC_HEADER + "@func\ndef f(caller: Caller):\n    return caller.address\n",
    )
    data = {"func_name": "f", "args": [], "version": xw.__version__, "runtime": None}
    with pytest.raises(xw.XlwingsError, match="Could not determine the calling cell"):
        asyncio.run(
            xw.server.custom_functions_call(
                data, module=module, typehint_to_value={Caller: None}
            )
        )


def test_streaming_function_without_caller_is_unaffected():
    module = _make_module(
        "streaming_no_caller",
        _FUNC_HEADER + "@func\nasync def streamed(x):\n    yield 1\n",
    )
    meta = xw.server.custom_functions_meta(module)
    assert meta["functions"][0]["options"] == {"stream": True}


def test_caller_is_exported_on_the_top_level_namespace():
    """xw.Caller is the documented import, next to xw.WithScript and friends."""
    assert xw.Caller is Caller
    assert "Caller" in xw.__all__


def test_every_name_in_all_is_importable():
    """`from xlwings import *` must never raise AttributeError.

    Caller is license-gated (it lives under pro/), so it's only appended to __all__ once
    the import succeeded. Advertising a name that isn't bound breaks star imports for
    everyone without a valid PRO license, which no other __all__ entry does.
    """
    unresolvable = [name for name in xw.__all__ if not hasattr(xw, name)]
    assert unresolvable == []


def test_star_import_works_without_a_valid_pro_license():
    """The license is validated at import time, so this needs a fresh interpreter."""
    env = {**os.environ, "XLWINGS_LICENSE_KEY": "invalid"}
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            "from xlwings import *\n"
            "import xlwings as xw\n"
            "assert not hasattr(xw, 'Caller'), 'expected the PRO import to fail'\n"
            "assert 'Caller' not in xw.__all__\n"
            "print('ok')",
        ],
        capture_output=True,
        text=True,
        env=env,
    )
    assert result.returncode == 0, result.stderr
    assert "ok" in result.stdout


def test_caller_is_immutable():
    caller = caller_from_address("Sheet1!B21")
    with pytest.raises(dataclasses.FrozenInstanceError):
        caller.address = "A1"


def test_caller_does_not_touch_global_xlwings_state():
    """Callers must not create workbooks in the process-global remote engine.

    An earlier design built a synthetic xw.Book per invocation, which accumulated books
    without bound and reassigned the global active book - a cross-request hazard. This is
    the canary for that regression.
    """
    books = xw.engines["remote"].apps.active.books
    count_before = len(books)
    active_before = books.active

    for caller_address in [
        "[Book1.xlsx]Sheet1!B21",
        "Sheet2!A1:C3",
        "'My Sheet'!$D$4",
    ]:
        caller_from_address(caller_address)

    assert len(books) == count_before
    # Compare by identity: Book.__eq__ only compares app and name, so two distinct books
    # with the same name would compare equal and hide a replaced active book.
    assert books.active.impl is active_before.impl
