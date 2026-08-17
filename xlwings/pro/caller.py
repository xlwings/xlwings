"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org

The cell that called a custom function, delivered via the Caller type hint.

Excel sends the calling cell as an address string (e.g. "[Book1.xlsx]Sheet1!B21"), not as
workbook data: a custom function receives its arguments and nothing else. Caller therefore
models the *reference* only - it deliberately isn't an xw.Range, which would advertise cell
values, reads and writes that this transport can't back. Use xw.WithScript to run a custom
script when you need to touch the workbook itself.

This lives in xlwings core rather than in a framework on top of it so that every
custom-functions runtime can reach it: xlwings Server imports it, and xlwings Lite (which
runs in Pyodide and only ever ships xlwings) would otherwise have no way to get at it.
"""

from __future__ import annotations

import dataclasses
import re
from typing import Optional, Tuple

from ..constants import MAX_COLUMNS, MAX_ROWS
from ..utils import a1_to_tuples, col_name
from .udfs_officejs import register_injectable_typehint

# [Book1.xlsx]Sheet1!B21 | Sheet1!B21 | [Book1.xlsx]'My Sheet'!A1 | 'It''s'!$A$1:$C$3
# The workbook prefix is optional: Office.js always sends it, but older clients don't.
# A quoted sheet name escapes a literal apostrophe by doubling it.
# Both regexes are applied with fullmatch(): "$" would also match just before a trailing
# newline, which would let "Sheet1!A1\n" through.
# Every quantifier below is greedy over a character class that can't also be matched by
# what follows it, so each character is consumed exactly one way and matching stays linear:
# the quoted-sheet body is written as runs of non-apostrophes separated by doubled
# apostrophes rather than the equivalent (?:[^']|'')*, whose two branches overlap, and the
# unquoted sheet name excludes "!" so it can't consume the separator and backtrack over it.
# Excluding "!" also rejects "!!A1" (sheet "!"), which the "!"-permitting version accepted:
# Excel has no "!" in sheet names, and a name needing one would have to be quoted anyway.
_CALLER_ADDRESS_RE = re.compile(
    r"(?:\[(?P<book>[^\]\n\r]*)\])?"
    r"(?:'(?P<quoted_sheet>[^'\n\r]*(?:''[^'\n\r]*)*)'|(?P<sheet>[^'!\n\r]*))"
    r"!(?P<address>[^!\n\r]+)"
)

# a1_to_tuples() accepts anything its regex matches and ignores the rest, so "A1junk"
# silently becomes A1. Match the address strictly and reject leftovers ourselves.
_A1_RE = re.compile(r"\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?", re.I)


@dataclasses.dataclass(frozen=True)
class Caller:
    """The cell that called a custom function (xlwings Lite and xlwings Server only).

    Carries the reference only - there are no cell values here, as custom functions never
    receive the workbook. Streaming functions aren't supported.
    """

    address: str
    """Normalized relative A1 notation, e.g. "B21" or "B21:D25".

    Excel's own wire format varies (``B21``, ``$B$21``, with or without the workbook
    prefix), so this is always rebuilt from the parsed row/column rather than echoed back.
    The ``$``-less form follows the modern convention used by Office Scripts.
    """

    row: int
    """1-based row of the top-left cell."""

    column: int
    """1-based column of the top-left cell."""

    sheet_name: str
    """Name of the sheet holding the calling cell."""

    book_name: str
    """Name of the workbook, or "" if the client didn't send one."""


def parse_caller_address(caller_address: str) -> Tuple[str, str, str]:
    """Split "[Book1.xlsx]'My Sheet'!$A$1:$C$3" into (book_name, sheet_name, address).

    The workbook name is "" when the client doesn't send the optional prefix.
    """
    # Imported here rather than at module level: xlwings/__init__.py imports the pro
    # subpackage, so a top-level `import xlwings` would be circular.
    from .. import XlwingsError

    if not isinstance(caller_address, str):
        raise XlwingsError(f"Invalid caller address: {caller_address!r}")
    match = _CALLER_ADDRESS_RE.fullmatch(caller_address)
    if not match:
        raise XlwingsError(f"Invalid caller address: '{caller_address}'")
    quoted_sheet = match.group("quoted_sheet")
    if quoted_sheet is not None:
        sheet_name = quoted_sheet.replace("''", "'")
    else:
        sheet_name = match.group("sheet")
    if not sheet_name:
        raise XlwingsError(f"Invalid caller address: '{caller_address}'")
    return match.group("book") or "", sheet_name, match.group("address")


def caller_from_address(caller_address: Optional[str]) -> Optional[Caller]:
    """Build a Caller from the address Excel sent, or None if it didn't send one.

    Returns None only for None and "", the two ways a client signals "no address".
    Anything else that doesn't parse into a valid Excel reference raises an XlwingsError,
    so that malformed input surfaces in the logs rather than being mistaken for absence.
    """
    from .. import XlwingsError

    if caller_address is None or caller_address == "":
        return None
    book_name, sheet_name, address = parse_caller_address(caller_address)
    if not _A1_RE.fullmatch(address):
        raise XlwingsError(f"Invalid caller address: '{caller_address}'")
    try:
        first, last = a1_to_tuples(address.replace("$", "").upper())
    except (IndexError, TypeError, ValueError) as e:
        raise XlwingsError(f"Invalid caller address: '{caller_address}'") from e
    if first is None:
        raise XlwingsError(f"Invalid caller address: '{caller_address}'")
    row, column = first
    last_row, last_column = last if last else first
    # a1_to_tuples() doesn't range-check, so "A0" and "XFE1" get this far. Validate before
    # col_name(), which raises IndexError (not an XlwingsError) beyond column XFD.
    if not (1 <= row <= last_row <= MAX_ROWS):
        raise XlwingsError(f"Caller address is out of range: '{caller_address}'")
    if not (1 <= column <= last_column <= MAX_COLUMNS):
        raise XlwingsError(f"Caller address is out of range: '{caller_address}'")
    try:
        normalized = f"{col_name(column)}{row}"
        if last:
            normalized += f":{col_name(last_column)}{last_row}"
    except (IndexError, TypeError, ValueError) as e:
        raise XlwingsError(f"Invalid caller address: '{caller_address}'") from e
    return Caller(
        address=normalized,
        row=row,
        column=column,
        sheet_name=sheet_name,
        book_name=book_name,
    )


# Tell xlwings that Caller is provided by the framework rather than by Excel, so that it's
# hidden from the Excel-facing signature and can be placed in any position, including
# keyword-only, i.e. after *args. Registered in the defining module so that importing
# Caller from anywhere can't bypass it.
register_injectable_typehint(Caller)
