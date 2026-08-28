"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import asyncio
import base64
import datetime as dt
import numbers
import re
import sys
from functools import lru_cache

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

from .. import NoSuchObjectError, XlwingsError, __version__, base_classes, utils
from ..constants import MAX_COLUMNS, MAX_ROWS

# Private marker set on a sheet's api dict once its cell values have been loaded
# on an async (lazy) book. Stored on the dict (not keyed by sheet name) so it
# survives renames; the sheet's api dict is preserved across metadata reloads by
# _update_api_in_place. Only read locally (never serialized back to JS).
_SHEET_VALUES_LOADED_KEY = "_xlwings_values_loaded"

# xlwings' public calculation modes mapped to Office.js' Excel.CalculationMode.
# xlwings' "semiautomatic" is Office.js' "AutomaticExceptTables".
_CALCULATION_PY2JS = {
    "automatic": "Automatic",
    "manual": "Manual",
    "semiautomatic": "AutomaticExceptTables",
}
_CALCULATION_JS2PY = {v: k for k, v in _CALCULATION_PY2JS.items()}


def _mark_sheet_values_loaded(sheet_api):
    sheet_api[_SHEET_VALUES_LOADED_KEY] = True


def _sheet_values_loaded(sheet_api):
    return sheet_api.get(_SHEET_VALUES_LOADED_KEY, False)


def _normalize_jsnull(obj):
    """Recursively replace Pyodide's `JsNull` sentinel with Python `None`.

    Pyodide >= 0.28 converts JS `null` to `pyodide.ffi.jsnull` instead of
    `None`. Book data coming back from Office.js via `.to_py()` therefore
    contains `JsNull` for empty/absent fields (e.g. `scope_sheet_index` on
    book-scoped names, `address` for non-cell selections), which breaks the
    `is None` checks throughout this module. Normalize everything back to
    `None` at the JS boundary so downstream code stays Pyodide-version
    agnostic.

    The `values` arrays (cell data) are skipped here for two reasons. First,
    *book* data (`getBookData()`) represents empty cells as `""`, not
    `null`, so a book's `values` cannot contain `JsNull`. Second, walking
    them would mean an extra full pass over every cell of an eagerly-loaded book
    (e.g. `xw.Book(json=...)` in xlwings Lite's notebook runner).

    Note this is *not* true for custom function *arguments*: Excel's custom
    functions runtime sends empty cells in a range argument as JS `null` ->
    `JsNull`. Those don't pass through here — UDFs use a separate engine, so
    they're normalized in `_xlofficejs._clean_value_data_element` instead.
    """
    try:
        from pyodide.ffi import JsNull
    except ImportError:
        return obj

    def walk(o):
        if isinstance(o, JsNull):
            return None
        if isinstance(o, dict):
            return {k: v if k == "values" else walk(v) for k, v in o.items()}
        if isinstance(o, list):
            return [walk(v) for v in o]
        return o

    return walk(obj)


# Time types (doesn't contain dt.date)
time_types = (dt.datetime,)
if np:
    time_types = time_types + (np.datetime64,)
if pd:
    time_types = time_types + (pd.Timestamp,)

datetime_pattern = r"^(-?(?:[1-9][0-9]*)?[0-9]{4})-(1[0-2]|0[1-9])-(3[01]|0[1-9]|[12][0-9])T(2[0-3]|[01][0-9]):([0-5][0-9]):([0-5][0-9])(\.[0-9]+)?(Z|[+-](?:2[0-3]|[01][0-9]):[0-5][0-9])?$"  # noqa: E501
datetime_regex = re.compile(datetime_pattern)


def _update_api_in_place(target, source):
    """Update target dict in-place from source, preserving references to nested dicts
    inside lists (matched by 'name' key). This ensures that e.g. Sheet or Name objects
    holding references to dicts inside target['sheets'] see the updated values."""
    for key, value in source.items():
        if isinstance(value, list) and key in target and isinstance(target[key], list):
            old_by_name = {
                item["name"]: item
                for item in target[key]
                if isinstance(item, dict) and "name" in item
            }
            new_list = []
            for item in value:
                if isinstance(item, dict) and item.get("name") in old_by_name:
                    old_by_name[item["name"]].update(item)
                    new_list.append(old_by_name[item["name"]])
                else:
                    new_list.append(item)
            target[key] = new_list
        elif (
            isinstance(value, dict) and key in target and isinstance(target[key], dict)
        ):
            target[key].update(value)
        else:
            target[key] = value


def _clean_value_data_element(
    value, datetime_builder, empty_as, number_builder, err_to_str
):
    if value == "":
        return empty_as
    if isinstance(value, str):
        # TODO: Send arrays back and forth with indices of the location of dt values
        if datetime_regex.match(value):
            value = dt.datetime.fromisoformat(
                value[:-1]
            )  # cut off "Z" (Python doesn't accept it and Excel doesn't support tz)
        elif not err_to_str and value in [
            "#DIV/0!",
            "#N/A",
            "#NAME?",
            "#NULL!",
            "#NUM!",
            "#REF!",
            "#VALUE!",
            "#DATA!",
        ]:
            value = None
        else:
            value = value
    if isinstance(value, dt.datetime) and datetime_builder is not dt.datetime:
        value = datetime_builder(
            month=value.month,
            day=value.day,
            year=value.year,
            hour=value.hour,
            minute=value.minute,
            second=value.second,
            microsecond=value.microsecond,
            tzinfo=None,
        )
    elif number_builder is not None and isinstance(value, float):
        value = number_builder(value)
    return value


class Engine:
    def __init__(self):
        self.apps = Apps()

    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder, err_to_str):
        return [
            [
                _clean_value_data_element(
                    c, datetime_builder, empty_as, number_builder, err_to_str
                )
                for c in row
            ]
            for row in data
        ]

    @staticmethod
    def prepare_xl_data_element(x, options):
        if x is None:
            return ""
        elif pd and pd.isna(x):
            return ""
        elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif np and isinstance(x, np.number):
            return float(x)
        elif pd and isinstance(x, type(pd.NaT)):
            return None
        elif isinstance(x, time_types):
            if np and isinstance(x, np.datetime64):
                x = utils.np_datetime_to_datetime(x)
            elif pd and isinstance(x, pd.Timestamp):
                x = x.to_pydatetime()
            if x.time() == dt.time(0, 0):
                x = x.date().isoformat()
            else:
                x = x.replace(tzinfo=None).isoformat(sep=" ").split(".")[0]
        elif isinstance(x, dt.date):
            x = x.isoformat()
        return x

    @property
    def name(self):
        return "remote"

    @property
    def type(self):
        return "remote"


class Apps(base_classes.Apps):
    def __init__(self):
        self._apps = [App(self)]

    def __iter__(self):
        return iter(self._apps)

    def __len__(self):
        return len(self._apps)

    def __getitem__(self, index):
        return self._apps[index]

    def add(self, **kwargs):
        self._apps.insert(0, App(self, **kwargs))
        return self._apps[0]


class App(base_classes.App):
    _next_pid = -1

    def __init__(self, apps, add_book=True, **kwargs):
        self.apps = apps
        self._pid = App._next_pid
        App._next_pid -= 1
        self._display_alerts = True
        self._books = Books(self)
        if add_book:
            self._books.add()

    def kill(self):
        self.apps._apps.remove(self)
        self.apps = None

    @property
    def engine(self):
        return engine

    @property
    def books(self):
        return self._books

    @property
    def pid(self):
        return self._pid

    @property
    def display_alerts(self):
        # Office.js never shows the alerts that this suppresses on the desktop
        # engines, but Range.merge() toggles it, so it has to be readable and
        # writable rather than raising.
        return self._display_alerts

    @display_alerts.setter
    def display_alerts(self, value):
        self._display_alerts = value

    def _unsupported(name, detail, read_only=False):
        # Office.js' Excel.Application doesn't expose these, so they raise with
        # the reason rather than a bare NotImplementedError. read_only mirrors
        # the public API: path, startup_path and version have no setter there,
        # so defining one here would turn AttributeError into the wrong error.
        message = f"App.{name} is not supported in Office.js: {detail}"

        def getter(self):
            raise NotImplementedError(message)

        if read_only:
            return property(getter)

        def setter(self, value):
            raise NotImplementedError(message)

        return property(getter, setter)

    cut_copy_mode = _unsupported("cut_copy_mode", "it has no clipboard API.")
    enable_events = _unsupported(
        "enable_events", "Excel.Application has no equivalent property."
    )
    interactive = _unsupported(
        "interactive", "Excel.Application has no equivalent property."
    )
    status_bar = _unsupported(
        "status_bar", "Excel.Application has no equivalent property."
    )
    path = _unsupported(
        "path",
        "an add-in has no access to the Excel installation's paths.",
        read_only=True,
    )
    startup_path = _unsupported(
        "startup_path",
        "an add-in has no access to the Excel installation's paths.",
        read_only=True,
    )
    version = _unsupported(
        "version",
        "Excel.Application only exposes calculationEngineVersion, which is not "
        "the application version.",
        read_only=True,
    )
    del _unsupported

    def quit(self):
        raise NotImplementedError(
            "App.quit() is not supported in Office.js: an add-in can't close the "
            "Excel application."
        )

    @property
    def calculation(self):
        calculation = self.books.active.api["book"].get("calculation")
        if calculation is None:
            raise NotImplementedError(
                "This client doesn't send the calculation mode. It requires a "
                "newer version of the xlwings JavaScript module."
            )
        return _CALCULATION_JS2PY[calculation]

    @calculation.setter
    def calculation(self, value):
        try:
            js_value = _CALCULATION_PY2JS[value]
        except KeyError:
            raise ValueError(
                f"Invalid calculation mode: {value!r}. Must be one of "
                f"{sorted(_CALCULATION_PY2JS)}."
            ) from None
        self.books.active.api["book"]["calculation"] = js_value
        self.books.active.append_json_action(func="setCalculation", args=[js_value])

    @property
    def screen_updating(self):
        # Office.js has no screen updating flag to read back, only a
        # suspend-until-next-sync call. See the setter.
        raise NotImplementedError(
            "App.screen_updating can't be read in Office.js, which has no screen "
            "updating property, only suspendScreenUpdatingUntilNextSync()."
        )

    @screen_updating.setter
    def screen_updating(self, value):
        # Office.js can only suspend screen updating until its next sync, not
        # turn it off indefinitely. Setting it back to True is therefore a
        # no-op: the suspension ends on its own at the next sync.
        self.books.active.append_json_action(
            func="setScreenUpdating", args=[bool(value)]
        )

    def calculate(self):
        # args=[] rather than omitting it: append_json_action wraps a missing
        # args as [None], and this action takes no arguments.
        self.books.active.append_json_action(func="calculate", args=[])

    @property
    def selection(self):
        book = self.books.active
        return Range(sheet=book.sheets.active, arg1=book.api["book"]["selection"])

    async def get_selection(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "App.get_selection() is only supported in xlwings Lite"
            )
        import js

        result = _normalize_jsnull((await js.xlwings.getSelection()).to_py())
        sheet_index = int(result["sheetIndex"])
        address = result["address"]
        book = self.books.active
        sheet = Sheet(
            api=book.api["sheets"][sheet_index],
            sheets=book.sheets,
            index=sheet_index + 1,
        )
        if address is None:
            return None  # Non-cell selection (e.g., shape)
        return Range(sheet=sheet, arg1=address)

    @property
    def visible(self):
        return True

    @visible.setter
    def visible(self, value):
        pass

    def activate(self, steal_focus=None):
        pass

    def alert(self, prompt, title, buttons, mode, callback):
        self.books.active.append_json_action(
            func="alert",
            args=[
                "" if prompt is None else prompt,
                "" if title is None else title,
                "" if buttons is None else buttons,
                "" if mode is None else mode,
                "" if callback is None else callback,
            ],
        )

    def run(self, macro, args):
        self.books.active.append_json_action(
            func="runMacro",
            args=[macro] + [args] if not isinstance(args, list) else [macro] + args,
        )


class Books(base_classes.Books):
    def __init__(self, app):
        self.app = app
        self.books = []
        self._active = None

    @property
    def active(self):
        return self._active

    async def get_active(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Books.get_active() is only supported in xlwings Lite"
            )
        import js
        from pyodide.ffi import to_js

        book_data_js = await js.xlwings.getBookData(
            to_js({"lazy": True}, dict_converter=js.Object.fromEntries)
        )
        # open() normalizes JsNull -> None
        book_data = book_data_js.to_py()
        # The book was fetched lazily (structure only, no cell values), so mark
        # it: sync `.value` reads raise until values are loaded, and `load()`
        # defaults to metadata-only. See Book.load / Range.api.
        return self.open(book_data, lazy=True)

    def open(self, json, lazy=False):
        # Normalize here (rather than only at the getBookData boundary) so that
        # callers passing raw `.to_py()` data straight to `xw.Book(json=...)`
        # (e.g. xlwings Lite's notebook runner) also get JsNull -> None.
        book = Book(api=_normalize_jsnull(json), books=self, lazy=lazy)
        self.books.append(book)
        self._active = book
        return book

    def add(self):
        book = Book(
            api={
                "version": __version__,
                "book": {"name": f"Book{len(self) + 1}", "active_sheet_index": 0},
                "sheets": [
                    {
                        "name": "Sheet1",
                        "values": [[]],
                    },
                ],
            },
            books=self,
        )
        self.books.append(book)
        self._active = book
        return book

    def _try_find_book_by_name(self, name):
        for book in self.books:
            if book.name == name or book.fullname == name:
                return book
        return None

    def __len__(self):
        return len(self.books)

    def __iter__(self):
        for book in self.books:
            yield book

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return self.books[name_or_index - 1]
        else:
            book = self._try_find_book_by_name(name_or_index)
            if book is None:
                raise KeyError(name_or_index)
            return book


class Book(base_classes.Book):
    def __init__(self, api, books, lazy=False):
        self.books = books
        self._api = api
        self._json = {"actions": []}
        # Async/lazy book: fetched with structure only (no cell values). While
        # lazy, sync `.value` reads on a sheet whose values haven't been loaded
        # raise (use `await get_value()` or `await load(values=True)`), and
        # `load()` defaults to metadata-only. Whether a sheet's values are loaded
        # is marked per sheet on its api dict (see `_SHEET_VALUES_LOADED_KEY`), so
        # it survives sheet renames.
        self._lazy = lazy
        if api["version"] != __version__ and api["client"] != "Office.js":
            raise XlwingsError(
                f"Your xlwings version is different on the client ({api['version']}) "
                f"and server ({__version__})."
            )

    def append_json_action(self, **kwargs):
        args = kwargs.get("args")
        self._json["actions"].append(
            {
                "func": kwargs.get("func"),
                "args": [args] if not isinstance(args, list) else args,
                "values": kwargs.get("values"),
                "sheet_position": kwargs.get("sheet_position"),
                "start_row": kwargs.get("start_row"),
                "start_column": kwargs.get("start_column"),
                "row_count": kwargs.get("row_count"),
                "column_count": kwargs.get("column_count"),
            }
        )

    @property
    def api(self):
        return self._api

    def json(self):
        return self._json

    async def flush(self):
        if sys.platform != "emscripten":
            raise NotImplementedError("Book.flush() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        actions = self._json.get("actions", [])
        if actions:
            actions_js = to_js(
                {"actions": actions}, dict_converter=js.Object.fromEntries
            )
            await js.xlwings.runActions(actions_js)
            self._json["actions"] = []
        # Yield to the browser event loop so it can repaint (to print to output pane)
        await asyncio.sleep(0.01)

    async def load(self, values=None):
        """(Re)load the book's data from Excel on demand.

        On an async (lazy) book, this defaults to loading only *metadata* -
        sheet structure, tables, pictures, names - and not cell values, since
        bulk-loading values would defeat the point of the async API. Pass
        `values=True` to also snapshot all cell values (after which sync
        `.value` reads work again).

        On a regular (eager) book, everything including values is loaded, as
        before.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError("Book.load() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        # Default: metadata-only for lazy books, full for eager books.
        load_values = (not self._lazy) if values is None else bool(values)
        # getBookData(lazy=True) returns structure only (empty values); a plain
        # call returns everything.
        opts = to_js({"lazy": not load_values}, dict_converter=js.Object.fromEntries)
        data = _normalize_jsnull((await js.xlwings.getBookData(opts)).to_py())
        if not load_values:
            # Metadata-only: don't let the empty `values` payload clobber any
            # values already loaded on this book.
            for sheet in data.get("sheets", []):
                sheet.pop("values", None)
        _update_api_in_place(self._api, data)
        if load_values:
            for sheet in self._api.get("sheets", []):
                _mark_sheet_values_loaded(sheet)
        get_range_api.cache_clear()

    @property
    def name(self):
        return self.api["book"]["name"]

    @property
    def fullname(self):
        return self.name

    @property
    def names(self):
        return Names(parent=self, api=self.api["names"])

    @property
    def sheets(self):
        return Sheets(api=self.api["sheets"], book=self)

    @property
    def app(self):
        return self.books.app

    def save(self, path=None, password=None):
        if path is not None:
            raise NotImplementedError(
                "Book.save() can't take a path in Office.js, which has no SaveAs "
                "equivalent: Workbook.save() only saves in place. Call save() "
                "without arguments instead."
            )
        if password is not None:
            raise NotImplementedError(
                "Book.save() can't take a password in Office.js, which has no API "
                "for setting one."
            )
        self.append_json_action(func="save", args=[])

    def to_pdf(self, path, quality):
        raise NotImplementedError(
            "Book.to_pdf() is not supported in Office.js, which has no PDF export "
            "API."
        )

    def close(self):
        assert self.api is not None, "Seems this book was already closed."
        self.books.books.remove(self)
        self.books = None
        self._api = None

    def activate(self):
        pass


class Sheets(base_classes.Sheets):
    def __init__(self, api, book):
        self._api = api
        self.book = book

    @property
    def active(self):
        ix = self.book.api["book"]["active_sheet_index"]
        return Sheet(api=self.api[ix], sheets=self, index=ix + 1)

    async def get_active(self):
        if sys.platform != "emscripten":
            raise NotImplementedError(
                "Sheets.get_active() is only supported in xlwings Lite"
            )
        import js

        ix = int(await js.xlwings.getActiveSheetIndex())
        return Sheet(api=self.api[ix], sheets=self, index=ix + 1)

    @property
    def api(self):
        return self._api

    def __call__(self, name_or_index):
        if isinstance(name_or_index, int):
            return Sheet(
                api=self.api[name_or_index - 1], sheets=self, index=name_or_index
            )
        else:
            for ix, sheet in enumerate(self.api):
                if sheet["name"] == name_or_index:
                    return Sheet(api=sheet, sheets=self, index=ix + 1)
        raise ValueError(f"Sheet '{name_or_index}' doesn't exist!")

    def add(self, before=None, after=None, name=None):
        # TODO: this is hardcoded to English
        if name is None:
            sheet_number = 1
            while True:
                if f"Sheet{sheet_number}" in [sheet.name for sheet in self]:
                    sheet_number += 1
                else:
                    break
            name = f"Sheet{sheet_number}"

        api = {
            "name": name,
            "visibility": "Visible",
            "values": [[]],
            "pictures": [],
            "tables": [],
        }

        if before:
            if before.index == 1:
                ix = 1
            else:
                ix = before.index - 1
        elif after:
            ix = after.index + 1
        else:
            # Default position is different from Desktop apps!
            ix = len(self) + 1

        self.api.insert(ix - 1, api)
        self.book.append_json_action(func="addSheet", args=[ix - 1, name])
        self.book.api["book"]["active_sheet_index"] = ix - 1

        return Sheet(api=api, sheets=self, index=ix)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for ix, sheet in enumerate(self.api):
            yield Sheet(api=sheet, sheets=self, index=ix + 1)


class Sheet(base_classes.Sheet):
    def __init__(self, api, sheets, index):
        self._api = api
        self._index = index
        self.sheets = sheets

    def append_json_action(self, **kwargs):
        self.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.append_json_action(
            func="setSheetName",
            args=value,
        )
        self.api["name"] = value

    @property
    def visible(self):
        # Office.js also knows "VeryHidden", which maps to False like "Hidden"
        # does: xlwings' public API is a bool, so the distinction is dropped.
        return self.api["visibility"] == "Visible"

    @visible.setter
    def visible(self, value):
        visibility = "Visible" if value else "Hidden"
        self.append_json_action(
            func="setSheetVisibility",
            args=visibility,
        )
        self.api["visibility"] = visibility

    @property
    def index(self):
        return self._index

    @property
    def book(self):
        return self.sheets.book

    def range(self, arg1, arg2=None):
        return Range(sheet=self, arg1=arg1, arg2=arg2)

    @property
    def cells(self):
        return Range(
            sheet=self,
            arg1=(1, 1),
            arg2=(MAX_ROWS, MAX_COLUMNS),
        )

    @property
    def names(self):
        api = [
            name
            for name in self.book.api["names"]
            if name["scope_sheet_index"] is not None
            and name["scope_sheet_index"] + 1 == self.index
            and not name["book_scope"]
        ]
        return Names(parent=self, api=api)

    def activate(self):
        ix = self.index - 1
        self.book.api["book"]["active_sheet_index"] = ix
        self.append_json_action(func="activateSheet", args=ix)

    @property
    def pictures(self):
        return Pictures(self)

    @property
    def tables(self):
        return Tables(parent=self)

    def delete(self):
        ix = self.index - 1
        del self.book.api["sheets"][ix]
        # Keep the locally tracked active sheet pointing at an existing sheet:
        # deleting a sheet at or before it shifts the remaining ones down, and
        # deleting the last sheet would otherwise leave the index out of range.
        book_api = self.book.api["book"]
        active_ix = book_api["active_sheet_index"]
        if active_ix > ix:
            book_api["active_sheet_index"] = active_ix - 1
        elif active_ix == ix:
            book_api["active_sheet_index"] = min(ix, len(self.book.api["sheets"]) - 1)
        self.append_json_action(func="sheetDelete")

    def clear(self):
        self.append_json_action(func="sheetClear")

    def clear_contents(self):
        self.append_json_action(func="sheetClearContents")

    def clear_formats(self):
        self.append_json_action(func="sheetClearFormats")

    @property
    def used_range(self):
        address = self.api.get("used_range_address")
        if address:
            return Range(sheet=self, arg1=address)
        if address is None and "used_range_address" in self.api:
            # The client reported an empty sheet, which has no used range.
            # Excel's COM API reports A1 in that case.
            return Range(sheet=self, arg1=(1, 1))
        # Fallback for clients that don't send `used_range_address` yet: derive
        # the extent from the shape of the values payload. Since those values
        # are anchored at A1, the used range's real top-left corner is lost,
        # i.e. a sheet whose used range is C5:D10 reports A1:D10.
        if self.book._lazy and not _sheet_values_loaded(self.api):
            raise XlwingsError(
                f"Cell values of sheet '{self.name}' haven't been loaded "
                "(async book). Use 'await book.load(values=True)' or "
                "'await sheet.load(values=True)' first."
            )
        values = self.api["values"]
        nrows = len(values)
        ncols = max((len(row) for row in values), default=0)
        if nrows == 0 or ncols == 0:
            # Empty sheet: Excel reports A1 as the used range
            return Range(sheet=self, arg1=(1, 1))
        return Range(sheet=self, arg1=(1, 1), arg2=(nrows, ncols))

    @property
    def freeze_panes(self):
        return FreezePanes(self)

    async def load(self, values=None):
        """(Re)load this sheet's data from Excel on demand.

        Like `Book.load`, this loads only *metadata* (tables, pictures, names,
        structure) by default on an async (lazy) book, and everything including
        values on a regular book. Pass `values=True` to also load this sheet's
        cell values.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError("Sheet.load() is only supported in xlwings Lite")
        import js
        from pyodide.ffi import to_js

        book = self.book
        load_values = (not book._lazy) if values is None else bool(values)
        opts = to_js(
            {"include": self.name, "lazy": not load_values},
            dict_converter=js.Object.fromEntries,
        )
        book_data_js = await js.xlwings.getBookData(opts)
        book_data = _normalize_jsnull(book_data_js.to_py())
        for sheet_data in book_data["sheets"]:
            if sheet_data["name"] == self.name:
                if not load_values:
                    # Don't clobber any values already present with the empty
                    # metadata-only payload.
                    sheet_data.pop("values", None)
                self._api.update(sheet_data)
                break
        if load_values:
            _mark_sheet_values_loaded(self._api)
        get_range_api.cache_clear()


@lru_cache(None)
def get_range_api(api_values, arg1, arg2=None):
    # Keeping this outside of the Range class allows us to cache it across multiple
    # instances of the same range
    if arg2:
        values = [
            row[arg1[1] - 1 : arg2[1]] for row in api_values[arg1[0] - 1 : arg2[0]]
        ]
        if not values:
            # Completely outside the used range
            return [(None,) * (arg2[1] + 1 - arg1[1])] * (arg2[0] + 1 - arg1[0])
        else:
            # Partly outside the used range
            nrows = arg2[0] + 1 - arg1[0]
            ncols = arg2[1] + 1 - arg1[1]
            nrows_actual = len(values)
            ncols_actual = len(values[0])
            delta_rows = nrows - nrows_actual
            delta_cols = ncols - ncols_actual
            if delta_rows != 0:
                values.extend([(None,) * ncols_actual] * delta_rows)
            if delta_cols != 0:
                v = []
                for row in values:
                    v.append(row + (None,) * delta_cols)
                values = v
            return values
    else:
        try:
            values = [(api_values[arg1[0] - 1][arg1[1] - 1],)]
            return values
        except IndexError:
            # Outside the used range
            return [(None,)]


class Range(base_classes.Range):
    def __init__(self, sheet, arg1, arg2=None):
        self.sheet = sheet
        self.arg1_input = arg1
        self.arg2_input = arg2

        # Handle None case (for app.selection if e.g., a shape is selected)
        if arg1 is None:
            self.arg1 = None
            self.arg2 = None
            return

        # Range
        if isinstance(arg1, Range) and isinstance(arg2, Range):
            cell1 = arg1.coords[1], arg1.coords[2]
            cell2 = arg2.coords[1], arg2.coords[2]
            arg1 = min(cell1[0], cell2[0]), min(cell1[1], cell2[1])
            arg2 = max(cell1[0], cell2[0]), max(cell1[1], cell2[1])
        # A1 notation
        if isinstance(arg1, str):
            # A1 notation
            tuple1, tuple2 = utils.a1_to_tuples(arg1)
            if not tuple1:
                # Named range
                for api_name in sheet.book.api["names"]:
                    if (
                        api_name["name"].split("!")[-1] == arg1
                        and api_name["sheet_index"] == sheet.index - 1
                    ):
                        tuple1, tuple2 = utils.a1_to_tuples(api_name["address"])
                        break
            if not tuple1:
                # Tables
                for api_table in sheet.api["tables"]:
                    if api_table["name"] == arg1:
                        tuple1, tuple2 = utils.a1_to_tuples(
                            api_table["data_body_range_address"]
                        )
                        break
            if not tuple1:
                raise NoSuchObjectError(
                    f"The address/named range '{arg1}' doesn't exist."
                )
            arg1, arg2 = tuple1, tuple2

        # Coordinates
        if len(arg1) == 4:
            row, col, nrows, ncols = arg1
            arg1 = (row, col)
            if nrows > 1 or ncols > 1:
                arg2 = (row + nrows - 1, col + ncols - 1)

        self.arg1 = arg1  # 1-based tuple
        self.arg2 = arg2  # 1-based tuple
        self.sheet = sheet

    def append_json_action(self, **kwargs):
        # Do nothing if the range is None (e.g., if a Shape is selected)
        if self.arg1 is None:
            return
        nrows, ncols = self.shape
        self.sheet.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.sheet.index - 1,
                    "start_row": self.row - 1,
                    "start_column": self.column - 1,
                    "row_count": nrows,
                    "column_count": ncols,
                },
            }
        )

    @property
    def api(self):
        return get_range_api(
            tuple(tuple(row) for row in self.sheet.api["values"]), self.arg1, self.arg2
        )

    @property
    def coords(self):
        return self.sheet.name, self.row, self.column, len(self.api), len(self.api[0])

    @property
    def row(self):
        return self.arg1[0]

    @property
    def column(self):
        return self.arg1[1]

    @property
    def shape(self):
        if self.arg2:
            return self.arg2[0] - self.arg1[0] + 1, self.arg2[1] - self.arg1[1] + 1
        else:
            return 1, 1

    def get_async_pipeline_overrides(self, options):
        """Return async stage replacements for the converter pipeline."""
        if sys.platform != "emscripten":
            raise NotImplementedError("get_value() is only supported in xlwings Lite")
        from ..conversion.standard import (
            AsyncExpandRangeStage,
            AsyncReadValueFromRangeStage,
            ExpandRangeStage,
            ReadValueFromRangeStage,
        )

        overrides = {ReadValueFromRangeStage: AsyncReadValueFromRangeStage(options)}
        if options.get("expand", None):
            overrides[ExpandRangeStage] = AsyncExpandRangeStage(options)
        return overrides

    @property
    def raw_value(self):
        # On an async (lazy) book, cell values aren't pre-loaded. Reading `.value`
        # synchronously would silently return None; raise instead and point to
        # the async API. `await get_value()` doesn't go through here (it fetches
        # via getRangeValues), and `await book.load(values=True)` marks the sheet
        # as loaded, after which sync reads work again.
        if self.sheet.book._lazy and not _sheet_values_loaded(self.sheet.api):
            raise XlwingsError(
                f"Cell values of sheet '{self.sheet.name}' haven't been loaded "
                "(async book). Use 'await myrange.get_value()' to read on demand, "
                "or 'await book.load(values=True)' to load all values first."
            )
        return self.api

    @raw_value.setter
    def raw_value(self, value):
        if not isinstance(value, list):
            # Covers also this case: myrange['A1:B2'].value = 'xyz'
            nrows, ncols = self.shape
            values = [[value] * ncols] * nrows
        else:
            values = value
        self.append_json_action(
            func="setValues",
            values=values,
        )

    def clear_contents(self):
        self.append_json_action(
            func="rangeClearContents",
        )

    def clear(self):
        self.append_json_action(
            func="rangeClear",
        )

    def clear_formats(self):
        self.append_json_action(
            func="rangeClearFormats",
        )

    @property
    def address(self):
        # Handle non-cell selection
        if self.arg1 is None:
            return
        nrows, ncols = self.shape
        address = f"${utils.col_name(self.column)}${self.row}"
        if nrows == 1 and ncols == 1:
            return address
        else:
            return (
                f"{address}"
                f":${utils.col_name(self.column + ncols - 1)}${self.row + nrows - 1}"
            )

    @property
    def has_array(self):
        # Not supported, but since this is only used for legacy CSE arrays, probably
        # not much of an issue. Here as there's currently a dependency in expansion.py.
        return None

    def end(self, direction):
        if direction == "down":
            i = 1
            while True:
                try:
                    if self.sheet.api["values"][self.row - 1 + i][
                        self.column - 1
                    ] not in (None, ""):
                        i += 1
                    else:
                        break
                except IndexError:
                    break  # outside used range
            nrows = i - 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == "up":
            i = -1
            while True:
                row_ix = self.row - 1 + i
                if row_ix >= 0 and self.sheet.api["values"][row_ix][
                    self.column - 1
                ] not in (None, ""):
                    i -= 1
                else:
                    break
            nrows = i + 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == "right":
            i = 1
            while True:
                try:
                    if self.sheet.api["values"][self.row - 1][
                        self.column - 1 + i
                    ] not in (None, ""):
                        i += 1
                    else:
                        break
                except IndexError:
                    break  # outside used range
            ncols = i - 1
            return self.sheet.range((self.row, self.column + ncols))
        if direction == "left":
            i = -1
            while True:
                col_ix = self.column - 1 + i
                if col_ix >= 0 and self.sheet.api["values"][self.row - 1][
                    col_ix
                ] not in (None, ""):
                    i -= 1
                else:
                    break
            ncols = i + 1
            return self.sheet.range((self.row, self.column + ncols))

    def autofit(self, axis=None):
        if axis == "rows" or axis == "r":
            self.append_json_action(func="setAutofit", args="rows")
        elif axis == "columns" or axis == "c":
            self.append_json_action(func="setAutofit", args="columns")
        elif axis is None:
            self.append_json_action(func="setAutofit", args="rows")
            self.append_json_action(func="setAutofit", args="columns")

    @property
    def color(self):
        raise NotImplementedError()

    @color.setter
    def color(self, value):
        if not isinstance(value, str):
            raise ValueError('Color must be supplied in hex format e.g., "#FFA500".')
        self.append_json_action(func="setRangeColor", args=value)

    @property
    def formula(self):
        # Formulas aren't part of the payload that the client sends with every
        # request, to keep it small. Use `await get_formula()` instead, which
        # fetches them on demand.
        raise NotImplementedError(
            "Reading formulas synchronously isn't supported on this engine. "
            "Use 'await myrange.get_formula()' to fetch them on demand."
        )

    async def get_formula(self):
        """Fetch this range's formulas as a raw 2D list.

        Always 2D and unsqueezed, which is what `AdjustDimensionsStage` expects.
        The public `Range.get_formula` applies the `ndim` option on top, so that
        what the *user* gets back matches the shape of reading `.value`.
        """
        if sys.platform != "emscripten":
            raise NotImplementedError("get_formula() is only supported in xlwings Lite")
        import js

        data_js = await js.xlwings.getRangeData(
            self.sheet.name, self.address, "formulas"
        )
        return _normalize_jsnull(data_js.to_py())["formulas"]

    @formula.setter
    def formula(self, value):
        nrows, ncols = self.shape
        if not isinstance(value, list):
            # Scalars broadcast over the whole range, like on the other engines.
            self.append_json_action(func="setFormula", values=[[value] * ncols] * nrows)
            return
        if value and not isinstance(value[0], list):
            # A flat list is a row, unless the target is a single column. This
            # mirrors how `.value` treats flat lists.
            value = [[item] for item in value] if ncols == 1 and nrows != 1 else [value]
        if not value or not value[0]:
            return
        # Like `.value`, the data wins over the target's current shape: writing
        # more formulas than the range holds expands it instead of raising.
        target = self
        if len(value) != nrows or len(value[0]) != ncols:
            target = Range(
                self.sheet,
                self.arg1,
                (self.arg1[0] + len(value) - 1, self.arg1[1] + len(value[0]) - 1),
            )
        target.append_json_action(func="setFormula", values=value)

    @property
    def formula2(self):
        # See `formula`: Office.js has no separate formula2, and reading is
        # async on this engine.
        raise NotImplementedError(
            "Reading formulas synchronously isn't supported on this engine. "
            "Use 'await myrange.get_formula()' to fetch them on demand."
        )

    @formula2.setter
    def formula2(self, value):
        # Office.js has no separate formula2: range.formulas already writes
        # dynamic array formulas, which is what formula2 stands for.
        self.formula = value

    @property
    def column_width(self):
        raise NotImplementedError()

    @column_width.setter
    def column_width(self, value):
        if (
            isinstance(value, bool)
            or not isinstance(value, numbers.Real)
            or not 0 <= value <= 255
        ):
            raise ValueError("column_width must be a number between 0 and 255.")
        # Keep the public xlwings unit (characters). The Office.js callback
        # converts it to points using the workbook's actual standard width.
        self.append_json_action(func="setColumnWidth", args=value)

    @property
    def row_height(self):
        raise NotImplementedError()

    @row_height.setter
    def row_height(self, value):
        self.append_json_action(func="setRowHeight", args=value)

    @property
    def wrap_text(self):
        raise NotImplementedError()

    @wrap_text.setter
    def wrap_text(self, value):
        self.append_json_action(func="setWrapText", args=bool(value))

    @property
    def formula_array(self):
        raise NotImplementedError()

    @formula_array.setter
    def formula_array(self, value):
        self.append_json_action(func="setFormulaArray", args=value)

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        self.append_json_action(
            func="addHyperlink", args=[address, text_to_display, screen_tip]
        )

    @property
    def number_format(self):
        raise NotImplementedError()

    @number_format.setter
    def number_format(self, value):
        self.append_json_action(func="setNumberFormat", args=value)

    @property
    def name(self):
        for name in self.sheet.book.api["names"]:
            if name["sheet_index"] == self.sheet.index - 1 and name[
                "address"
            ] == self.address.replace("$", ""):
                return Name(
                    parent=self.sheet.book if name["book_scope"] else self.sheet,
                    api=name,
                )

    @name.setter
    def name(self, value):
        self.append_json_action(
            func="setRangeName",
            args=value,
        )

    def autofill(self, destination, type_):
        types = {
            "fill_copy": "FillCopy",
            "fill_days": "FillDays",
            "fill_default": "FillDefault",
            "fill_formats": "FillFormats",
            "fill_months": "FillMonths",
            "fill_series": "FillSeries",
            "fill_values": "FillValues",
            "fill_weekdays": "FillWeekdays",
            "fill_years": "FillYears",
            "growth_trend": "GrowthTrend",
            "linear_trend": "LinearTrend",
            "flash_fill": "FlashFill",
        }
        if type_ not in types:
            raise XlwingsError(
                f"Invalid autofill type '{type_}'. "
                f"Must be one of: {', '.join(sorted(types))}."
            )
        destination_impl = destination.impl
        if (
            destination_impl.sheet.book is not self.sheet.book
            or destination_impl.sheet.api is not self.sheet.api
        ):
            raise XlwingsError(
                "range.autofill() requires the destination to be on the same sheet."
            )
        self.append_json_action(
            func="rangeAutofill", args=[destination.address, types[type_]]
        )

    def copy(self, destination=None):
        if destination is None:
            raise XlwingsError("range.copy() requires a destination argument.")
        self.append_json_action(
            func="copyRange",
            args=[destination.sheet.index - 1, destination.address],
        )

    def copy_from(self, source_range, copy_type=None, skip_blanks=None, transpose=None):
        self.append_json_action(
            func="copyFromRange",
            args=[
                source_range.sheet.index - 1,
                source_range.address,
                copy_type,
                skip_blanks,
                transpose,
            ],
        )

    def delete(self, shift=None):
        if shift not in ("up", "left"):
            # Non-remote version allows shift=None
            raise XlwingsError(
                "range.delete() requires either 'up' or 'left' as shift arguments."
            )
        self.append_json_action(func="rangeDelete", args=shift)

    def insert(self, shift=None, copy_origin=None):
        if shift not in ("down", "right"):
            raise XlwingsError(
                "range.insert() requires either 'down' or 'right' as shift arguments."
            )
        if copy_origin not in (
            "format_from_left_or_above",
            "format_from_right_or_below",
        ):
            raise XlwingsError(
                "range.insert() requires either 'format_from_left_or_above' or "
                "'format_from_right_or_below' as copy_origin arguments."
            )
        # copy_origin is only supported by VBA clients
        self.append_json_action(func="rangeInsert", args=[shift, copy_origin])

    def select(self):
        self.append_json_action(
            func="rangeSelect",
        )

    def merge(self, across):
        self.append_json_action(func="rangeMerge", args=bool(across))

    def unmerge(self):
        self.append_json_action(func="rangeUnmerge")

    def group(self, by):
        self.append_json_action(func="rangeGroup", args=[by])

    def ungroup(self, by):
        self.append_json_action(func="rangeUngroup", args=[by])

    def adjust_indent(self, amount):
        self.append_json_action(func="rangeAdjustIndent", args=amount)

    def to_png(self, path):
        self.append_json_action(func="rangeToPng", args=[path])

    @property
    def font(self):
        return Font(self, self.sheet.book.api)

    def __len__(self):
        nrows, ncols = self.shape
        return nrows * ncols

    def __call__(self, arg1, arg2=None):
        if arg2 is None:
            col = (arg1 - 1) % self.shape[1]
            row = int((arg1 - 1 - col) / self.shape[1])
            return self(row + 1, col + 1)
        else:
            return Range(
                sheet=self.sheet,
                arg1=(self.row + arg1 - 1, self.column + arg2 - 1),
            )


class Collection(base_classes.Collection):
    def __init__(self, parent):
        self._parent = parent
        self._api = parent.api[self._attr]

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    def __call__(self, key):
        if isinstance(key, numbers.Number):
            if key > len(self):
                raise KeyError(key)
            else:
                return self._wrap(self.parent, key)
        else:
            for ix, i in enumerate(self.api):
                if i["name"] == key:
                    return self._wrap(self.parent, ix + 1)
            raise KeyError(key)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for ix, api in enumerate(self.api):
            yield self._wrap(self._parent, ix + 1)

    def __contains__(self, key):
        if isinstance(key, numbers.Number):
            return 1 <= key <= len(self)
        else:
            for i in self.api:
                if i["name"] == key:
                    return True
            return False


class Picture(base_classes.Picture):
    def __init__(self, parent, key):
        self._parent = parent
        self._api = self.parent.api["pictures"][key - 1]
        self.key = key

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.api["name"] = value
        self.append_json_action(func="setPictureName", args=[self.index - 1, value])

    @property
    def width(self):
        return self.api["width"]

    @width.setter
    def width(self, value):
        self.append_json_action(func="setPictureWidth", args=[self.index - 1, value])

    @property
    def height(self):
        return self.api["height"]

    @height.setter
    def height(self, value):
        self.append_json_action(func="setPictureHeight", args=[self.index - 1, value])

    @property
    def index(self):
        # TODO: make available in public API
        if isinstance(self.key, numbers.Number):
            return self.key
        else:
            for ix, obj in self.api:
                if obj["name"] == self.key:
                    return ix + 1
            raise KeyError(self.key)

    def delete(self):
        self.parent._api["pictures"].pop(self.index - 1)
        self.append_json_action(func="deletePicture", args=self.index - 1)

    def update(self, filename):
        with open(filename, "rb") as image_file:
            encoded_image_string = base64.b64encode(image_file.read()).decode("utf-8")
        self.append_json_action(
            func="updatePicture",
            args=[
                encoded_image_string,
                self.index - 1,
                self.name,
                self.width,
                self.height,
            ],
        )
        return self


class Pictures(Collection, base_classes.Pictures):
    _attr = "pictures"
    _wrap = Picture

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    def add(
        self,
        filename,
        link_to_file=None,
        save_with_document=None,
        left=None,
        top=None,
        width=None,
        height=None,
        anchor=None,
    ):
        if self.parent.book.api["client"] == "Google Apps Script" and (left or top):
            raise ValueError(
                "'left' and 'top' are not supported with Google Sheets. "
                "Use 'anchor' instead."
            )
        if anchor is None:
            column_index = 0
            row_index = 0
        else:
            column_index = anchor.column - 1
            row_index = anchor.row - 1
        # Google Sheets allows a max size of 1 million pixels. For matplotlib, you
        # can control the pixels like so: fig = plt.figure(figsize=(6, 4), dpi=200)
        # This sample has (6 * 200) * (4 * 200) = 960,000 px
        # Note that savefig(bbox_inches="tight") crops the image and therefore will
        # reduce the number of pixels in a non-deterministic way. Existing figure
        # size can be checked via fig.get_size_inches(). pandas accepts figsize also:
        # ax = df.plot(figsize=(3,3))
        # fig = ax.get_figure()
        with open(filename, "rb") as image_file:
            encoded_image_string = base64.b64encode(image_file.read()).decode("utf-8")
        # TODO: width and height are currently ignored but can be set via obj properties
        self.append_json_action(
            func="addPicture",
            args=[
                encoded_image_string,
                column_index,
                row_index,
                left if left else 0,
                top if top else 0,
            ],
        )
        self.parent._api["pictures"].append(
            {"name": "Image", "width": None, "height": None}
        )
        return Picture(self.parent, len(self.parent.api["pictures"]))


class Name(base_classes.Name):
    def __init__(self, parent, api):
        self.parent = parent
        self.api = api

    @property
    def name(self):
        if self.api["book_scope"]:
            return self.api["name"]
        else:
            sheet_name = self.api["scope_sheet_name"]
            if "!" not in self.api["name"]:
                # VBA/Google Sheets already do this
                sheet_name = f"'{sheet_name}'" if " " in sheet_name else sheet_name
                return f"{sheet_name}!{self.api['name']}"
            else:
                return self.api["name"]

    @property
    def refers_to(self):
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        sheet = book.sheets(self.api["sheet_index"] + 1)
        sheet_name = f"'{sheet.name}'" if " " in sheet.name else sheet.name
        return f"={sheet_name}!{sheet.range(self.api['address']).address}"

    @property
    def refers_to_range(self):
        book = self.parent if isinstance(self.parent, Book) else self.parent.book
        sheet = book.sheets(self.api["sheet_index"] + 1)
        return sheet.range(self.api["address"])

    def delete(self):
        # TODO: delete in api
        self.parent.append_json_action(
            func="nameDelete",
            args=[
                self.name,  # this includes the sheet name for sheet scope
                self.refers_to,
                self.api["name"],  # no sheet name
                self.api["sheet_index"],
                self.api["book_scope"],
                self.api["scope_sheet_index"],
            ],
        )


class Names(base_classes.Names):
    def __init__(self, parent, api):
        self.parent = parent
        self.api = api

    def add(self, name, refers_to):
        # TODO: raise backend error in case of duplicates
        if isinstance(self.parent, Book):
            is_parent_book = True
        else:
            is_parent_book = False
        self.parent.append_json_action(func="namesAdd", args=[name, refers_to])

        def _get_sheet_index(parent):
            if is_parent_book:
                sheets = parent.sheets
            else:
                sheets = parent.book.sheets
            for sheet in sheets:
                if sheet.name == refers_to.split("!")[0].replace("=", "").replace(
                    "'", ""
                ):
                    return sheet.index - 1

        return Name(
            self.parent,
            {
                "name": name,
                "sheet_index": _get_sheet_index(self.parent),
                "address": refers_to.split("!")[1].replace("$", ""),
                "book_scope": True if is_parent_book else False,
            },
        )

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            name_or_index -= 1
            if name_or_index > len(self):
                raise KeyError(name_or_index)
            else:
                return Name(self.parent, api=self.api[name_or_index])
        else:
            for ix, i in enumerate(self.api):
                name = Name(self.parent, api=self.api[ix])
                if name.name == name_or_index:
                    # Sheet scope names have the sheet name prepended
                    return name
            raise KeyError(name_or_index)

    def contains(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return 1 <= name_or_index <= len(self)
        else:
            for i in self.api:
                if i["name"] == name_or_index:
                    return True
            return False

    def __len__(self):
        return len(self.api)


engine = Engine()


class Table(base_classes.Table):
    @property
    def show_autofilter(self):
        return self.api["show_autofilter"]

    @show_autofilter.setter
    def show_autofilter(self, value):
        self.append_json_action(
            func="showAutofilterTable", args=[self.index - 1, value]
        )

    def __init__(self, parent, key):
        self._parent = parent
        self._api = self.parent.api["tables"][key - 1]
        self.key = key

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    @property
    def api(self):
        return self._api

    @property
    def parent(self):
        return self._parent

    @property
    def name(self):
        return self.api["name"]

    @name.setter
    def name(self, value):
        self.api["name"] = value
        self.append_json_action(func="setTableName", args=[self.index - 1, value])

    @property
    def range(self):
        if self.api["range_address"]:
            return self.parent.range(self.api["range_address"])
        else:
            return None

    @property
    def header_row_range(self):
        if self.api["header_row_range_address"]:
            return self.parent.range(self.api["header_row_range_address"])
        else:
            return None

    @property
    def data_body_range(self):
        if self.api["data_body_range_address"]:
            return self.parent.range(self.api["data_body_range_address"])
        else:
            return None

    @property
    def totals_row_range(self):
        if self.api["total_row_range_address"]:
            return self.parent.range(self.api["total_row_range_address"])
        else:
            return None

    @property
    def show_headers(self):
        return self.api["show_headers"]

    @show_headers.setter
    def show_headers(self, value):
        self.append_json_action(func="showHeadersTable", args=[self.index - 1, value])

    @property
    def show_totals(self):
        return self.api["show_totals"]

    @show_totals.setter
    def show_totals(self, value):
        self.append_json_action(func="showTotalsTable", args=[self.index - 1, value])

    @property
    def insert_row_range(self):
        # Office.js' Excel.Table has no InsertRowRange equivalent: its only
        # range accessors are getRange(), getDataBodyRange(),
        # getHeaderRowRange() and getTotalRowRange(). Returning None would be
        # indistinguishable from the documented "table isn't empty" answer, so
        # raise instead of answering wrongly.
        raise NotImplementedError(
            "Table.insert_row_range is not supported in Office.js, which has no "
            "InsertRowRange equivalent."
        )

    @property
    def display_name(self):
        # Office.js' Excel.Table only has `name`, so display_name aliases it.
        # The two are equivalent in practice anyway: on macOS, setting
        # display_name changes the name too, and Office Scripts dropped the
        # distinction as well.
        return self.name

    @display_name.setter
    def display_name(self, value):
        self.name = value

    @property
    def show_table_style_first_column(self):
        return self.api["show_table_style_first_column"]

    @show_table_style_first_column.setter
    def show_table_style_first_column(self, value):
        self.api["show_table_style_first_column"] = value
        self.append_json_action(
            func="showTableStyleFirstColumn", args=[self.index - 1, value]
        )

    @property
    def show_table_style_last_column(self):
        return self.api["show_table_style_last_column"]

    @show_table_style_last_column.setter
    def show_table_style_last_column(self, value):
        self.api["show_table_style_last_column"] = value
        self.append_json_action(
            func="showTableStyleLastColumn", args=[self.index - 1, value]
        )

    @property
    def show_table_style_row_stripes(self):
        return self.api["show_table_style_row_stripes"]

    @show_table_style_row_stripes.setter
    def show_table_style_row_stripes(self, value):
        self.api["show_table_style_row_stripes"] = value
        self.append_json_action(
            func="showTableStyleRowStripes", args=[self.index - 1, value]
        )

    @property
    def show_table_style_column_stripes(self):
        return self.api["show_table_style_column_stripes"]

    @show_table_style_column_stripes.setter
    def show_table_style_column_stripes(self, value):
        self.api["show_table_style_column_stripes"] = value
        self.append_json_action(
            func="showTableStyleColumnStripes", args=[self.index - 1, value]
        )

    @property
    def table_style(self):
        return self.api["table_style"]

    @table_style.setter
    def table_style(self, value):
        self.append_json_action(func="setTableStyle", args=[self.index - 1, value])

    @property
    def index(self):
        # TODO: make available in public API
        if isinstance(self.key, numbers.Number):
            return self.key
        else:
            for ix, obj in self.api:
                if obj["name"] == self.key:
                    return ix + 1
            raise KeyError(self.key)

    def resize(self, range):
        self.append_json_action(
            func="resizeTable", args=[self.index - 1, range.address]
        )


class Tables(Collection, base_classes.Tables):
    _attr = "tables"
    _wrap = Table

    def append_json_action(self, **kwargs):
        self.parent.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.parent.index - 1,
                },
            }
        )

    def add(
        self,
        source_type=None,
        source=None,
        link_source=None,
        has_headers=None,
        destination=None,
        table_style_name=None,
        name=None,
    ):
        self.append_json_action(
            func="addTable",
            args=[source.address, has_headers, table_style_name, name],
        )
        self.parent._api["tables"].append(
            {
                "name": "",
                "range_address": None,
                "header_row_range_address": None,
                "data_body_range_address": None,
                "total_row_range_address": None,
                "show_headers": None,
                "show_totals": None,
                "table_style": "",
                # Excel's defaults for a new table, so the getters work before
                # the next round-trip refreshes the payload.
                "show_autofilter": True,
                "show_table_style_first_column": False,
                "show_table_style_last_column": False,
                "show_table_style_row_stripes": True,
                "show_table_style_column_stripes": False,
            }
        )
        return Table(self.parent, len(self.parent.api["tables"]))


class FreezePanes(base_classes.FreezePanes):
    def __init__(self, sheet):
        self.sheet = sheet

    def append_json_action(self, **kwargs):
        self.sheet.book.append_json_action(
            **{
                **kwargs,
                **{
                    "sheet_position": self.sheet.index - 1,
                },
            }
        )

    def freeze_at(self, frozen_range):
        self.append_json_action(func="freezePaneAtRange", args=[frozen_range])

    def unfreeze(self):
        self.append_json_action(func="freezePaneUnfreeze")


class Font(base_classes.Font):
    # TODO: support Shape and getters
    def __init__(self, parent, api):
        self.parent = parent
        self._api = api

    def append_json_action(self, **kwargs):
        if isinstance(self.parent, Range):
            self.parent.append_json_action(
                **{
                    **kwargs,
                }
            )
        else:
            raise NotImplementedError()

    @property
    def api(self):
        return self._api

    @property
    def bold(self):
        raise NotImplementedError()

    @bold.setter
    def bold(self, value):
        self.append_json_action(func="setFontProperty", args=["bold", value])

    @property
    def italic(self):
        raise NotImplementedError()

    @italic.setter
    def italic(self, value):
        self.append_json_action(func="setFontProperty", args=["italic", value])

    @property
    def size(self):
        raise NotImplementedError()

    @size.setter
    def size(self, value):
        self.append_json_action(func="setFontProperty", args=["size", value])

    @property
    def color(self):
        raise NotImplementedError()

    @color.setter
    def color(self, color_or_rgb):
        if not isinstance(color_or_rgb, str):
            raise ValueError('Color must be supplied in hex format e.g., "#FFA500".')
        self.append_json_action(func="setFontProperty", args=["color", color_or_rgb])

    @property
    def name(self):
        raise NotImplementedError()

    @name.setter
    def name(self, value):
        self.append_json_action(func="setFontProperty", args=["name", value])


if __name__ == "__main__":
    # python -m xlwings.pro._xlremote
    import inspect

    def print_unimplemented_attributes(class_name, base_class, derived_class=None):
        if class_name == "Apps":
            return
        base_attributes = set(
            attr
            for attr in vars(base_class)
            if not (attr.startswith("_") or attr == "api")
        )
        if derived_class:
            derived_attributes = set(
                attr for attr in vars(derived_class) if not attr.startswith("_")
            )
        else:
            derived_attributes = set()
        unimplemented_attributes = base_attributes - derived_attributes

        if unimplemented_attributes:
            print("")
            print(f"    xlwings.{class_name}")
            print("")
            for attribute in unimplemented_attributes:
                if not attribute.startswith("__") and attribute not in (
                    "api",
                    "xl",
                    "hwnd",
                ):
                    if callable(getattr(base_class, attribute)):
                        print(f"        - {attribute}()")
                    else:
                        print(f"        - {attribute}")

    for name, obj in inspect.getmembers(base_classes):
        if inspect.isclass(obj):
            print_unimplemented_attributes(name, obj, globals().get(name))
