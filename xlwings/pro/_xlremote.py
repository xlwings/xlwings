"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import base64
import datetime as dt
import numbers
import re
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

# Time types (doesn't contain dt.date)
time_types = (dt.datetime,)
if np:
    time_types = time_types + (np.datetime64,)

datetime_pattern = r"^(-?(?:[1-9][0-9]*)?[0-9]{4})-(1[0-2]|0[1-9])-(3[01]|0[1-9]|[12][0-9])T(2[0-3]|[01][0-9]):([0-5][0-9]):([0-5][0-9])(\.[0-9]+)?(Z|[+-](?:2[0-3]|[01][0-9]):[0-5][0-9])?$"  # noqa: E501
datetime_regex = re.compile(datetime_pattern)


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
        elif np and isinstance(x, np.datetime64):
            return utils.np_datetime_to_datetime(x).replace(tzinfo=None).isoformat()
        elif pd and isinstance(x, pd.Timestamp):
            return x.to_pydatetime().replace(tzinfo=None).isoformat()
        elif pd and isinstance(x, type(pd.NaT)):
            return None
        elif isinstance(x, time_types):
            x = x.replace(tzinfo=None).isoformat()
        elif isinstance(x, dt.date):
            # JS applies tz conversion with "2021-01-01" when calling
            # toLocaleDateString() while it leaves "2021-01-01T00:00:00" unchanged
            x = dt.datetime(x.year, x.month, x.day).isoformat()
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
    def selection(self):
        book = self.books.active
        return Range(sheet=book.sheets.active, arg1=book.api["book"]["selection"])

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

    def open(self, json):
        book = Book(api=json, books=self)
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
    def __init__(self, api, books):
        self.books = books
        self._api = api
        self._json = {"actions": []}
        if api["version"] != __version__:
            raise XlwingsError(
                f'Your xlwings version is different on the client ({api["version"]}) '
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
        # Default naming logic is different from Desktop apps!
        sheet_number = 1
        while True:
            if f"Sheet{sheet_number}" in [sheet.name for sheet in self]:
                sheet_number += 1
            else:
                break
        api = {
            "name": f"Sheet{sheet_number}",
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
            arg2=(1_048_576, 16_384),
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
        del self.book.api["sheets"][self.index - 1]
        self.append_json_action(func="sheetDelete")

    def clear(self):
        self.append_json_action(func="sheetClear")

    def clear_contents(self):
        self.append_json_action(func="sheetClearContents")

    def clear_formats(self):
        self.append_json_action(func="sheetClearFormats")


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

    @property
    def raw_value(self):
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
                    if self.sheet.api["values"][self.row - 1 + i][self.column - 1]:
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
                if row_ix >= 0 and self.sheet.api["values"][row_ix][self.column - 1]:
                    i -= 1
                else:
                    break
            nrows = i + 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == "right":
            i = 1
            while True:
                try:
                    if self.sheet.api["values"][self.row - 1][self.column - 1 + i]:
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
                if col_ix >= 0 and self.sheet.api["values"][self.row - 1][col_ix]:
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

    def copy(self, destination=None):
        # TODO: introduce the new copy_from from Office Scripts
        if destination is None:
            raise XlwingsError("range.copy() requires a destination argument.")
        self.append_json_action(
            func="copyRange",
            args=[destination.sheet.index - 1, destination.address],
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
            encoded_image_string = base64.b64encode(image_file.read())
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
            encoded_image_string = base64.b64encode(image_file.read())
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
            }
        )
        return Table(self.parent, len(self.parent.api["tables"]))
