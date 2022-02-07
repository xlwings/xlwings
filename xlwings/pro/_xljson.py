"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import re
import numbers
import datetime as dt
from functools import lru_cache

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

from .. import utils, base_classes, __version__, XlwingsError


# Time types (doesn't contain dt.date)
time_types = (dt.datetime,)
if np:
    time_types = time_types + (np.datetime64,)


def _clean_value_data_element(value, datetime_builder, empty_as, number_builder):
    if value == "":
        return empty_as
    if isinstance(value, str):
        # TODO: Send arrays back and forth with indices of the location of dt values
        pattern = r"^(-?(?:[1-9][0-9]*)?[0-9]{4})-(1[0-2]|0[1-9])-(3[01]|0[1-9]|[12][0-9])T(2[0-3]|[01][0-9]):([0-5][0-9]):([0-5][0-9])(\.[0-9]+)?(Z|[+-](?:2[0-3]|[01][0-9]):[0-5][0-9])?$"
        if re.compile(pattern).match(value):
            value = dt.datetime.fromisoformat(
                value[:-1]
            )  # cut off "Z" (Python doesn't accept it and Excel doesn't support tz)
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
    elif number_builder is not None and type(value) == float:
        value = number_builder(value)
    return value


class Engine:
    def __init__(self):
        self.apps = Apps()

    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder):
        return [
            [
                _clean_value_data_element(c, datetime_builder, empty_as, number_builder)
                for c in row
            ]
            for row in data
        ]

    @staticmethod
    def prepare_xl_data_element(x):
        if x is None:
            return ""
        elif np and isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif np and isinstance(x, np.number):
            return float(x)
        elif np and isinstance(x, np.datetime64):
            return utils.np_datetime_to_datetime(x).replace(tzinfo=None).isoformat()
        elif pd and isinstance(x, pd.Timestamp):
            return x.to_pydatetime().replace(tzinfo=None)
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
        return "json"


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
            api = self.api[name_or_index - 1]
            ix = name_or_index - 1
        else:
            api = None
            for ix, sheet in enumerate(self.api):
                if sheet["name"] == name_or_index:
                    api = sheet
                    break
                else:
                    continue
        if api is None:
            raise ValueError(f"Sheet '{name_or_index}' doesn't exist!")
        else:
            return Sheet(api=api, sheets=self, index=ix + 1)

    def add(self, before=None, after=None):
        # Default naming logic is different from Desktop apps!
        sheet_number = 1
        while True:
            if f"Sheet{sheet_number}" in [sheet.name for sheet in self]:
                sheet_number += 1
            else:
                break
        api = {"name": f"Sheet{sheet_number}", "values": [[]]}
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
        self.book.append_json_action(func="addSheet", args=ix - 1)
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

    def activate(self):
        ix = self.index - 1
        self.book.api["book"]["active_sheet_index"] = ix
        self.append_json_action(func="activateSheet", args=ix)


@lru_cache(None)
def get_range_api(api_values, arg1, arg2=None):
    # Keeping this outside of the Range class allows us to cache it across multiple
    # instances of the same range
    if arg2:
        values = [
            row[arg1[1] - 1 : arg2[1]] for row in api_values[arg1[0] - 1 : arg2[0]]
        ]
        if not values:
            # Outside the used range
            values = [[None] * (arg2[1] + 1 - arg1[1])] * (arg2[0] + 1 - arg1[0])
        return values
    else:
        try:
            values = [[api_values[arg1[0] - 1][arg1[1] - 1]]]
            return values
        except IndexError:
            # Outside the used range
            return [[None]]


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
            if ":" in arg1:
                address1, address2 = arg1.split(":")
                arg1 = utils.address_to_index_tuple(address1.upper())
                arg2 = utils.address_to_index_tuple(address2.upper())
            else:
                arg1 = utils.address_to_index_tuple(arg1.upper())
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
            func="clearContents",
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
        raise NotImplemented()

    @color.setter
    def color(self, value):
        if not isinstance(value, str):
            raise ValueError('Color must be supplied in hex format e.g., "#FFA500".')
        self.append_json_action(func="setRangeColor", args=value)

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


engine = Engine()
