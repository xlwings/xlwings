"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import datetime as dt
import numbers
from pathlib import Path

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

from xlwings import xlwingslib

from .. import NoSuchObjectError, base_classes, utils

MAX_ROWS = 1_048_576
MAX_COLUMNS = 16_384


def _clean_value_data_element(value, datetime_builder, empty_as, number_builder):
    if value is None:
        return empty_as
    elif isinstance(value, dt.datetime) and datetime_builder is not dt.datetime:
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
        # err_to_str is handled in raw_values for efficiency
        if empty_as or number_builder or datetime_builder is not dt.datetime:
            return [
                [
                    _clean_value_data_element(
                        c, datetime_builder, empty_as, number_builder
                    )
                    for c in row
                ]
                for row in data
            ]
        else:
            return data

    @staticmethod
    def prepare_xl_data_element(x, options):
        return x

    @property
    def name(self):
        return "calamine"

    @property
    def type(self):
        return "reader"


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
    def visible(self):
        return True

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

    def open(self, filename):
        filename = str(Path(filename).resolve())
        sheet_names = xlwingslib.get_sheet_names(filename)
        names = []
        for name, ref in xlwingslib.get_defined_names(filename):
            if ref.split("!")[0].strip("'") in sheet_names:
                names.append(
                    {
                        "name": name,
                        "sheet_index": sheet_names.index(ref.split("!")[0].strip("'")),
                        "address": ref.split("!")[1],
                        "book_scope": True,  # TODO: not provided by calamine
                    }
                )
        book = Book(
            api={
                "sheet_names": sheet_names,
                "names": names,
            },
            books=self,
            path=filename,
        )
        self.books.append(book)
        self._active = book
        return book

    def add(self):
        book = Book(api={"sheet_names": ["Sheet1"]}, books=self, path="dummy")
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
    def __init__(self, api, books, path):
        self._api = api
        self.books = books
        self.path = path

    @property
    def api(self):
        return self._api

    @property
    def name(self):
        return Path(self.fullname).name

    @property
    def fullname(self):
        return self.path

    @property
    def names(self):
        return Names(parent=self, api=self.api["names"])

    @property
    def sheets(self):
        return Sheets(book=self)

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
    def __init__(self, book):
        self.book = book

    @property
    def api(self):
        return None

    def __call__(self, name_or_index):
        if isinstance(name_or_index, str):
            sheet_names = self.book.api["sheet_names"]
            if name_or_index not in sheet_names:
                raise NoSuchObjectError(f"Sheet {name_or_index} doesn't exist.")
            else:
                ix = self.book.api["sheet_names"].index(name_or_index) + 1
        else:
            ix = name_or_index

        return Sheet(book=self.book, sheet_index=ix)

    def __len__(self):
        return len(self.book.api["sheet_names"])

    def __iter__(self):
        for ix, sheet in enumerate(self.book.api["sheet_names"]):
            yield Sheet(book=self.book, sheet_index=ix + 1)


class Sheet(base_classes.Sheet):
    def __init__(self, book, sheet_index):
        self._api = {}  # used by e.g., Range.end()
        self._book = book
        self.sheet_index = sheet_index

    @property
    def api(self):
        return self._api

    @property
    def name(self):
        return self.book.api["sheet_names"][self.index - 1]

    @property
    def index(self):
        return self.sheet_index

    @property
    def book(self):
        return self._book

    def range(self, arg1, arg2=None):
        return Range(sheet=self, book=self.book, arg1=arg1, arg2=arg2)

    @property
    def cells(self):
        return Range(
            sheet=self,
            book=self.book,
            arg1=(1, 1),
            arg2=(MAX_ROWS, MAX_COLUMNS),
        )


class Range(base_classes.Range):
    def __init__(self, sheet, book, arg1, arg2=None):
        self.sheet = sheet
        self.book = book
        self.options = None  # Assigned by main.Range to keep API of sheet.range clean

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

    @property
    def api(self):
        return self.raw_value

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
        err_to_str = self.options.get("err_to_str", False)
        if self.arg2 is None:
            self.arg2 = self.arg1
        if self.arg2[0] == MAX_ROWS and self.arg2[1] == MAX_COLUMNS:
            # Whole sheet via sheet.cells
            if not self.sheet.api.get(f"values_err_to_str_{err_to_str}"):
                values = xlwingslib.get_sheet_values(
                    self.book.fullname, self.sheet.index - 1, err_to_str
                )
                self.sheet._api[f"values_err_to_str_{err_to_str}"] = values
                return values
        else:
            # Specific range
            return xlwingslib.get_range_values(
                self.book.fullname,
                self.sheet.index - 1,
                (self.arg1[0] - 1, self.arg1[1] - 1),
                (self.arg2[0] - 1, self.arg2[1] - 1),
                err_to_str,
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
        err_to_str = self.options.get("err_to_str", False)
        if not self.sheet.api.get(f"values_err_to_str_{err_to_str}"):
            values = xlwingslib.get_sheet_values(
                self.book.fullname, self.sheet.index - 1, err_to_str
            )
            self.sheet._api[f"values_err_to_str_{err_to_str}"] = values
        else:
            values = self.sheet.api[f"values_err_to_str_{err_to_str}"]
        if direction == "down":
            i = 1
            while True:
                try:
                    if values[self.row - 1 + i][self.column - 1]:
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
                if row_ix >= 0 and values[row_ix][self.column - 1]:
                    i -= 1
                else:
                    break
            nrows = i + 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == "right":
            i = 1
            while True:
                try:
                    if values[self.row - 1][self.column - 1 + i]:
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
                if col_ix >= 0 and values[self.row - 1][col_ix]:
                    i -= 1
                else:
                    break
            ncols = i + 1
            return self.sheet.range((self.row, self.column + ncols))

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
                book=self.book,
                arg1=(self.row + arg1 - 1, self.column + arg2 - 1),
            )


class Name(base_classes.Name):
    def __init__(self, parent, api):
        self.parent = parent  # only implemented for Book, not Sheet
        self.api = api

    @property
    def name(self):
        return self.api["name"]

    @property
    def refers_to(self):
        sheet_name = self.parent.sheets(self.api["sheet_index"] + 1).name
        sheet_name = f"'{sheet_name}'" if " " in sheet_name else sheet_name
        return f"={sheet_name}!{self.api['address']}"

    @property
    def refers_to_range(self):
        return self.parent.sheets(self.api["sheet_index"] + 1).range(
            self.api["address"]
        )


class Names(base_classes.Names):
    def __init__(self, parent, api):
        self.parent = parent  # only implemented for Book, not Sheet
        self.api = api

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            name_or_index -= 1
            if name_or_index > len(self):
                raise KeyError(name_or_index)
            else:
                return Name(self.parent, api=self.api[name_or_index])
        else:
            for ix, i in enumerate(self.api):
                if i["name"] == name_or_index:
                    return Name(self.parent, api=self.api[ix])
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
