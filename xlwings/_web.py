import re
import json
import numbers
import datetime as dt
import logging
from functools import lru_cache

try:
    import numpy as np
except ImportError:
    np = None

from . import utils
from . import platform_base_classes

logger = logging.getLogger(__name__)


USER_CONFIG_FILE = '/todo'

# Time types
time_types = (dt.date, dt.datetime)
if np:
    time_types = time_types + (np.datetime64,)


class Engine:
    def __init__(self):
        self.apps = Apps()

    @property
    def name(self):
        return "web"


def _clean_value_data_element(value, datetime_builder, empty_as, number_builder):
    if value == '':
        return empty_as
    if isinstance(value, str):
        pattern = r'^(-?(?:[1-9][0-9]*)?[0-9]{4})-(1[0-2]|0[1-9])-(3[01]|0[1-9]|[12][0-9])T(2[0-3]|[01][0-9]):([0-5][0-9]):([0-5][0-9])(\.[0-9]+)?(Z|[+-](?:2[0-3]|[01][0-9]):[0-5][0-9])?$'
        if re.compile(pattern).match(value):
            value = dt.datetime.fromisoformat(value[:-1])  # cutting off "Z" (Excel doesn't support time-zones)
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


def clean_value_data(data, datetime_builder, empty_as, number_builder):
    return [
        [_clean_value_data_element(c, datetime_builder, empty_as, number_builder) for c in row]
        for row in data
    ]


Engine.clean_value_data = staticmethod(clean_value_data)


def prepare_xl_data_element(x):
    return x


Engine.prepare_xl_data_element = staticmethod(prepare_xl_data_element)


class Apps:
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


class App:

    _next_pid = -1

    def __init__(self, apps, add_book=True, **kwargs):
        self.apps = apps
        self._pid = App._next_pid
        App._next_pid -= 1
        self._books = Books(self)
        if add_book:
            self._books.add(json=[])

    @property
    def engine(self):
        return engine

    def activate(self, steal_focus=False):
        raise NotImplementedError()

    @property
    def books(self):
        return self._books

    @property
    def pid(self):
        return self._pid

    def kill(self):
        self.apps._apps.remove(self)
        self.apps = None

    def range(self, arg1, arg2=None):
        # TODO: better implementation
        return self.books.active.sheets.active.range(arg1)


class Books:
    def __init__(self, app):
        self.app = app
        self.books = []
        self.active = None

    def open(self, json):
        book = Book(api=json, books=self)
        self.books.append(book)
        self.active = book
        return book

    @property
    def api(self):
        return None

    def add(self, json):
        book = Book(api=json, books=self)
        self.books.append(book)
        self.active = book
        return book

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return self.books[name_or_index - 1]
        else:
            book = self._try_find_book_by_name(name_or_index)
            if book is None:
                raise KeyError(name_or_index)
            return book

    def __len__(self):
        return len(self.books)

    def __iter__(self):
        for book in self.books:
            yield book


class Book:
    def __init__(self, api, books):
        self.api = api
        self.books = books
        self._json = []

    def json(self):
        return json.dumps(self._json, default=lambda d: d.isoformat())

    @property
    def name(self):
        return self.api['book']['name']

    @property
    def fullname(self):
        return self.name

    @property
    def sheets(self):
        return Sheets(api=self.api['sheets'], book=self)

    @property
    def app(self):
        return self.books.app

    @property
    def index(self):
        return self.app.books.index(self)

    def close(self):
        assert self.api is not None, "Seems this book was already closed."
        self.books.books.remove(self)
        self.books = None
        self.api = None

    def save(self, path=None):
        pass


class Sheets:
    def __init__(self, api, book):
        self.api = api
        self.book = book

    @property
    def active(self):
        return Sheet(api=self.api[self.book.api['book']['active_sheet_index']], sheets=self)

    def __call__(self, name_or_index):
        api = None
        if isinstance(name_or_index, int):
            api = self.api[name_or_index - 1]
        else:
            api = None
            for sheet in self.api:
                if sheet['name'] == name_or_index:
                    api = sheet
                    break
                else:
                    continue
        if api is None:
            raise ValueError(f"Sheet '{name_or_index}' doesn't exist!")
        else:
            return Sheet(api=api, sheets=self)

    def __len__(self):
        return len(self.api)

    def __iter__(self):
        for sheet in self.api:
            yield Sheet(api=sheet, sheets=self)

    # def add(self, before=None, after=None):
    #     return Sheet(api=self.book.api.create_sheet(), sheets=self)


class Sheet:
    def __init__(self, api, sheets):
        self.api = api
        self.sheets = sheets

    @property
    def name(self):
        return self.api['name']

    @name.setter
    def name(self, value):
        self.api.title = value

    @property
    def book(self):
        return self.sheets.book

    @property
    def index(self):
        return self.sheets.book.index(self.api)

    def range(self, arg1, arg2=None):
        return Range(sheet=self, api=self.api, arg1=arg1, arg2=arg2)

    def activate(self):
        self.sheets.book.api.active_sheet = self.sheets.book.api.index(self.api)

    def select(self):
        self.sheets.book.api.active_sheet = self.sheets.book.api.index(self.api)

    def clear(self):
        return NotImplementedError()

    def autofit(self, axis=None):
        logger.warning("Autofit doesn't do anything in openpyxl engine.")

    @property
    def cells(self):
        return Range(api=self.api['values'], sheet=self, row_ix=1, col_ix=1)


class Range(platform_base_classes.Range):
    def __init__(self, sheet, api, arg1, arg2=None):
        # Range
        if isinstance(arg1, Range):
            arg1 = arg1.coords[1], arg1.coords[2]
        if isinstance(arg2, Range):
            arg2 = arg2.coords[1], arg2.coords[2]
        # A1 notation
        if isinstance(arg1, str):
            # A1 notation
            if ":" in arg1:
                address1, address2 = arg1.split(':')
                arg1 = utils.address_to_index_tuple(address1.upper())
                arg2 = utils.address_to_index_tuple(address2.upper())
            else:
                arg1 = utils.address_to_index_tuple(arg1.upper())
        # Coordinates
        if len(arg1) == 4:
            row, col, nrows, ncols = arg1
            arg1 = (row, col)
            if nrows > 1 or ncols > 1:
                arg2 = (row + nrows, col + ncols)

        self.arg1 = arg1  # 1-based tuple
        self.arg2 = arg2  # 1-based tuple
        self.sheet = sheet
        self._api = api

    @property
    @lru_cache(None)
    def api(self):
        if self.arg2:
            values = [row[self.arg1[1] - 1 : self.arg2[1]] for row in self._api['values'][self.arg1[0] - 1 : self.arg2[0]]]
            if not values:
                # Outside the used range
                values = [[None] * (self.arg2[1] + 1 - self.arg1[1])] * (self.arg2[0] + 1 - self.arg1[0])
            return values
        else:
            try:
                values = [[self._api['values'][self.arg1[0] - 1][self.arg1[1] - 1]]]
                return values
            except IndexError:
                # Outside the used range
                return [[None]]

    @property
    def coords(self):
        return self.sheet.name, self.row, self.column, len(self.api), len(self.api[0])

    def __len__(self):
        return len(self.api) * len(self.api[0])

    @property
    def row(self):
        return self.arg1[0]

    @property
    def column(self):
        return self.arg1[1]

    @property
    def shape(self):
        return len(self.api), len(self.api[0])

    @property
    def raw_value(self):
        return self.api

    @raw_value.setter
    def raw_value(self, value):
        self.sheet.book._json.append(
            {
                'data': [[value]] if not isinstance(value, list) else value,
                'sheet_name': self.sheet.name,
                'start_row': self.coords[1] - 1,
                'start_column': self.coords[2] - 1
            }
        )

    @property
    def address(self):
        nrows, ncols = self.shape
        address = f'${utils.col_name(self.column)}${self.row}'
        if nrows == 1 and ncols == 1:
            return address
        else:
            return f'{address}:${utils.col_name(self.column + ncols - 1)}${self.row + nrows - 1}'

    @property
    def has_array(self):
        # TODO
        return False

    def end(self, direction):
        # TODO: left, up
        if direction == 'down':
            i = 1
            while True:
                if self.sheet.api['values'][self.row - 1 + i][self.column - 1]:
                    i += 1
                else:
                    break
            nrows = i - 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == 'right':
            i = 1
            while True:
                if self.sheet.api['values'][self.row - 1][self.column - 1 + i]:
                    i += 1
                else:
                    break
            ncols = i - 1
            return self.sheet.range((self.row, self.column + ncols))

    def __call__(self, row, col):
        print(row, col)
        return Range(sheet=self.sheet, api=self.sheet.api, arg1=(self.row + row - 1, self.column + col - 1))


engine = Engine()
