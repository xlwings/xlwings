import re
import json
import datetime as dt
import logging
from functools import lru_cache

try:
    import numpy as np
except ImportError:
    np = None

from . import utils, platform_base_classes

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
            value = dt.datetime.fromisoformat(
                value[:-1]
            )  # cutting off "Z" (Excel doesn't support time-zones)
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


def generate_response(**kwargs):
    return {
        'data': kwargs.get('data'),
        'sheet_name': kwargs.get('sheet_name'),
        'start_row': kwargs.get('start_row'),
        'start_column': kwargs.get('start_column'),
        'row_count': kwargs.get('row_count'),
        'column_count': kwargs.get('column_count'),
        'func': kwargs.get('func'),
    }


class Apps:
    def __init__(self):
        self._apps = [App(self)]

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

    @property
    def books(self):
        return self._books

    @property
    def pid(self):
        return self._pid


class Books:
    def __init__(self, app):
        self.app = app
        self.books = []

    def open(self, json):
        book = Book(api=json, books=self)
        self.books.append(book)
        return book

    def add(self, json):
        pass


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


class Sheets:
    def __init__(self, api, book):
        self.api = api
        self.book = book

    @property
    def active(self):
        return Sheet(api=self.api[self.book.api['book']['active_sheet_index']], sheets=self)

    def __call__(self, name_or_index):
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


class Sheet:
    def __init__(self, api, sheets):
        self.api = api
        self.sheets = sheets

    @property
    def name(self):
        return self.api['name']

    @property
    def book(self):
        return self.sheets.book

    def range(self, arg1, arg2=None):
        return Range(sheet=self, api=self.api, arg1=arg1, arg2=arg2)

    @property
    def cells(self):
        return Range(
            sheet=self,
            api=self.api,
            arg1=(1, 1),
            arg2=(len(self.api['values']), len(self.api['values'][0])),
        )


class Range(platform_base_classes.Range):
    def __init__(self, sheet, api, arg1, arg2=None):
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
            values = [
                row[self.arg1[1] - 1 : self.arg2[1]]
                for row in self._api['values'][self.arg1[0] - 1 : self.arg2[0]]
            ]
            if not values:
                # Outside the used range
                values = [[None] * (self.arg2[1] + 1 - self.arg1[1])] * (
                    self.arg2[0] + 1 - self.arg1[0]
                )
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
        # TODO: should 1x1 and 1xn be returned as scalar and list?
        return self.api

    @raw_value.setter
    def raw_value(self, value):
        data = [[value]] if not isinstance(value, list) else value
        self.sheet.book._json.append(
            generate_response(
                data=data,
                sheet_name=self.sheet.name,
                start_row=self.row - 1,
                start_column=self.column - 1,
                row_count=len(data),
                column_count=len(data[0]),
            )
        )

    def clear_contents(self):
        nrows, ncols = self.shape
        self.sheet.book._json.append(
            generate_response(
                sheet_name=self.sheet.name,
                start_row=self.row - 1,
                start_column=self.column - 1,
                row_count=nrows,
                column_count=ncols,
                func='clear_contents',
            )
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
        # TODO: left, up, 2d case
        if direction == 'down':
            i = 1
            while True:
                try:
                    if self.sheet.api['values'][self.row - 1 + i][self.column - 1]:
                        i += 1
                    else:
                        break
                except IndexError:
                    break  # outside of used range
            nrows = i - 1
            return self.sheet.range((self.row + nrows, self.column))
        if direction == 'right':
            i = 1
            while True:
                try:
                    if self.sheet.api['values'][self.row - 1][self.column - 1 + i]:
                        i += 1
                    else:
                        break
                except IndexError:
                    break  # outside of used range
            ncols = i - 1
            return self.sheet.range((self.row, self.column + ncols))

    def __len__(self):
        return len(self.api) * len(self.api[0])

    def __call__(self, row, col):
        return Range(
            sheet=self.sheet, api=self.sheet.api, arg1=(self.row + row - 1, self.column + col - 1)
        )


engine = Engine()
