try:
    import openpyxl
except ImportError:
    openpyxl = None

import os


class App(object):

    def __init__(self):
        pass

    @property
    def books(self):
        return Books()


class Books(object):

    def __init__(self):
        self.books = {}

    def open(self, filename):
        book = self.books.get(filename, None)
        if book is None:
            self.books[filename] = book = Book(filename)
        return book

    def __call__(self, name_or_index):
        return self.books[name_or_index]

    def __len__(self):
        return len(self.books)

    def __iter__(self):
        for book in self.books.values():
            yield book


class Book(object):

    def __init__(self, filename):
        self.api = openpyxl.load_workbook(filename)
        self.filename = filename

    @property
    def name(self):
        _, n = os.path.split(self.filename)
        return n

    @property
    def sheets(self):
        return Sheets(book=self)


class Sheets(object):

    def __init__(self, book):
        self.book = book

    def __call__(self, name_or_index):
        if isinstance(name_or_index, int):
            api = self.book.api.worksheets[name_or_index - 1]
        else:
            api = self.book.api[name_or_index]
        return Sheet(api=api, book=self.book)

    def __len__(self):
        return len(self.book.api.worksheets)

    def __iter__(self):
        for sheet in self.book.api.worksheets:
            yield Sheet(api=sheet, book=self.book)


class Sheet(object):

    def __init__(self, api, book):
        self.api = api
        self.book = book

    @property
    def name(self):
        return self.api.title

    def range(self, arg1, arg2=None):
        return Range(api=self.api[arg1], sheet=self)


class Range(object):

    def __init__(self, api, sheet):
        self.api = api
        self.sheet = sheet

    @property
    def raw_value(self):
        return self.api.value

    @property
    def address(self):
        return self.api.coordinate

    def __len__(self):
        return 1