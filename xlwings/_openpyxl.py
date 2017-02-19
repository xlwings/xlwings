try:
    import openpyxl
except ImportError:
    openpyxl = None

import os
import numbers


class Engine(object):

    @property
    def apps(self):
        return Apps()

    @property
    def name(self):
        return "Openpyxl"

engine = Engine()


class Apps(object):

    def __iter__(self):
        yield the_app

    def __len__(self):
        return 1

    def __getitem__(self, index):
        return the_app


class App(object):

    def __init__(self):
        self._books = Books()


    @property
    def engine(self):
        return engine

    @property
    def books(self):
        return self._books

    @property
    def pid(self):
        return 0


class Books(object):

    def __init__(self):
        self.books = []
        self.books_by_filename = {}
        self.active = self.add()

    def open(self, filename):
        book = self.books_by_filename.get(filename, None)
        if book is None:
            self.books_by_filename[filename] = book = Book(openpyxl.load_workbook(filename), filename)
            self.books.append(book)
        return book

    @property
    def api(self):
        return None

    def add(self):
        name = "Book" + str(len(self.books) + 1)
        self.books_by_filename[name] = book = Book(openpyxl.Workbook(), name)
        self.books.append(book)
        self.active = book
        return book

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return self.books[name_or_index - 1]
        else:
            return self.books_by_filename[name_or_index]

    def __len__(self):
        return len(self.books)

    def __iter__(self):
        for book in self.books:
            yield book


class Book(object):

    def __init__(self, api, filename):
        self.api = api
        self.filename = filename

    @property
    def name(self):
        _, n = os.path.split(self.filename)
        return n

    @property
    def sheets(self):
        return Sheets(book=self)

    @property
    def app(self):
        return the_app

    @property
    def index(self):
        return self.app.books.index(self)

    def close(self):
        self.app.books.books.remove(self)
        del self.app.books.books_by_filename[self.filename]

    def save(self, path=None):
        if self.path:
            del self.app.books[self.filename]
            self.filename = path
            self.app.books[self.filename] = self
        self.api.save(self.filename)


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


the_app = App()
