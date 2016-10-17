# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import unittest

import xlwings as xw
from xlwings.tests.common import TestBase, this_dir
from xlwings import PY3

try:
    import matplotlib.pyplot as plt
except ImportError:
    plt = None


class TestActive(TestBase):
    def test_apps_active(self):
        self.assertEqual(xw.apps.active, self.app2)

    def test_books_active(self):
        wb = xw.Book()
        self.assertEqual(xw.books.active, wb)

    def test_sheets_active(self):
        self.wb2.sheets[0].name = 'active sheet test'
        self.assertEqual(self.wb2.sheets.active.name, 'active sheet test')

    def test_range(self):
        xw.sheets.active.range('B2:C3').value = 123.
        self.assertEqual(xw.Range('B2:C3').value, [[123., 123.], [123., 123.]])

    def test_book_fullname_closed(self):
        wb = xw.Book(os.path.join(this_dir, 'test book.xlsx'))
        self.assertEqual(wb, xw.apps.active.books['test book.xlsx'])

    def test_book_fullname_open(self):
        wb1 = self.app1.books.open(os.path.join(this_dir, 'test book.xlsx'))
        wb2 = xw.Book(os.path.join(this_dir, 'test book.xlsx'))
        self.assertEqual(wb1, wb2)

    def test_book_name_closed(self):
        os.chdir(this_dir)
        wb = xw.Book('test book.xlsx')
        self.assertEqual(wb, xw.apps.active.books['test book.xlsx'])

    def test_book_name_open(self):
        wb1 = self.app1.books.open(os.path.join(this_dir, 'test book.xlsx'))
        wb2 = xw.Book('test book.xlsx')
        self.assertEqual(wb1, wb2)
    
    def test_book_open_bad_name(self):        
        if PY3:
            with self.assertRaises(FileNotFoundError):
                xw.Book('bad name.xlsx')
        else:
            with self.assertRaises(IOError):
                xw.Book('bad name.xlsx')

    def test_book(self):
        wb = xw.Book()
        self.assertEqual(wb, xw.apps.active.books.active)

    def test_book_name_unsaved(self):
        wb = xw.Book()
        self.assertEqual(wb, xw.Book(wb.name))

    def test_books(self):
        wb = xw.Book()
        self.assertEqual(xw.books[-1], wb)


class TestView(TestBase):
    def test_list_new_book(self):
        n_books = xw.books.count
        xw.view([1, 2, 3])
        self.assertEqual(xw.books.count, n_books + 1)

    def test_list_sheet(self):
        n_books = xw.books.count
        xw.view([1, 2, 3], sheet=xw.books[0].sheets[0])
        self.assertEqual(xw.books.count, n_books)
        self.assertEqual(xw.books[0].sheets[0].range('A1:C1').value, [1., 2., 3.])


if __name__ == '__main__':
    unittest.main()
