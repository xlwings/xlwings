# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os

from nose.tools import assert_equal

import xlwings as xw
from .common import TestBase, this_dir


class TestActive(TestBase):
    def test_apps_active(self):
        assert_equal(xw.apps.active, self.app2)

    def test_books_active(self):
        wb = xw.Book()
        assert_equal(xw.books.active, wb)

    def test_sheets_active(self):
        self.wb2.sheets[0].name = 'active sheet test'
        assert_equal(self.wb2.sheets.active.name, 'active sheet test')

    def test_range(self):
        xw.sheets.active.range('B2:C3').value = 123.
        assert_equal(xw.Range('B2:C3').value, [[123., 123.], [123., 123.]])

    def test_book_fullname_closed(self):
        wb = xw.Book(os.path.join(this_dir, 'test book.xlsx'))
        assert_equal(wb, self.app2.books['test book.xlsx'])

    def test_book_fullname_open(self):
        wb1 = self.app1.books.open(os.path.join(this_dir, 'test book.xlsx'))
        wb2 = xw.Book(os.path.join(this_dir, 'test book.xlsx'))
        assert_equal(wb1, wb2)

    def test_book_name_closed(self):
        os.chdir(this_dir)
        wb = xw.Book('test book.xlsx')
        assert_equal(wb, xw.apps.active.books['test book.xlsx'])

    def test_book_name_open(self):
        wb1 = self.app1.books.open(os.path.join(this_dir, 'test book.xlsx'))
        wb2 = xw.Book('test book.xlsx')
        assert_equal(wb1, wb2)

    def test_book(self):
        wb = xw.Book()
        assert_equal(wb, self.app2.books.active)

    def test_book_name_unsaved(self):
        wb = xw.Book()
        assert_equal(wb, xw.Book(wb.name))

    def test_books(self):
        wb = xw.Book()
        assert_equal(xw.books[-1], wb)
