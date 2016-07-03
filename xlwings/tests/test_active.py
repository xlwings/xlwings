# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os

from nose.tools import assert_equal

import xlwings as xw
from .common import TestBase, this_dir


class TestActive(TestBase):
    def test_apps_active(self):
        self.wb1.activate()
        assert_equal(xw.apps.active, self.app1)

    def test_books_active(self):
        self.app1.activate()
        assert_equal(xw.books.active, self.wb1)

    def test_sheets_active(self):
        self.wb2.activate()
        self.wb2.sheets[2].name = 'active sheet test'
        self.wb2.sheets[2].activate()
        assert_equal(self.wb2.sheets.active.name, 'active sheet test')

    def test_range(self):
        self.wb2.sheets[2].range('B2:C3').value = 123.
        self.wb2.sheets[2].activate()
        assert_equal(xw.Range('B2:C3').value, [[123., 123.], [123., 123.]])

    def test_book_fullname(self):
        self.app1.activate()
        wb = xw.Book(os.path.join(this_dir, 'test book.xlsx'))
        assert_equal(wb, self.app1.books['test book.xlsx'])

    def test_book(self):
        self.app1.activate()
        wb = xw.Book()
        assert_equal(wb, self.app1.books.active)

    def test_books(self):
        self.app1.activate()
        wb = xw.Book()
        assert_equal(xw.books[1], wb)

