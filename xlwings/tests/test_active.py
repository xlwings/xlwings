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
        wb = xw.Book()
        wb.activate()
        assert_equal(xw.books.active.name, wb.name)

    def test_book_fullname(self):
        fullname = os.path.join(this_dir, 'test book.xlsx')
        wb = xw.Book(fullname)

    def test_sheets_active(self):
        self.wb2.activate()
        self.wb2.sheets[2].name = 'active sheet test'
        self.wb2.sheets[2].activate()
        assert_equal(self.wb2.sheets.active.name, 'active sheet test')

    def test_active_range(self):
        self.wb2.sheets[2].range('B2:C3').value = 123.
        self.wb2.activate()
        self.wb2.sheets[2].activate()
        assert_equal(xw.Range('B2:C3').value, [[123., 123.], [123., 123.]])

