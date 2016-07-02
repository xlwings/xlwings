# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from nose.tools import assert_equal

import xlwings as xw
from .common import TestBase


class TestActive(TestBase):
    def test_active_app(self):
        self.wb1.activate()
        assert_equal(xw.apps.active, self.app1)

    def test_apps_active(self):
        self.wb1.activate()
        assert_equal(xw.apps.active, self.app1)

    def test_active_book(self):
        wb = xw.Book()
        wb.activate()
        assert_equal(xw.books.active.name, wb.name)

    def test_books_active(self):
        wb = xw.Book()
        wb.activate()
        assert_equal(xw.apps[0].books.active.name, wb.name)

    def test_active_sheet(self):
        self.wb2.activate()
        self.wb2.sheets[2].name = 'active sheet test'
        self.wb2.sheets[2].activate()
        assert_equal(xw.sheets.active.name, 'active sheet test')

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

