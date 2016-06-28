# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os

from nose.tools import assert_equal, assert_true

import xlwings as xw
from xlwings.constants import Calculation
from .common import TestBase, this_dir


class TestApps(TestBase):
    def test_active(self):
        assert_equal(xw.apps[0], xw.apps.active)

    def test_len_apps(self):
        n_original = len(xw.apps)
        app = xw.App()
        wb = app.book()
        assert_equal(n_original + 1, len(xw.apps))
        assert_equal(xw.apps[0], app)
        app.quit()

    def test_app_iter(self):
        for app in xw.apps:
            assert_equal(len(app.books), 1)


class TestApp(TestBase):
    def test_activate(self):
        assert_equal(self.app2, xw.apps.active)
        self.app1.activate()
        assert_equal(self.app1, xw.apps.active)

    def test_visible(self):
        # Can't successfully test for False on Mac...?
        self.app1.visible = True
        assert_true(self.app1.visible)

    def test_quit(self):
        n_apps = len(xw.apps)
        while len(self.app2.books) > 0:
            self.app2.books[0].close()
        self.app2.quit()
        assert_equal(n_apps - 1, len(xw.apps))

    def test_kill(self):
        app = xw.App()
        n_apps = len(xw.apps)
        app.kill()
        assert_equal(n_apps - 1, len(xw.apps))

    def test_screen_updating(self):
        self.app1.screen_updating = False
        assert_equal(self.app1.screen_updating, False)

        self.app1.screen_updating = True
        assert_true(self.app1.screen_updating)

    def test_calculation_calculate(self):
        sht = self.wb1.sheets[0]
        sht.range('A1').value = 2
        sht.range('B1').formula = '=A1 * 2'

        self.app1.calculation = Calculation.xlCalculationManual
        sht.range('A1').value = 4
        assert_equal(sht.range('B1').value, 4)

        self.app1.calculation = Calculation.xlCalculationAutomatic
        self.app1.calculate()  # This is needed on Mac Excel 2016 but not on Mac Excel 2011 (changed behaviour)
        assert_equal(sht.range('B1').value, 8)

        sht.range('A1').value = 2
        assert_equal(sht.range('B1').value, 4)

    def test_version(self):
        assert_true(int(self.app1.version.split('.')[0]) > 0)

    def test_wb_across_instances(self):
        app1_wb_count = len(self.app1.books)
        app2_wb_count = len(self.app2.books)

        wb2 = self.app1.book()
        wb3 = self.app2.book()
        wb4 = self.app2.book()
        wb5 = self.app2.book()

        assert_equal(len(self.app1.books), app1_wb_count + 1)
        assert_equal(len(self.app2.books), app2_wb_count + 3)

        wb2.close()
        wb3.close()
        wb4.close()
        wb5.close()

        self.app2.quit()

    def test_selection(self):
        assert_equal(self.app1.selection.address, '$A$1')

    def test_books(self):
        assert_equal(len(self.app2.books), 1)

    def test_pid(self):
        assert_true(self.app1.pid > 0)

    def test_range(self):
        n_books = len(self.app1.books)
        self.app1.book()
        assert_equal(len(self.app1.books), n_books + 1)

    def test_macro(self):
        self.app1.book(os.path.join(this_dir, 'macro book.xlsm'))

        test1 = self.app1.macro('Module1.Test1')
        res1 = test1('Test1a', 'Test1b')
        assert_equal(res1, 1)
