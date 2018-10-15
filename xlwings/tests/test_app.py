# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys
import time
import unittest

import xlwings as xw
from xlwings.tests.common import TestBase, this_dir, SPEC


class TestApps(TestBase):
    def test_active(self):
        self.assertTrue(xw.apps.active in [self.app1, self.app2])

    def test_len(self):
        n_original = len(xw.apps)
        app = xw.App(spec=SPEC)
        wb = app.books.add()
        self.assertEqual(n_original + 1, len(xw.apps))
        app.quit()

    def test_count(self):
        self.assertEqual(xw.apps.count, len(xw.apps))

    def test_iter(self):
        for app in xw.apps:
            if app == (self.app1 or self.app2):
                self.assertEqual(len(app.books), 2)

    def test_keys(self):
        k = xw.apps.keys()[0]
        self.assertEqual(xw.apps[k], xw.apps(k))


class TestApp(TestBase):
    def test_activate(self):
        if sys.platform.startswith('win') and self.app1.version.major > 14:
            # Excel >= 2013 on Win has issues with activating hidden apps correctly
            # over two instances
            with self.assertRaises(Exception):
                self.app1.activate()
        else:
            self.assertEqual(self.app2, xw.apps.active)
            self.app1.activate()
            self.assertEqual(self.app1, xw.apps.active)

    def test_visible(self):
        # Can't successfully test for False on Mac...?
        self.app1.visible = True
        self.assertTrue(self.app1.visible)

    def test_quit(self):
        app = xw.App()
        n_apps = len(xw.apps)
        app.quit()
        time.sleep(1)  # needed for Mac Excel 2011
        self.assertEqual(n_apps - 1, len(xw.apps))

    def test_kill(self):
        app = xw.App(spec=SPEC)
        n_apps = len(xw.apps)
        app.kill()
        import time
        time.sleep(0.5)
        self.assertEqual(n_apps - 1, len(xw.apps))

    def test_screen_updating(self):
        self.app1.screen_updating = False
        self.assertEqual(self.app1.screen_updating, False)

        self.app1.screen_updating = True
        self.assertTrue(self.app1.screen_updating)

    def test_display_alerts(self):
        self.app1.display_alerts = False
        self.assertEqual(self.app1.display_alerts, False)

        self.app1.display_alerts = True
        self.assertTrue(self.app1.display_alerts)

    def test_calculation_calculate(self):
        sht = self.wb1.sheets[0]
        sht.range('A1').value = 2
        sht.range('B1').formula = '=A1 * 2'

        self.app1.calculation = 'manual'
        sht.range('A1').value = 4
        self.assertEqual(sht.range('B1').value, 4)

        self.app1.calculation = 'automatic'
        self.app1.calculate()  # This is needed on Mac Excel 2016 but not on Mac Excel 2011 (changed behaviour)
        self.assertEqual(sht.range('B1').value, 8)

        sht.range('A1').value = 2
        self.assertEqual(sht.range('B1').value, 4)

    def test_calculation(self):
        self.app1.calculation = 'automatic'
        self.assertEqual(self.app1.calculation, 'automatic')

        self.app1.calculation = 'manual'
        self.assertEqual(self.app1.calculation, 'manual')

        self.app1.calculation = 'semiautomatic'
        self.assertEqual(self.app1.calculation, 'semiautomatic')

    def test_version(self):
        self.assertTrue(self.app1.version.major > 0)

    def test_wb_across_instances(self):
        app1_wb_count = len(self.app1.books)
        app2_wb_count = len(self.app2.books)

        wb2 = self.app1.books.add()
        wb3 = self.app2.books.add()
        wb4 = self.app2.books.add()
        wb5 = self.app2.books.add()

        self.assertEqual(len(self.app1.books), app1_wb_count + 1)
        self.assertEqual(len(self.app2.books), app2_wb_count + 3)

        wb2.close()
        wb3.close()
        wb4.close()
        wb5.close()

    def test_selection(self):
        self.assertEqual(self.app1.selection.address, '$A$1')

    def test_books(self):
        self.assertEqual(len(self.app2.books), 2)

    def test_pid(self):
        self.assertTrue(self.app1.pid > 0)

    def test_len(self):
        n_books = len(self.app1.books)
        self.app1.books.add()
        self.assertEqual(len(self.app1.books), n_books + 1)

    def test_macro(self):
        wb = self.app1.books.open(os.path.join(this_dir, 'macro book.xlsm'))
        test1 = self.app1.macro('Module1.Test1')
        res1 = test1('Test1a', 'Test1b')
        self.assertEqual(res1, 1)


if __name__ == '__main__':
    unittest.main()
