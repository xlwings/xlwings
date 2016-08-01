# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import inspect
import unittest

import xlwings as xw

this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))

# Uncomment to run tests in Mac Excel 2011
SPEC = None
# SPEC = '/Applications/Microsoft Office 2011/Microsoft Excel'


class TestBase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app1 = xw.App(visible=False, spec=SPEC)
        cls.app2 = xw.App(visible=False, spec=SPEC)

    def setUp(self):
        self.wb1 = self.app1.books.add()
        self.wb2 = self.app2.books.add()
        for wb in [self.wb1, self.wb2]:
            if len(wb.sheets) == 1:
                wb.sheets.add(after=1)
                wb.sheets.add(after=2)
                wb.sheets[0].select()

    def tearDown(self):
        for app in [self.app1, self.app2]:
            app.books[-1].close()

    @classmethod
    def tearDownClass(cls):
        cls.app1.kill()
        cls.app2.kill()
