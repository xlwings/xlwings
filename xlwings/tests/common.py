# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import inspect
import unittest

import nose

import xlwings as xw

this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))

# Optional dependencies
try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None
try:
    import matplotlib
except ImportError:
    matplotlib = None
try:
    import PIL
except ImportError:
    PIL = None


# Test skips and fixtures
def _skip_if_no_numpy():
    if np is None:
        raise nose.SkipTest('numpy missing')


def _skip_if_no_pandas():
    if pd is None:
        raise nose.SkipTest('pandas missing')


def _skip_if_no_matplotlib():
    if matplotlib is None:
        raise nose.SkipTest('matplotlib missing')

# Uncomment to run tests in Mac Excel 2011
SPEC = None
# SPEC = '/Applications/Microsoft Office 2011/Microsoft Excel'


class TestBase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.existing_apps = list(xw.apps)
        cls.app1 = xw.App(visible=True, spec=SPEC)
        cls.app2 = xw.App(visible=True, spec=SPEC)

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
            while len(app.books) > 0:
                app.books[0].close()

    @classmethod
    def tearDownClass(cls):
        cls.app1.books.add()
        cls.app2.books.add()
        for app in [cls.app1, cls.app2]:
            app.quit()
