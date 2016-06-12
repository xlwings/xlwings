# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import inspect

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


class TestBase:
    @classmethod
    def setUpClass(cls):
        cls.existing_apps = list(xw.apps)
        cls.app = xw.Application(visible=False)

    def setUp(self, xlsx=None):
        if len(self.app.workbooks) == 0:
            self.wb = self.app.workbook()
        else:
            self.wb = self.app.workbooks[0]
        if len(self.wb.sheets) == 1:
            self.wb.sheets.add(after=1)
            self.wb.sheets.add(after=2)

    def tearDown(self):
        for app in xw.applications:
            if app.pid not in [i.pid for i in self.existing_apps]:
                for wb in app:
                    wb.close()

    @classmethod
    def tearDownClass(cls):
        for app in xw.applications:
            if app.pid not in [i.pid for i in cls.existing_apps]:
                app.kill()
