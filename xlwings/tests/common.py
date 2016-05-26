# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys
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
    def setUp(self, xlsx=None):
        self.app = xw.Application(visible=False)
        self.wb = self.app.workbook()
        if len(self.wb.sheets) == 1:
            self.wb.sheets.add(after=1)
            self.wb.sheets.add(after=2)

    def tearDown(self):
        self.wb.close()
        if sys.platform.startswith('win'):
            self.app.kill()
