# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import nose
from nose.tools import assert_equal
from datetime import datetime
from xlwings import Workbook, Range

# Optional imports
try:
    import numpy as np
    from numpy.testing import assert_array_equal
except ImportError:
    np = None
try:
    import pandas as pd
    from pandas import DataFrame, Series
    from pandas.util.testing import assert_frame_equal, assert_series_equal
except ImportError:
    pd = None

# Connect to test file and make Sheet1 the active sheet
xl_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_chart_1.xlsx')
wb = Workbook(xl_file)
wb.activate('Sheet1')

# Test Data


if __name__ == '__main__':
    nose.main()