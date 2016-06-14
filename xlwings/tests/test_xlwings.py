# -*- coding: utf-8 -*-
# TODO: clean up used workbooks

from __future__ import unicode_literals
import os
import sys
import shutil
from datetime import datetime, date

import pytz
import inspect
import nose
from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

from xlwings import Application, Book, Sheet, Range, Chart, Picture, Plot, ShapeAlreadyExists
from xlwings.constants import ChartType, RgbColor
from .test_data import data, test_date_1, test_date_2, list_row_1d, list_row_2d, list_col, chart_data
from .common import TestBase, _skip_if_no_matplotlib, _skip_if_no_numpy, _skip_if_no_pandas


this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))

# Mac imports
if sys.platform.startswith('darwin'):
    from appscript import k as kw
    # TODO: uncomment the desired Excel installation or set to None for default installation
    # APP_TARGET = None
    APP_TARGET = '/Applications/Microsoft Office 2011/Microsoft Excel'
else:
    APP_TARGET = None

# Optional dependencies
try:
    import numpy as np
    from numpy.testing import assert_array_equal
    from .test_data import array_1d, array_2d
except ImportError:
    np = None
try:
    import pandas as pd
    from pandas import DataFrame, Series
    from pandas.util.testing import assert_frame_equal, assert_series_equal
    from .test_data import series_1, timeseries_1, df_1, df_2, df_dateindex, df_multiheader, df_multiindex
except ImportError:
    pd = None
try:
    import matplotlib
    from matplotlib.figure import Figure
except ImportError:
    matplotlib = None
try:
    import PIL
except ImportError:
    PIL = None


class TestPicture(TestBase):
    def setUp(self):
        super(TestPicture, self).setUp('test_chart_1.xlsx')

    def test_two_wkb(self):
        wb2 = Book(app_visible=False, app_target=APP_TARGET)
        pic1 = Picture.add(sheet=1, name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        pic2 = Picture.add(sheet=self.wb.sheet(1), name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic1.name, 'pic1')
        assert_equal(pic2.name, 'pic1')
        wb2.close()

    def test_name(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.name, 'pic1')

        pic.name = 'pic_new'
        assert_equal(pic.name, 'pic_new')

    def test_left(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.left, 0)

        pic.left = 20
        assert_equal(pic.left, 20)

    def test_top(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.left, 0)

        pic.top = 20
        assert_equal(pic.top, 20)

    def test_width(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        if PIL:
            assert_equal(pic.width, 60)
        else:
            assert_equal(pic.width, 100)

        pic.width = 50
        assert_equal(pic.width, 50)

    def test_picture_object(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.name, Picture('pic1').name)

    def test_height(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        if PIL:
            assert_equal(pic.height, 60)
        else:
            assert_equal(pic.height, 100)

        pic.height = 50
        assert_equal(pic.height, 50)

    @raises(Exception)
    def test_delete(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        pic.delete()
        pic.name

    @raises(ShapeAlreadyExists)
    def test_duplicate(self):
        pic1 = Picture.add(os.path.join(this_dir, 'sample_picture.png'), name='pic1')
        pic2 = Picture.add(os.path.join(this_dir, 'sample_picture.png'), name='pic1')

    def test_picture_update(self):
        pic1 = Picture.add(os.path.join(this_dir, 'sample_picture.png'), name='pic1')
        pic1.update(os.path.join(this_dir, 'sample_picture.png'))


class TestPlot(TestBase):
    def setUp(self):
        super(TestPlot, self).setUp('test_chart_1.xlsx')

    def test_add_plot(self):
        _skip_if_no_matplotlib()

        fig = Figure(figsize=(8, 6))
        ax = fig.add_subplot(111)
        ax.plot([1, 2, 3, 4, 5])

        plot = Plot(fig)
        pic = plot.show('Plot1')
        assert_equal(pic.name, 'Plot1')

        plot.show('Plot2', sheet=2)
        pic2 = Picture(2, 'Plot2')
        assert_equal(pic2.name, 'Plot2')


class TestChart(TestBase):

    def setUp(self):
        super(TestChart, self).setUp('test_chart_1.xlsx')

    def test_add_keywords(self):
        name = 'My Chart'
        chart_type = ChartType.xlLine
        Range('A1').value = chart_data
        chart = Chart.add(chart_type=chart_type, name=name, source_data=Range('A1').table)

        chart_actual = Chart(name)
        name_actual = chart_actual.name
        chart_type_actual = chart_actual.chart_type
        assert_equal(name, name_actual)
        if sys.platform.startswith('win'):
            assert_equal(chart_type, chart_type_actual)
        else:
            assert_equal(kw.line_chart, chart_type_actual)

    def test_add_properties(self):
        name = 'My Chart'
        chart_type = ChartType.xlLine
        Range('Sheet2', 'A1').value = chart_data
        chart = Chart.add('Sheet2')
        chart.chart_type = chart_type
        chart.name = name
        chart.set_source_data(Range('Sheet2', 'A1').table)

        chart_actual = Chart('Sheet2', name)
        name_actual = chart_actual.name
        chart_type_actual = chart_actual.chart_type
        assert_equal(name, name_actual)
        if sys.platform.startswith('win'):
            assert_equal(chart_type, chart_type_actual)
        else:
            assert_equal(kw.line_chart, chart_type_actual)


if __name__ == '__main__':
    nose.main()
