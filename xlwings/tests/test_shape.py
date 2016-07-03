# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys

from nose.tools import assert_equal, assert_not_equal, assert_true, raises, assert_false

import xlwings as xw
from xlwings.constants import ChartType
from .common import TestBase, this_dir, _skip_if_no_matplotlib


try:
    import matplotlib
    from matplotlib.figure import Figure
except ImportError:
    matplotlib = None


if sys.platform.startswith('darwin'):
    from appscript import k as kw


class TestPicture(TestBase):
    def test_two_books(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        pic2 = self.wb2.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic1.name, 'pic1')
        assert_equal(pic2.name, 'pic1')

    def test_name(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.name, 'pic1')

        pic.name = 'pic_new'
        assert_equal(pic.name, 'pic_new')

    def test_left(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.left, 0)

        pic.left = 20
        assert_equal(pic.left, 20)

    def test_top(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.left, 0)

        pic.top = 20
        assert_equal(pic.top, 20)

    def test_width(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.width, 60)

        pic.width = 50
        assert_equal(pic.width, 50)

    def test_picture_object(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.name, self.wb1.pictures['pic1'].name)

    def test_height(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.height, 60)

        pic.height = 50
        assert_equal(pic.height, 50)

    @raises(Exception)
    def test_delete(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_true('pic1' in [i.name for i in self.wb1.sheets[0].pictures])
        pic.delete()
        assert_false('pic1' in [i.name for i in self.wb1.sheets[0].pictures])

    @raises(xw.ShapeAlreadyExists)
    def test_duplicate(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        pic2 = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)

    def test_picture_update(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        pic1.update(filename)


class TestPlot(TestBase):
    def test_add_plot(self):
        _skip_if_no_matplotlib()

        fig = Figure(figsize=(8, 6))
        ax = fig.add_subplot(111)
        ax.plot([1, 2, 3, 4, 5])

        plot = Plot(fig)
        pic = plot.show('Plot1')
        assert_equal(pic.name, 'Plot1')

        plot.show('Plot2', sheet=2)
        pic2 = self.wb1.sheets[1].pictures['Plot2']
        assert_equal(pic2.name, 'Plot2')


class TestChart(TestBase):
    def test_add_keywords(self):
        self.wb1.sheets[0].range('A1').value = [['one', 'two'], [1.1, 2.2]]
        chart = self.wb1.sheets[0].charts.add(chart_type=ChartType.xlLine,
                                              name='My Chart',
                                              source_data=self.wb1.sheets[0].range('A1').expand('table'))

        assert_equal('My Chart', chart.name)
        if sys.platform.startswith('win'):
            assert_equal(self.wb1.sheets[0].charts[0].chart_type, ChartType.xlLine)
        else:
            assert_equal(self.wb1.sheets[0].charts[0].chart_type, kw.line_chart)

    def test_add_properties(self):
        self.wb1.sheets[1].range('A1').value = [['one', 'two'], [1.1, 2.2]]
        chart = self.wb1.sheets[1].charts.add()
        chart.chart_type = ChartType.xlLine
        chart.name = 'My Chart'
        chart.set_source_data(self.wb1.sheets[1].range('A1').expand('table'))

        assert_equal('My Chart', chart.name)
        if sys.platform.startswith('win'):
            assert_equal(self.wb1.sheets[0].charts[0].chart_type, ChartType.xlLine)
        else:
            assert_equal(self.wb1.sheets[0].charts[0].chart_type, kw.line_chart)