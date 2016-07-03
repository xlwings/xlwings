# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys

from nose.tools import assert_equal, assert_not_equal, assert_true, raises, assert_false

import xlwings as xw
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
        assert_equal(pic.name, self.wb1.sheets[0].pictures['pic1'].name)

    def test_height(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(pic.height, 60)

        pic.height = 50
        assert_equal(pic.height, 50)

    def test_delete(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_true('pic1' in self.wb1.sheets[0].pictures)
        pic.delete()
        assert_false('pic1' in self.wb1.sheets[0].pictures)

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


class TestCharts(TestBase):
    def test_add_properties(self):
        sht = self.wb1.sheets[0]
        sht.range('A1').value = [['one', 'two'], [1.1, 2.2]]

        self.assertEqual(len(sht.charts), 0)
        chart = sht.charts.add()
        self.assertEqual(len(sht.charts), 1)

        chart.name = 'My Chart'
        chart.source_data = sht.range('A1').expand('table')
        chart.chart_type = 'line'

        self.assertEqual('My Chart', chart.name)
        self.assertEqual(sht.charts[0].chart_type, 'line')

        chart.chart_type = 'pie'
        self.assertEqual(sht.charts[0].chart_type, 'pie')

        for a in ('left', 'top', 'width', 'height'):
            setattr(chart, a, 400)
            self.assertEqual(getattr(sht.charts[0], a), 400)
            setattr(sht.charts[0], a, 500)
            self.assertEqual(getattr(chart, a), 500)

        chart.delete()
        self.assertEqual(sht.charts.count, 0)


class TestChart(TestBase):
    def test_len(self):
        chart = self.wb1.sheets[0].charts.add()
        assert_equal(len(self.wb1.sheets[0].charts), 1)

    def test_count(self):
        chart = self.wb1.sheets[0].charts.add()
        assert_equal(len(self.wb1.sheets[0].charts), self.wb1.sheets[0].charts.count)