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


class TestShape(TestBase):
    def test_name(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)

        sh = self.wb1.sheets[0].shapes[0]
        assert_equal(sh.name, 'pic1')
        sh.name = "yoyoyo"
        assert_equal(sh.name, 'yoyoyo')

    def test_coordinates(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(name='pic1', filename=filename, left=0, top=0, width=200, height=100)

        sh = self.wb1.sheets[0].shapes[0]
        for a, init, neu in (('left', 0, 50), ('top', 0, 50), ('width', 200, 150), ('height', 100, 160)):
            assert_equal(getattr(sh, a), init)
            setattr(sh, a, neu)
            assert_equal(getattr(sh, a), neu)

    def test_picture_object(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)

        assert_equal(self.wb1.sheets[0].shapes[0], self.wb1.sheets[0].shapes['pic1'])

    def test_delete(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_true('pic1' in self.wb1.sheets[0].shapes)
        self.wb1.sheets[0].shapes[0].delete()
        assert_false('pic1' in self.wb1.sheets[0].shapes)

    def test_type(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(name='pic1', filename=filename)
        assert_equal(self.wb1.sheets[0].shapes[0].type, 'picture')


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