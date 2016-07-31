# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import unittest

import xlwings as xw
from xlwings.tests.common import TestBase, this_dir

try:
    import matplotlib as mpl
    import matplotlib.pyplot as plt
except ImportError:
    mpl = None

try:
    import PIL
except ImportError:
    PIL = None


class TestShape(TestBase):
    def test_name(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(filename, name='pic1')

        sh = self.wb1.sheets[0].shapes[0]
        self.assertEqual(sh.name, 'pic1')
        sh.name = "yoyoyo"
        self.assertEqual(sh.name, 'yoyoyo')

    def test_coordinates(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(filename, name='pic1', left=0, top=0, width=200, height=100)

        sh = self.wb1.sheets[0].shapes[0]
        for a, init, neu in (('left', 0, 50), ('top', 0, 50), ('width', 200, 150), ('height', 100, 160)):
            self.assertEqual(getattr(sh, a), init)
            setattr(sh, a, neu)
            self.assertEqual(getattr(sh, a), neu)

    def test_picture_object(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(filename, name='pic1')

        self.assertEqual(self.wb1.sheets[0].shapes[0], self.wb1.sheets[0].shapes['pic1'])

    def test_delete(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertTrue('pic1' in self.wb1.sheets[0].shapes)
        self.wb1.sheets[0].shapes[0].delete()
        self.assertFalse('pic1' in self.wb1.sheets[0].shapes)

    def test_type(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(self.wb1.sheets[0].shapes[0].type, 'picture')


class TestPicture(TestBase):
    def test_two_books(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        pic2 = self.wb2.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(pic1.name, 'pic1')
        self.assertEqual(pic2.name, 'pic1')

    def test_name(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(pic.name, 'pic1')

        pic.name = 'pic_new'
        self.assertEqual(pic.name, 'pic_new')

    def test_left(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(pic.left, 0)

        pic.left = 20
        self.assertEqual(pic.left, 20)

    def test_top(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(pic.left, 0)

        pic.top = 20
        self.assertEqual(pic.top, 20)

    def test_width(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        if PIL:
            self.assertEqual(pic.width, 60)
        else:
            self.assertEqual(pic.width, 100)

        pic.width = 50
        self.assertEqual(pic.width, 50)

    def test_picture_object(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(pic.name, self.wb1.sheets[0].pictures['pic1'].name)

    def test_height(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        if PIL:
            self.assertEqual(pic.height, 60)
        else:
            self.assertEqual(pic.height, 100)

        pic.height = 50
        self.assertEqual(pic.height, 50)

    def test_delete(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertTrue('pic1' in self.wb1.sheets[0].pictures)
        pic.delete()
        self.assertFalse('pic1' in self.wb1.sheets[0].pictures)

    def test_duplicate(self):
        with self.assertRaises(xw.ShapeAlreadyExists):
            filename = os.path.join(this_dir, 'sample_picture.png')
            pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1')
            pic2 = self.wb1.sheets[0].pictures.add(filename, name='pic1')

    def test_picture_update(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        pic1.update(filename)

    def test_picture_auto_update(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1', update=True)
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1', update=True)
        self.assertEqual(len(self.wb1.sheets[0].pictures), 1)

    def test_picture_auto_update_without_name(self):
        with self.assertRaises(ValueError):
            filename = os.path.join(this_dir, 'sample_picture.png')
            pic1 = self.wb1.sheets[0].pictures.add(filename, update=True)

    def test_picture_index(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        self.assertEqual(self.wb1.sheets[0].pictures[0], self.wb1.sheets[0].pictures['pic1'])
        self.assertEqual(self.wb1.sheets[0].pictures(1), self.wb1.sheets[0].pictures[0])

    def test_len(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic1')
        pic2 = self.wb1.sheets[0].pictures.add(filename, name='pic2')
        self.assertEqual(len(self.wb1.sheets[0].pictures), 2)

    def test_iter(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        names = ['pic1', 'pic2']
        pic1 = self.wb1.sheets[0].pictures.add(filename, name=names[0])
        pic2 = self.wb1.sheets[0].pictures.add(filename, name=names[1])
        for ix, pic in enumerate(self.wb1.sheets[0].pictures):
            self.assertEqual(self.wb1.sheets[0].pictures[ix].name, names[ix])

    def test_contains(self):
        filename = os.path.join(this_dir, 'sample_picture.png')
        pic1 = self.wb1.sheets[0].pictures.add(filename, name='pic 1')
        self.assertTrue('pic 1' in self.wb1.sheets[0].pictures)


@unittest.skipIf(mpl is None, 'matplotlib missing')
class TestMatplotlib(TestBase):
    def test_add_no_name(self):
        fig = plt.figure()
        plt.plot([-1, 1, -2, 2, -3, 3, 2])
        self.wb1.sheets[0].pictures.add(fig)
        self.assertEqual(len(self.wb1.sheets[0].pictures), 1)

    def test_add_with_name(self):
        fig = plt.figure()
        plt.plot([-1, 1, -2, 2, -3, 3, 2])
        self.wb1.sheets[0].pictures.add(fig, name='Test1')
        self.assertEqual(self.wb1.sheets[0].pictures[0].name, 'Test1')


class TestCharts(TestBase):
    def test_add_properties(self):
        sht = self.wb1.sheets[0]
        sht.range('A1').value = [['one', 'two'], [1.1, 2.2]]

        self.assertEqual(len(sht.charts), 0)
        chart = sht.charts.add()
        self.assertEqual(len(sht.charts), 1)

        chart.name = 'My Chart'
        chart.set_source_data(sht.range('A1').expand('table'))
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
        self.assertEqual(len(self.wb1.sheets[0].charts), 1)

    def test_count(self):
        chart = self.wb1.sheets[0].charts.add()
        self.assertEqual(len(self.wb1.sheets[0].charts), self.wb1.sheets[0].charts.count)

if __name__ == '__main__':
    unittest.main()
