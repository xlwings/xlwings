# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from nose.tools import assert_equal, assert_not_equal, assert_true, raises, assert_false

import xlwings as xw
from .common import TestBase


class TestSheets(TestBase):
    def test_active(self):
        self.wb1.sheets[2].activate()
        assert_equal(self.wb1.sheets.active.name, self.wb1.sheets[2].name)

    def test_index(self):
        assert_equal(self.wb1.sheets[0].name, self.wb1.sheets(1).name)

    def test_len(self):
        assert_equal(len(self.wb1.sheets), 3)

    def del_sheet(self):
        name = self.wb1.sheets[0].name
        del self.wb1.sheets[0]
        assert_equal(len(self.wb1.sheets), 2)
        assert_false(self.wb1.sheets[0].name, name)

    def test_iter(self):
        for ix, sht in enumerate(self.wb1.sheets):
            assert_equal(self.wb1.sheets[ix].name, sht.name)

    def test_add(self):
        self.wb1.sheets.add()
        assert_equal(len(self.wb1.sheets), 4)

    def test_add_before(self):
        new_sheet = self.wb1.sheets.add(before='Sheet1')
        assert_equal(self.wb1.sheets[0].name, new_sheet.name)

    def test_add_after(self):
        self.wb1.sheets.add(after=len(self.wb1.sheets))
        assert_equal(self.wb1.sheets[(len(self.wb1.sheets) - 1)].name, self.wb1.sheets.active.name)

        self.wb1.sheets.add(after=1)
        assert_equal(self.wb1.sheets[1].name, self.wb1.sheets.active.name)

    def test_add_default(self):
        current_index = self.wb1.sheets.active.index
        self.wb1.sheets.add()
        assert_equal(self.wb1.sheets.active.index, current_index)

    def test_add_named(self):
        self.wb1.sheets.add('test', before=1)
        assert_equal(self.wb1.sheets[0].name, 'test')

    @raises(Exception)
    def test_add_name_already_taken(self):
        self.wb1.sheets.add('Sheet1')


class TestSheet(TestBase):
    def test_name(self):
        self.wb1.sheets[0].name = 'NewName'
        assert_equal(self.wb1.sheets[0].name, 'NewName')

    def test_names(self):
        self.wb1.sheets[0].range('A1').name = 'test1'
        assert_equal(len(self.wb1.sheets[0].names), 1)

    def test_book(self):
        assert_equal(self.wb1.sheets[0].book.name, self.wb1.name)

    def test_index(self):
        assert_equal(self.wb1.sheets['Sheet1'].index, 1)

    def test_range(self):
        self.wb1.sheets[0].range('A1').value = 123.
        assert_equal(self.wb1.sheets[0].range('A1').value, 123.)

    def test_cells(self):
        pass  # TODO

    def test_activate(self):
        self.wb1.sheets['Sheet2'].activate()
        assert_equal(self.wb1.sheets.active.name, 'Sheet2')
        self.wb1.sheets[2].activate()
        assert_equal(self.wb1.sheets.active.index, 3)
        self.wb1.sheets(1).activate()
        assert_equal(self.wb1.sheets.active.index, 1)

    def test_clear_content(self):
        self.wb1.sheets['Sheet2'].range('G10').value = 22
        self.wb1.sheets['Sheet2'].clear_contents()
        assert_equal(self.wb1.sheets['Sheet2'].range('G10').value, None)

    def test_clear(self):
        self.wb1.sheets['Sheet2'].range('G10').value = 22
        self.wb1.sheets['Sheet2'].range('G10').color = (255, 255, 255)
        self.wb1.sheets['Sheet2'].clear()
        assert_equal(self.wb1.sheets['Sheet2'].range('G10').value, None)
        assert_equal(self.wb1.sheets['Sheet2'].range('G10').color, None)

    def test_autofit(self):
        sht = self.wb1.sheets['Sheet1']
        sht.range('A1:D4').value = 'test_string'
        sht.range('A1:D4').row_height = 40
        sht.range('A1:D4').column_width = 40
        assert_equal(sht.range('A1:D4').row_height, 40)
        assert_equal(sht.range('A1:D4').column_width, 40)

        sht.autofit()

        assert_not_equal(sht.range('A1:D4').row_height, 40)
        assert_not_equal(sht.range('A1:D4').column_width, 40)

        # Just checking if they don't throw an error
        sht.autofit('r')
        sht.autofit('c')
        sht.autofit('rows')
        sht.autofit('columns')

    def test_delete(self):
        assert_true('Sheet1' in [i.name for i in self.wb1.sheets])
        self.wb1.sheets['Sheet1'].delete()
        assert_false('Sheet1' in [i.name for i in self.wb1.sheets])

    def test_charts(self):
        pass  # TODO

    def test_chart(self):
        pass  # TODO

    def test_shapes(self):
        pass  # TODO

    def test_shape(self):
        pass  # TODO

    def test_pictures(self):
        pass  # TODO

    def test_picture(self):
        pass  # TODO