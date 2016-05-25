# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from nose.tools import assert_equal, assert_not_equal, assert_true, raises, assert_false

import xlwings as xw
from .common import TestBase


class TestSheet(TestBase):
    def test_activate(self):
        self.wb['Sheet2'].activate()
        assert_equal(xw.active.sheet.name, 'Sheet2')
        self.wb[2].activate()
        assert_equal(xw.active.sheet.index, 3)

    def test_name(self):
        self.wb[0].name = 'NewName'
        assert_equal(self.wb[0].name, 'NewName')

    def test_index(self):
        assert_equal(self.wb['Sheet1'].index, 1)

    def test_clear_content_active_sheet(self):
        self.wb[0].range('G10').value = 22
        xw.active.sheet.clear_contents()
        assert_equal(self.wb[0].range('G10').value, None)

    def test_clear_active_sheet(self):
        self.wb[0].range('G10').value = 22
        xw.active.sheet.clear()
        assert_equal(self.wb[0].range('G10').value, None)

    def test_clear_content(self):
        self.wb['Sheet2'].range('G10').value = 22
        self.wb['Sheet2'].clear_contents()
        assert_equal(self.wb['Sheet2'].range('G10').value, None)

    def test_clear(self):
        self.wb['Sheet2'].range('G10').value = 22
        self.wb['Sheet2'].clear()
        assert_equal(self.wb['Sheet2'].range('G10').value, None)

    def test_autofit(self):
        sht = self.wb['Sheet1']
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

    def test_add_before(self):
        new_sheet = self.wb.sheets.add(before='Sheet1')
        assert_equal(self.wb[0].name, new_sheet.name)

    def test_add_after(self):
        self.wb.sheets.add(after=len(self.wb.sheets))
        assert_equal(self.wb[(len(self.wb.sheets) - 1)].name, xw.active.sheet.name)

        self.wb.sheets.add(after=1)
        assert_equal(self.wb[1].name, xw.active.sheet.name)

    def test_add_default(self):
        current_index = xw.active.sheet.index
        self.wb.sheets.add()
        assert_equal(xw.active.sheet.index, current_index)

    def test_add_named(self):
        self.wb.sheets.add('test', before=1)
        assert_equal(self.wb[0].name, 'test')

    @raises(Exception)
    def test_add_name_already_taken(self):
        self.wb.sheets.add('Sheet1')

    def test_count(self):
        count = self.wb.sheets.count()
        assert_equal(count, 3)

    def test_sheets_names(self):
        all_names = [i.name for i in self.wb.sheets]
        assert_equal(all_names, ['Sheet1', 'Sheet2', 'Sheet3'])

    def test_delete(self):
        assert_true('Sheet1' in [i.name for i in self.wb.sheets])
        self.wb['Sheet1'].delete()
        assert_false('Sheet1' in [i.name for i in self.wb.sheets])