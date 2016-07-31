# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import unittest

import xlwings as xw
from xlwings.tests.common import TestBase


class TestSheets(TestBase):
    def test_active(self):
        self.assertEqual(self.wb2.sheets.active.name, self.wb2.sheets[0].name)

    def test_index(self):
        self.assertEqual(self.wb1.sheets[0].name, self.wb1.sheets(1).name)

    def test_len(self):
        self.assertEqual(len(self.wb1.sheets), 3)

    def del_sheet(self):
        name = self.wb1.sheets[0].name
        del self.wb1.sheets[0]
        self.assertEqual(len(self.wb1.sheets), 2)
        self.assertFalse(self.wb1.sheets[0].name, name)

    def test_iter(self):
        for ix, sht in enumerate(self.wb1.sheets):
            self.assertEqual(self.wb1.sheets[ix].name, sht.name)

    def test_add(self):
        self.wb1.sheets.add()
        self.assertEqual(len(self.wb1.sheets), 4)

    def test_add_before(self):
        new_sheet = self.wb1.sheets.add(before='Sheet1')
        self.assertEqual(self.wb1.sheets[0].name, new_sheet.name)

    def test_add_after(self):
        self.wb1.sheets.add(after=len(self.wb1.sheets))
        self.assertEqual(self.wb1.sheets[(len(self.wb1.sheets) - 1)].name, self.wb1.sheets.active.name)

        self.wb1.sheets.add(after=1)
        self.assertEqual(self.wb1.sheets[1].name, self.wb1.sheets.active.name)

    def test_add_default(self):
        current_index = self.wb1.sheets.active.index
        self.wb1.sheets.add()
        self.assertEqual(self.wb1.sheets.active.index, current_index)

    def test_add_named(self):
        self.wb1.sheets.add('test', before=1)
        self.assertEqual(self.wb1.sheets[0].name, 'test')

    def test_add_name_already_taken(self):
        with self.assertRaises(Exception):
            self.wb1.sheets.add('Sheet1')


class TestSheet(TestBase):
    def test_name(self):
        self.wb1.sheets[0].name = 'NewName'
        self.assertEqual(self.wb1.sheets[0].name, 'NewName')

    def test_names(self):
        self.wb1.sheets[0].range('A1').name = 'test1'
        self.assertEqual(len(self.wb1.sheets[0].names), 0)
        self.wb1.sheets[0].names.add('Sheet1!test2', 'Sheet1!B2')
        self.assertEqual(len(self.wb1.sheets[0].names), 1)

    def test_book(self):
        self.assertEqual(self.wb1.sheets[0].book.name, self.wb1.name)

    def test_index(self):
        self.assertEqual(self.wb1.sheets['Sheet1'].index, 1)

    def test_range(self):
        self.wb1.sheets[0].range('A1').value = 123.
        self.assertEqual(self.wb1.sheets[0].range('A1').value, 123.)

    def test_cells(self):
        pass  # TODO

    def test_activate(self):
        if sys.platform.startswith('win') and self.app1.version.major > 14:
            # Excel >= 2013 on Win has issues with activating hidden apps correctly
            # over two instances
            with self.assertRaises(Exception):
                self.app1.activate()
        else:
            self.wb2.activate()
            self.wb1.sheets['Sheet2'].activate()
            self.assertEqual(self.wb1.sheets.active.name, 'Sheet2')
            self.assertEqual(xw.apps[0], self.app1)
            self.wb1.sheets[2].activate()
            self.assertEqual(self.wb1.sheets.active.index, 3)
            self.wb1.sheets(1).activate()
            self.assertEqual(self.wb1.sheets.active.index, 1)

    def test_select(self):
        self.wb2.sheets[1].select()
        self.assertEqual(self.wb2.sheets.active, self.wb2.sheets[1])

    def test_clear_content(self):
        self.wb1.sheets['Sheet2'].range('G10').value = 22
        self.wb1.sheets['Sheet2'].clear_contents()
        self.assertEqual(self.wb1.sheets['Sheet2'].range('G10').value, None)

    def test_clear(self):
        self.wb1.sheets['Sheet2'].range('G10').value = 22
        self.wb1.sheets['Sheet2'].range('G10').color = (255, 255, 255)
        self.wb1.sheets['Sheet2'].clear()
        self.assertEqual(self.wb1.sheets['Sheet2'].range('G10').value, None)
        self.assertEqual(self.wb1.sheets['Sheet2'].range('G10').color, None)

    def test_autofit(self):
        sht = self.wb1.sheets['Sheet1']
        sht.range('A1:D4').value = 'test_string'
        sht.range('A1:D4').row_height = 40
        sht.range('A1:D4').column_width = 40
        self.assertEqual(sht.range('A1:D4').row_height, 40)
        self.assertEqual(sht.range('A1:D4').column_width, 40)

        sht.autofit()

        self.assertNotEqual(sht.range('A1:D4').row_height, 40)
        self.assertNotEqual(sht.range('A1:D4').column_width, 40)

        # Just checking if they don't throw an error
        sht.autofit('r')
        sht.autofit('c')
        sht.autofit('rows')
        sht.autofit('columns')

    def test_delete(self):
        self.assertTrue('Sheet1' in [i.name for i in self.wb1.sheets])
        self.wb1.sheets['Sheet1'].delete()
        self.assertFalse('Sheet1' in [i.name for i in self.wb1.sheets])


if __name__ == '__main__':
    unittest.main()
