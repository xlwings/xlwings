# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import unittest

from xlwings.tests.common import TestBase


class TestNames(TestBase):
    def test_get_names_index(self):
        self.wb1.sheets[0].range('B2:D10').name = 'test1'
        self.wb1.sheets[0].range('A1').name = 'test2'
        self.assertEqual(self.wb1.names(1).name, 'test1')
        self.assertEqual(self.wb1.names[1].name, 'test2')

    def test_names_contain(self):
        self.wb1.sheets[0].range('B2:D10').name = 'test1'
        self.assertTrue('test1' in self.wb1.names)

    def test_len(self):
        self.wb1.sheets[0].range('B2:D10').name = 'test1'
        self.wb1.sheets[0].range('A1').name = 'test2'
        self.assertEqual(len(self.wb1.names), 2)

    def test_count(self):
        self.assertEqual(len(self.wb1.names), self.wb1.names.count)

    def test_names_iter(self):
        self.wb1.sheets[0].range('B2:D10').name = 'test1'
        self.wb1.sheets[0].range('A1').name = 'test2'
        for ix, n in enumerate(self.wb1.names):
            if ix == 0:
                self.assertEqual(n.name, 'test1')
            if ix == 1:
                self.assertEqual(n.name, 'test2')

    def test_get_inexisting_name(self):
        self.assertIsNone(self.wb1.sheets[0].range('A1').name)

    def test_get_set_named_range(self):
        self.wb1.sheets[0].range('A100').name = 'test1'
        self.assertEqual(self.wb1.sheets[0].range('A100').name.name, 'test1')

        self.wb1.sheets[0].range('A200:B204').name = 'test2'
        self.assertEqual(self.wb1.sheets[0].range('A200:B204').name.name, 'test2')

    def test_delete_named_item1(self):
        self.wb1.sheets[0].range('B10:C11').name = 'to_be_deleted'
        self.assertEqual(self.wb1.sheets[0].range('to_be_deleted').name.name, 'to_be_deleted')

        del self.wb1.names['to_be_deleted']
        self.assertIsNone(self.wb1.sheets[0].range('B10:C11').name)

    def test_delete_named_item2(self):
        self.wb1.sheets[0].range('B10:C11').name = 'to_be_deleted'
        self.assertEqual(self.wb1.sheets[0].range('to_be_deleted').name.name, 'to_be_deleted')

        self.wb1.names['to_be_deleted'].delete()
        self.assertIsNone(self.wb1.sheets[0].range('B10:C11').name)

    def test_delete_named_item3(self):
        self.wb1.sheets[0].range('B10:C11').name = 'to_be_deleted'
        self.assertEqual(self.wb1.sheets[0].range('to_be_deleted').name.name, 'to_be_deleted')

        self.wb1.sheets[0].range('to_be_deleted').name.delete()
        self.assertIsNone(self.wb1.sheets[0].range('B10:C11').name)

    def test_names_collection(self):
        self.wb1.sheets[0].range('A1').name = 'name1'
        self.wb1.sheets[0].range('A2').name = 'name2'
        self.assertTrue('name1' in self.wb1.names and 'name2' in self.wb1.names)

        self.wb1.sheets[0].range('A3').name = 'name3'
        self.assertTrue('name1' in self.wb1.names and 'name2' in self.wb1.names and
                    'name3' in self.wb1.names)

    def test_sheet_scope(self):
        self.wb2.sheets[0].range('B2:C3').name = 'Sheet1!sheet_scope1'
        self.wb2.sheets[0].range('sheet_scope1').value = [[1., 2.], [3., 4.]]
        self.assertEqual(self.wb2.sheets[0].range('B2:C3').value, [[1., 2.], [3., 4.]])
        with self.assertRaises(Exception):
            values = self.wb2.sheets[1].range('sheet_scope1').value

    def test_workbook_scope(self):
        self.wb1.sheets[0].range('A1').name = 'test1'
        self.wb1.sheets[0].range('test1').value = 123.
        self.assertEqual(self.wb1.names['test1'].refers_to_range.value, 123.)

    def test_contains_name(self):
        self.wb1.sheets[0].range('A1').name = 'test1'
        self.assertTrue(self.wb1.names.contains('test1'))
        self.assertFalse(self.wb1.names.contains('test2'))

    def test_names_add(self):
        self.wb1.names.add('test1', '=Sheet1!$A$1:$B$3')
        self.assertEqual(self.wb1.sheets[0].range('A1:B3').name.name, 'test1')

    def test_refers_to_range(self):
        self.wb1.sheets[0].range('B2:D10').name = 'test1'
        self.assertEqual(self.wb1.sheets[0].range('B2:D10').address, self.wb1.sheets[0].range('B2:D10').name.refers_to_range.address)

if __name__ == '__main__':
    unittest.main()
