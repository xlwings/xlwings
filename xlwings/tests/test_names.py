# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

import xlwings as xw
from .common import TestBase


class TestNames(TestBase):
    def test_delete_named_item1(self):
        self.wb[0].range('B10:C11').name = 'to_be_deleted'
        assert_equal(self.wb[0].range('to_be_deleted').name.name, 'to_be_deleted')

        del self.wb.names['to_be_deleted']

        assert_not_equal(self.wb[0].range('B10:C11').name, 'to_be_deleted')

    def test_delete_named_item2(self):
        self.wb[0].range('B10:C11').name = 'to_be_deleted'
        assert_equal(self.wb[0].range('to_be_deleted').name.name, 'to_be_deleted')

        self.wb.names['to_be_deleted'].delete()

        assert_not_equal(self.wb[0].range('B10:C11').name, 'to_be_deleted')

    def test_delete_named_item3(self):
        self.wb[0].range('B10:C11').name = 'to_be_deleted'
        assert_equal(self.wb[0].range('to_be_deleted').name.name, 'to_be_deleted')

        self.wb[0].range('to_be_deleted').name.delete()

        assert_not_equal(self.wb[0].range('B10:C11').name, 'to_be_deleted')

    def test_names_collection(self):
        self.wb[0].range('A1').name = 'name1'
        self.wb[0].range('A2').name = 'name2'
        assert_true('name1' in self.wb.names and 'name2' in self.wb.names)

        self.wb[0].range('A3').name = 'name3'
        assert_true('name1' in self.wb.names and 'name2' in self.wb.names and
                    'name3' in self.wb.names)

    def test_sheet_scope(self):
        self.wb[0].range('B2:C3').name = 'Sheet1!sheet_scope1'
        self.wb[0].range('sheet_scope1').value = [[1., 2.], [3., 4.]]
        assert_equal(self.wb[0].range('B2:C3').value, [[1., 2.], [3., 4.]])
        self.wb[1].activate()
        assert_equal(self.wb[0].range('sheet_scope1').value, [[1., 2.], [3., 4.]])

