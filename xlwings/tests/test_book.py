# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os
import sys
import unittest

import xlwings as xw
from xlwings.tests.common import TestBase, this_dir
from xlwings import PY3


class TestBooks(TestBase):
    def test_indexing(self):
        self.assertEqual(self.app1.books[0], self.app1.books(1))

    def test_len(self):
        count = self.app1.books.count
        wb = self.app1.books.add()
        self.assertEqual(len(self.app1.books), count + 1)

    def test_count(self):
        self.assertEqual(len(self.app1.books), self.app1.books.count)

    def test_add(self):
        current_count = self.app1.books.count
        self.app1.books.add()
        self.assertEqual(len(self.app1.books), current_count + 1)

    def test_open(self):
        fullname = os.path.join(this_dir, 'test book.xlsx')
        wb = self.app1.books.open(fullname)
        self.assertEqual(self.app1.books.active, wb)

        wb2 = self.app1.books.open(fullname)  # Should not reopen
        self.assertEqual(wb, wb2)

    def test_open_bad_name(self):
        fullname = os.path.join(this_dir, 'no book.xlsx')  
        if PY3:
            with self.assertRaises(FileNotFoundError):
                self.app1.books.open(fullname)
        else:
            with self.assertRaises(IOError):
                self.app1.books.open(fullname)
                
    def test_iter(self):
        for ix, wb in enumerate(self.app1.books):
            self.assertEqual(self.app1.books[ix], wb)


class TestBook(TestBase):
    def test_instantiate_unsaved(self):
        self.wb1.sheets[0].range('B2').value = 123
        wb2 = self.app1.books[self.wb1.name]
        self.assertEqual(wb2.sheets[0].range('B2').value, 123)

    def test_instantiate_two_unsaved(self):
        """Covers GH Issue #63"""
        wb1 = self.wb1
        wb2 = self.app1.books.add()

        wb2.sheets[0].range('A1').value = 2.
        wb1.sheets[0].range('A1').value = 1.

        self.assertEqual(wb2.sheets[0].range('A1').value, 2.)
        self.assertEqual(wb1.sheets[0].range('A1').value, 1.)

    def test_instantiate_saved_by_name(self):
        wb1 = self.app1.books.open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test book.xlsx'))
        wb1.sheets[0].range('A1').value = 'xx'
        wb2 = self.app1.books['test book.xlsx']
        self.assertEqual(wb2.sheets[0].range('A1').value, 'xx')

    def test_instantiate_saved_by_fullpath(self):
        # unicode name of book, but not unicode path
        wb = self.app1.books.add()
        if sys.platform.startswith('darwin') and self.app1.version.major >= 15:
            dst = os.path.join(os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/', 'üni cöde.xlsx')
        else:
            dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'üni cöde.xlsx')
        if os.path.isfile(dst):
            os.remove(dst)
        wb.save(dst)
        wb2 = self.app1.books.open(dst)  # Book is open
        wb2.sheets[0].range('A1').value = 1
        wb2.save()
        wb2.close()
        wb3 = self.app1.books.open(dst)  # Book is closed
        self.assertEqual(wb3.sheets[0].range('A1').value, 1.)
        wb3.close()
        os.remove(dst)

    def test_active(self):
        self.wb2.sheets[0].range('A1').value = 'active_book'  # 2nd instance
        self.assertEqual(self.app2.books.active.sheets[0].range('A1').value, 'active_book')

    def test_mock_caller(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test book.xlsx')

        wb = self.app1.books.open(path)
        wb.set_mock_caller()
        wb2 = xw.Book.caller()
        wb2.sheets[0].range('A1').value = 333
        self.assertEqual(wb2.sheets[0].range('A1').value, 333)

    def test_macro(self):
        # NOTE: Uncheck Macro security check in Excel
        _none = None if sys.platform.startswith('win') else ''

        src = os.path.abspath(os.path.join(this_dir, 'macro book.xlsm'))
        wb = self.app1.books.open(src)

        test1 = wb.macro('Module1.Test1')
        test2 = wb.macro('Module1.Test2')
        test3 = wb.macro('Module1.Test3')
        test4 = wb.macro('Test4')

        res1 = test1('Test1a', 'Test1b')

        self.assertEqual(res1, 1)
        self.assertEqual(test2(), 2)
        self.assertEqual(test3('Test3a', 'Test3b'), _none)
        self.assertEqual(test4(), _none)
        self.assertEqual(wb.sheets[0].range('A1').value, 'Test1a')
        self.assertEqual(wb.sheets[0].range('A2').value, 'Test1b')
        self.assertEqual(wb.sheets[0].range('A3').value, 'Test2')
        self.assertEqual(wb.sheets[0].range('A4').value, 'Test3a')
        self.assertEqual(wb.sheets[0].range('A5').value, 'Test3b')
        self.assertEqual(wb.sheets[0].range('A6').value, 'Test4')

    def test_name(self):
        wb = self.app1.books.open(os.path.join(this_dir, 'test book.xlsx'))
        self.assertEqual(wb.name, 'test book.xlsx')

    def test_sheets(self):
        self.assertEqual(len(self.wb1.sheets), 3)

    def test_app(self):
        # Win Excel 2016 struggles with this test in any other more meaningful way
        self.assertEqual(self.app1.books[0].app.books[0], self.app1.books[0])

    def test_close(self):
        wb = self.app1.books.add()
        count = self.app1.books.count
        wb.close()
        self.assertEqual(len(self.app1.books), count - 1)

    def test_save_naked(self):
        if sys.platform.startswith('darwin') and self.app1.version.major >= 15:
            folder = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'
            if os.path.isdir(folder):
                os.chdir(folder)

        cwd = os.getcwd()
        target_file_path = os.path.join(cwd, self.wb1.name + '.xlsx')
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

        self.wb1.save()

        self.assertTrue(os.path.isfile(target_file_path))

        self.app1.books[os.path.basename(target_file_path)].close()
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

    def test_save_path(self):
        if sys.platform.startswith('darwin') and self.app1.version.major >= 15:
            folder = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/'
            if os.path.isdir(folder):
                os.chdir(folder)

        cwd = os.getcwd()
        target_file_path = os.path.join(cwd, 'TestFile.xlsx')
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

        self.wb1.save(target_file_path)

        self.assertTrue(os.path.isfile(target_file_path))

        self.app1.books[os.path.basename(target_file_path)].close()
        if os.path.isfile(target_file_path):
            os.remove(target_file_path)

    def test_fullname(self):
        fullname = os.path.join(this_dir, 'test book.xlsx')
        wb = self.app1.books.open(fullname)
        self.assertEqual(wb.fullname.lower(), fullname.lower())

    def test_names(self):
        names = self.wb1.names
        self.assertEqual(len(names), 0)

    def test_activate(self):
        if sys.platform.startswith('win') and self.app1.version.major > 14:
            with self.assertRaises(Exception):
                wb1 = self.app1.books.add()
                wb2 = self.app2.books.add()
                wb1.activate()
        else:
            wb1 = self.app1.books.add()
            wb2 = self.app2.books.add()
            wb1.activate()
            self.assertEqual(xw.books.active, wb1)
            wb2.activate()
            self.assertEqual(xw.books.active, wb2)

    def test_selection(self):
        self.wb1.sheets[0].range('B10').select()
        self.assertEqual(self.wb1.selection.address, '$B$10')
        self.wb2.sheets[0].range('A2:C3').select()
        self.assertEqual(self.wb2.selection.address, '$A$2:$C$3')

    def test_sheet(self):
        self.wb1.sheets.add()
        self.assertEqual(len(self.wb1.sheets), 4)

if __name__ == '__main__':
    unittest.main()
