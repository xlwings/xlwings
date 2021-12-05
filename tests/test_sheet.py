import sys
import unittest

import xlwings as xw
from .common import TestBase


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
            self.assertEqual(xw.apps.keys()[0], self.app1)
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

    def test_used_range(self):
        self.wb1.sheets[0].range('A1:C7').value = 1
        self.assertEqual(self.wb1.sheets[0].used_range, self.wb1.sheets[0].range("A1:C7"))

    def test_visible(self):
        self.assertTrue(self.wb1.sheets[0].visible)
        self.wb1.sheets[0].visible = False
        self.assertFalse(self.wb1.sheets[0].visible)

    def test_sheet_copy_without_arguments(self):
        original_name = self.wb1.sheets[0].name
        self.wb1.sheets[0]['A1'].value = 'xyz'
        self.wb1.sheets[0].copy()
        self.assertEqual(self.wb1.sheets[-1].name, original_name + ' (2)')
        self.assertEqual(self.wb1.sheets[-1]['A1'].value, 'xyz')

    def test_sheet_copy_with_before_and_after(self):
        with self.assertRaises(AssertionError):
            self.wb1.sheets[0].copy(before=self.wb1.sheets[0], after=self.wb1.sheets[0])

    def test_sheet_copy_before_same_book(self):
        original_name = self.wb1.sheets[0].name
        self.wb1.sheets[0]['A1'].value = 'xyz'
        copied_sheet = self.wb1.sheets[0].copy(before=self.wb1.sheets[0])
        self.assertNotEqual(self.wb1.sheets[0].name, original_name)
        self.assertEqual(self.wb1.sheets[0]['A1'].value, 'xyz')
        self.assertEqual(copied_sheet.name, self.wb1.sheets[0].name)

    def test_sheet_copy_after_same_book(self):
        original_name = self.wb1.sheets[0].name
        self.wb1.sheets[0]['A1'].value = 'xyz'
        self.wb1.sheets[0].copy(after=self.wb1.sheets[0])
        self.assertNotEqual(self.wb1.sheets[1].name, original_name)
        self.assertEqual(self.wb1.sheets[1]['A1'].value, 'xyz')

    def test_sheet_copy_before_same_book_new_name(self):
        self.wb1.sheets[0]['A1'].value = 'xyz'
        self.wb1.sheets[0].copy(before=self.wb1.sheets[0], name='mycopy')
        self.assertEqual(self.wb1.sheets[0].name, 'mycopy')
        self.assertEqual(self.wb1.sheets[0]['A1'].value, 'xyz')

    def test_sheet_copy_before_same_book_new_name_already_exists(self):
        self.wb1.sheets[0]['A1'].value = 'xyz'
        self.wb1.sheets[0].copy(before=self.wb1.sheets[0], name='mycopy')
        with self.assertRaises(ValueError):
            self.wb1.sheets[0].copy(before=self.wb1.sheets[0], name='mycopy')

    def test_sheet_copy_before_different_book(self):
        self.wb1.sheets[0]['A1'].value = 'xyz'
        wb2 = self.wb1.app.books.add()
        self.wb1.sheets[0].copy(before=wb2.sheets[0])
        self.assertEqual(wb2.sheets[0]['A1'].value, self.wb1.sheets[0]['A1'].value)

    def test_sheet_copy_before_different_book_same_name(self):
        mysheet = self.wb1.sheets.add('mysheet')
        mysheet['A1'].value = 'xyz'
        wb2 = self.wb1.app.books.add()
        self.wb1.sheets[0].copy(after=wb2.sheets[0], name='mysheet')
        self.assertEqual(wb2.sheets[1]['A1'].value, mysheet['A1'].value)
        self.assertEqual(wb2.sheets[1].name, 'mysheet')
        with self.assertRaises(ValueError):
            self.wb1.sheets[0].copy(after=wb2.sheets[0], name='mysheet')


class TestPageSetup(unittest.TestCase):
    def test_print_area(self):
        sheet = xw.Book().sheets[0]
        self.assertIsNone(sheet.page_setup.print_area)
        sheet.page_setup.print_area = 'A1:B2'
        self.assertEqual(sheet.page_setup.print_area, '$A$1:$B$2')
        sheet.page_setup.print_area = None
        self.assertIsNone(sheet.page_setup.print_area)
        sheet.book.close()


if __name__ == '__main__':
    unittest.main()
