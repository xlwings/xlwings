# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import sys
import os

from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

import xlwings as xw
from xlwings.constants import RgbColor
from .common import TestBase, this_dir
from .test_data import data

# Mac imports
if sys.platform.startswith('darwin'):
    from appscript import k as kw


class TestRangeInstantiation(TestBase):
    def test_range1(self):
        r = self.wb1.sheets[0].range('A1')
        assert_equal(r.address, '$A$1')

    def test_range2(self):
        r = self.wb1.sheets[0].range('A1:A1')
        assert_equal(r.address, '$A$1')

    def test_range3(self):
        r = self.wb1.sheets[0].range('B2:D5')
        assert_equal(r.address, '$B$2:$D$5')

    def test_range4(self):
        r = self.wb1.sheets[0].range((1, 1))
        assert_equal(r.address, '$A$1')

    def test_range5(self):
        r = self.wb1.sheets[0].range((1, 1), (1, 1))
        assert_equal(r.address, '$A$1')

    def test_range6(self):
        r = self.wb1.sheets[0].range((2, 2), (5, 4))
        assert_equal(r.address, '$B$2:$D$5')

    def test_range7(self):
        r = self.wb1.sheets[0].range('A1', (2, 2))
        assert_equal(r.address, '$A$1:$B$2')

    def test_range8(self):
        r = self.wb1.sheets[0].range((1, 1), 'B2')
        assert_equal(r.address, '$A$1:$B$2')

    def test_range9(self):
        r = self.wb1.sheets[0].range(xw.Range('A1'), xw.Range('B2'))
        assert_equal(r.address, '$A$1:$B$2')


class TestRangeAttributes(TestBase):
    def test_iterator(self):
        self.wb1.sheets[0].range('A20').value = [[1., 2.], [3., 4.]]
        r = self.wb1.sheets[0].range('A20:B21')

        assert_equal([c.value for c in r], [1., 2., 3., 4.])

        # check that reiterating on same range works properly
        assert_equal([c.value for c in r], [1., 2., 3., 4.])

    def test_sheet(self):
        assert_equal(self.wb1.sheets[1].range('A1').sheet.name, self.wb1.sheets[1].name)

    def test_len(self):
        assert_equal(len(self.wb1.sheets[0].range('A1:C4')), 12)

    def test_row(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').row, 3)

    def test_column(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').column, 2)

    def test_row_count(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').row_count, 3)

    def test_column_count(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').column_count, 5)

    def raw_value(self):
        pass  # TODO

    def test_clear_content(self):
        self.wb1.sheets[0].range('G1').value = 22
        self.wb1.sheets[0].range('G1').clear_contents()
        assert_equal(self.wb1.sheets[0].range('G1').value, None)

    def test_clear(self):
        self.wb1.sheets[0].range('G1').value = 22
        self.wb1.sheets[0].range('G1').clear()
        assert_equal(self.wb1.sheets[0].range('G1').value, None)

    def test_end(self):
        self.wb1.sheets[0].range('A1:C5').value = 1.
        assert_equal(self.wb1.sheets[0].range('A1').end('d').address, self.wb1.sheets[0].range('A5').address)
        assert_equal(self.wb1.sheets[0].range('A1').end('down').address, self.wb1.sheets[0].range('A5').address)
        assert_equal(self.wb1.sheets[0].range('C5').end('u').address, self.wb1.sheets[0].range('C1').address)
        assert_equal(self.wb1.sheets[0].range('C5').end('up').address, self.wb1.sheets[0].range('C1').address)
        assert_equal(self.wb1.sheets[0].range('A1').end('right').address, self.wb1.sheets[0].range('C1').address)
        assert_equal(self.wb1.sheets[0].range('A1').end('r').address, self.wb1.sheets[0].range('C1').address)
        assert_equal(self.wb1.sheets[0].range('C5').end('left').address, self.wb1.sheets[0].range('A5').address)
        assert_equal(self.wb1.sheets[0].range('C5').end('l').address, self.wb1.sheets[0].range('A5').address)

    def test_formula(self):
        self.wb1.sheets[0].range('A1').formula = '=SUM(A2:A10)'
        assert_equal(self.wb1.sheets[0].range('A1').formula, '=SUM(A2:A10)')

    def test_formula_array(self):
        self.wb1.sheets[0].range('A1').value = [[1, 4], [2, 5], [3, 6]]
        self.wb1.sheets[0].range('D1').formula_array = '=SUM(A1:A3*B1:B3)'
        assert_equal(self.wb1.sheets[0].range('D1').value, 32.)

    def test_column_width(self):
        self.wb1.sheets[0].range('A1:B2').column_width = 10.0
        result = self.wb1.sheets[0].range('A1').column_width
        assert_equal(10.0, result)

        self.wb1.sheets[0].range('A1:B2').value = 'ensure cells are used'
        self.wb1.sheets[0].range('B2').column_width = 20.0
        result = self.wb1.sheets[0].range('A1:B2').column_width
        if sys.platform.startswith('win'):
            assert_equal(None, result)
        else:
            assert_equal(kw.missing_value, result)

    def test_row_height(self):
        self.wb1.sheets[0].range('A1:B2').row_height = 15.0
        result = self.wb1.sheets[0].range('A1').row_height
        assert_equal(15.0, result)

        self.wb1.sheets[0].range('A1:B2').value = 'ensure cells are used'
        self.wb1.sheets[0].range('B2').row_height = 20.0
        result = self.wb1.sheets[0].range('A1:B2').row_height
        if sys.platform.startswith('win'):
            assert_equal(None, result)
        else:
            assert_equal(kw.missing_value, result)

    def test_width(self):
        """test_width: Width depends on default style text size, so do not test absolute widths"""
        self.wb1.sheets[0].range('A1:D4').column_width = 10.0
        result_before = self.wb1.sheets[0].range('A1').width
        self.wb1.sheets[0].range('A1:D4').column_width = 12.0
        result_after = self.wb1.sheets[0].range('A1').width
        assert_true(result_after > result_before)

    def test_height(self):
        self.wb1.sheets[0].range('A1:D4').row_height = 60.0
        result = self.wb1.sheets[0].range('A1:D4').height
        assert_equal(240.0, result)

    def test_left(self):
        assert_equal(self.wb1.sheets[0].range('A1').left, 0.0)
        self.wb1.sheets[0].range('A1').column_width = 20.0
        assert_equal(self.wb1.sheets[0].range('B1').left, self.wb1.sheets[0].range('A1').width)

    def test_top(self):
        assert_equal(self.wb1.sheets[0].range('A1').top, 0.0)
        self.wb1.sheets[0].range('A1').row_height = 20.0
        assert_equal(self.wb1.sheets[0].range('A2').top, self.wb1.sheets[0].range('A1').height)

    def test_number_format_cell(self):
        format_string = "mm/dd/yy;@"
        self.wb1.sheets[0].range('A1').number_format = format_string
        result = self.wb1.sheets[0].range('A1').number_format
        assert_equal(format_string, result)

    def test_number_format_range(self):
        format_string = "mm/dd/yy;@"
        self.wb1.sheets[0].range('A1:D4').number_format = format_string
        result = self.wb1.sheets[0].range('A1:D4').number_format
        assert_equal(format_string, result)

    def test_get_address(self):
        wb1 = xw.Book(os.path.join(this_dir, 'test book.xlsx'))

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address()
        assert_equal(res, '$A$1:$C$3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(False)
        assert_equal(res, '$A1:$C3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(True, False)
        assert_equal(res, 'A$1:C$3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(False, False)
        assert_equal(res, 'A1:C3')

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(include_sheetname=True)
        assert_equal(res, "'Sheet1'!$A$1:$C$3")

        res = wb1.sheets[1].range((1, 1), (3, 3)).get_address(include_sheetname=True)
        assert_equal(res, "'Sheet2'!$A$1:$C$3")

        res = wb1.sheets[0].range((1, 1), (3, 3)).get_address(external=True)
        assert_equal(res, "'[test book.xlsx]Sheet1'!$A$1:$C$3")

        wb1.close()

    def test_address(self):
        assert_equal(self.wb1.sheets[0].range('A1:B2').address, '$A$1:$B$2')

    def test_current_region(self):
        values = [[1., 2.], [3., 4.]]
        self.wb1.sheets[0].range('A20').value = values
        assert_equal(self.wb1.sheets[0].range('B21').current_region.value, values)

    def test_autofit_range(self):
        self.wb1.sheets[0].range('A1:D4').value = 'test_string'

        self.wb1.sheets[0].range('A1:D4').row_height = 40
        self.wb1.sheets[0].range('A1:D4').column_width = 40
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').row_height)
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').column_width)
        self.wb1.sheets[0].range('A1:D4').autofit()
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').column_width)
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').row_height)

        self.wb1.sheets[0].range('A1:D4').row_height = 40
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').row_height)
        self.wb1.sheets[0].range('A1:D4').autofit('r')
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').row_height)

        self.wb1.sheets[0].range('A1:D4').column_width = 40
        assert_equal(40, self.wb1.sheets[0].range('A1:D4').column_width)
        self.wb1.sheets[0].range('A1:D4').autofit('c')
        assert_true(40 != self.wb1.sheets[0].range('A1:D4').column_width)

        self.wb1.sheets[0].range('A1:D4').autofit('rows')
        self.wb1.sheets[0].range('A1:D4').autofit('columns')

    def test_autofit_col(self):
        self.wb1.sheets[0].range('A1:D4').value = 'test_string'
        self.wb1.sheets[0].range('A:D').column_width = 40
        assert_equal(40, self.wb1.sheets[0].range('A:D').column_width)
        self.wb1.sheets[0].range('A:D').autofit()
        assert_true(40 != self.wb1.sheets[0].range('A:D').column_width)

        # Just checking if they don't throw an error
        self.wb1.sheets[0].range('A:D').autofit('r')
        self.wb1.sheets[0].range('A:D').autofit('c')
        self.wb1.sheets[0].range('A:D').autofit('rows')
        self.wb1.sheets[0].range('A:D').autofit('columns')

    def test_autofit_row(self):
        self.wb1.sheets[0].range('A1:D4').value = 'test_string'
        self.wb1.sheets[0].range('1:10').row_height = 40
        assert_equal(40, self.wb1.sheets[0].range('1:10').row_height)
        self.wb1.sheets[0].range('1:10').autofit()
        assert_true(40 != self.wb1.sheets[0].range('1:10').row_height)

        # Just checking if they don't throw an error
        self.wb1.sheets[0].range('1:1000000').autofit('r')
        self.wb1.sheets[0].range('1:1000000').autofit('c')
        self.wb1.sheets[0].range('1:1000000').autofit('rows')
        self.wb1.sheets[0].range('1:1000000').autofit('columns')

    def test_color(self):
        rgb = (30, 100, 200)

        self.wb1.sheets[0].range('A1').color = rgb
        assert_equal(rgb, self.wb1.sheets[0].range('A1').color)

        self.wb1.sheets[0].range('A2').color = RgbColor.rgbAqua
        assert_equal((0, 255, 255), self.wb1.sheets[0].range('A2').color)

        self.wb1.sheets[0].range('A2').color = None
        assert_equal(self.wb1.sheets[0].range('A2').color, None)

        self.wb1.sheets[0].range('A1:D4').color = rgb
        assert_equal(rgb, self.wb1.sheets[0].range('A1:D4').color)

    def test_len_rows(self):
        assert_equal(len(self.wb1.sheets[0].range('A1:C4').rows), 4)

    def test_len_cols(self):
        assert_equal(len(self.wb1.sheets[0].range('A1:C4').columns), 3)

    def test_shape(self):
        assert_equal(self.wb1.sheets[0].range('A1:C4').shape, (4, 3))

    def test_size(self):
        assert_equal(self.wb1.sheets[0].range('A1:C4').size, 12)

    def test_table(self):
        self.wb1.sheets[0].range('A1').value = data
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range('A3:B3').number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = self.wb1.sheets[0].range('A1').table.value
        assert_equal(cells, data)

    def test_vertical(self):
        self.wb1.sheets[0].range('A10').value = data
        if sys.platform.startswith('win') and self.wb1.app.version == '14.0':
            self.wb1.sheets[0].range('A12:B12').number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = self.wb1.sheets[0].range('A10').vertical.value
        assert_equal(cells, [row[0] for row in data])

    def test_horizontal(self):
        self.wb1.sheets[0].range('A20').value = data
        cells = self.wb1.sheets[0].range('A20').horizontal.value
        assert_equal(cells, data[0])

    def test_hyperlink(self):
        address = 'www.xlwings.org'
        # Naked address
        self.wb1.sheets[0].range('A1').add_hyperlink(address)
        assert_equal(self.wb1.sheets[0].range('A1').value, address)
        hyperlink = self.wb1.sheets[0].range('A1').hyperlink
        if not hyperlink.endswith('/'):
            hyperlink += '/'
        assert_equal(hyperlink, 'http://' + address + '/')

        # Address + FriendlyName
        self.wb1.sheets[0].range('A2').add_hyperlink(address, 'test_link')
        assert_equal(self.wb1.sheets[0].range('A2').value, 'test_link')
        hyperlink = self.wb1.sheets[0].range('A2').hyperlink
        if not hyperlink.endswith('/'):
            hyperlink += '/'
        assert_equal(hyperlink, 'http://' + address + '/')

    def test_hyperlink_formula(self):
        self.wb1.sheets[0].range('B10').formula = '=HYPERLINK("http://xlwings.org", "xlwings")'
        assert_equal(self.wb1.sheets[0].range('B10').hyperlink, 'http://xlwings.org')

    def test_resize(self):
        r = self.wb1.sheets[0].range('A1').resize(4, 5)
        assert_equal(r.address, '$A$1:$E$4')

        r = self.wb1.sheets[0].range('A1').resize(row_size=4)
        assert_equal(r.address, '$A$1:$A$4')

        r = self.wb1.sheets[0].range('A1:B4').resize(column_size=5)
        assert_equal(r.address, '$A$1:$E$4')

        r = self.wb1.sheets[0].range('A1:B4').resize(row_size=5)
        assert_equal(r.address, '$A$1:$B$5')

        r = self.wb1.sheets[0].range('A1:B4').resize()
        assert_equal(r.address, '$A$1:$B$4')

        r = self.wb1.sheets[0].range('A1:C5').resize(row_size=1)
        assert_equal(r.address, '$A$1:$C$1')

        assert_raises(AssertionError, self.wb1.sheets[0].range('A1:B4').resize, row_size=0)
        assert_raises(AssertionError, self.wb1.sheets[0].range('A1:B4').resize, column_size=0)

    def test_offset(self):
        o = self.wb1.sheets[0].range('A1:B3').offset(3, 4)
        assert_equal(o.address, '$E$4:$F$6')

        o = self.wb1.sheets[0].range('A1:B3').offset(row_offset=3)
        assert_equal(o.address, '$A$4:$B$6')

        o = self.wb1.sheets[0].range('A1:B3').offset(column_offset=4)
        assert_equal(o.address, '$E$1:$F$3')

    def test_last_cell(self):
        assert_equal(self.wb1.sheets[0].range('B3:F5').last_cell.row, 5)
        assert_equal(self.wb1.sheets[0].range('B3:F5').last_cell.column, 6)

    def test_select(self):
        self.wb1.sheets[0].range('C10').select()
        assert_equal(self.app1.selection.address, self.wb1.sheets[0].range('C10').address)


class TestRangeIndexing(TestBase):
    @raises(IndexError)
    def test_zero_based_index1(self):
        self.wb1.sheets[0].range((0, 1)).value = 123

    @raises(IndexError)
    def test_zero_based_index2(self):
        a = self.wb1.sheets[0].range((1, 1), (1, 0)).value

    @raises(IndexError)
    def test_zero_based_index3(self):
        xw.Range((1, 0)).value = 123

    @raises(IndexError)
    def test_zero_based_index4(self):
        a = xw.Range((1, 0), (1, 0)).value


class TestRangeSlicing(TestBase):
    pass


