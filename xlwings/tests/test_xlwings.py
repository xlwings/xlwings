# -*- coding: utf-8 -*-
# TODO: clean up used workbooks

from __future__ import unicode_literals
import os
import sys
import shutil
from datetime import datetime, date

import pytz
import inspect
import nose
from nose.tools import assert_equal, raises, assert_raises, assert_true, assert_false, assert_not_equal

from xlwings import Application, Workbook, Sheet, Range, Chart, Picture, Plot, ShapeAlreadyExists
from xlwings.constants import ChartType, RgbColor
from .test_data import data, test_date_1, test_date_2, list_row_1d, list_row_2d, list_col, chart_data
from .common import TestBase, _skip_if_no_matplotlib, _skip_if_no_numpy, _skip_if_no_pandas


this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))

# Mac imports
if sys.platform.startswith('darwin'):
    from appscript import k as kw
    # TODO: uncomment the desired Excel installation or set to None for default installation
    # APP_TARGET = None
    APP_TARGET = '/Applications/Microsoft Office 2011/Microsoft Excel'
else:
    APP_TARGET = None

# Optional dependencies
try:
    import numpy as np
    from numpy.testing import assert_array_equal
    from .test_data import array_1d, array_2d
except ImportError:
    np = None
try:
    import pandas as pd
    from pandas import DataFrame, Series
    from pandas.util.testing import assert_frame_equal, assert_series_equal
    from .test_data import series_1, timeseries_1, df_1, df_2, df_dateindex, df_multiheader, df_multiindex
except ImportError:
    pd = None
try:
    import matplotlib
    from matplotlib.figure import Figure
except ImportError:
    matplotlib = None
try:
    import PIL
except ImportError:
    PIL = None


class TestSheet(TestBase):
    def setUp(self):
        super(TestSheet, self).setUp('test_workbook_1.xlsx')

    def test_activate(self):
        Sheet('Sheet2').activate()
        assert_equal(Sheet.active.name, 'Sheet2')
        Sheet(3).activate()
        assert_equal(Sheet.active.index, 3)

    def test_name(self):
        Sheet(1).name = 'NewName'
        assert_equal(Sheet(1).name, 'NewName')

    def test_index(self):
        assert_equal(Sheet('Sheet1').index, 1)

    def test_clear_content_active_sheet(self):
        Range('G10').value = 22
        Sheet.active().clear_contents()
        cell = Range('G10').value
        assert_equal(cell, None)

    def test_clear_active_sheet(self):
        Range('G10').value = 22
        Sheet.active().clear()
        cell = Range('G10').value
        assert_equal(cell, None)

    def test_clear_content(self):
        Range('Sheet2', 'G10').value = 22
        Sheet('Sheet2').clear_contents()
        cell = Range('Sheet2', 'G10').value
        assert_equal(cell, None)

    def test_clear(self):
        Range('Sheet2', 'G10').value = 22
        Sheet('Sheet2').clear()
        cell = Range('Sheet2', 'G10').value
        assert_equal(cell, None)

    def test_autofit(self):
        Range('Sheet1', 'A1:D4').value = 'test_string'
        Sheet('Sheet1').autofit()
        Sheet('Sheet1').autofit('r')
        Sheet('Sheet1').autofit('c')
        Sheet('Sheet1').autofit('rows')
        Sheet('Sheet1').autofit('columns')

    def test_add_before(self):
        new_sheet = Sheet.add(before='Sheet1')
        assert_equal(Sheet(1).name, new_sheet.name)

    def test_add_after(self):
        Sheet.add(after=Sheet.count())
        assert_equal(Sheet(Sheet.count()).name, Sheet.active.name)

        Sheet.add(after=1)
        assert_equal(Sheet(2).name, Sheet.active.name)

    def test_add_default(self):
        # TODO: test call without args properly
        Sheet.add()

    def test_add_wkb(self):
        # test use of add with wkb argument

        # Connect to an alternative test file and make Sheet1 the active sheet
        xl_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_range_1.xlsx')
        wb_2nd = Workbook(xl_file, app_visible=False, app_target=APP_TARGET)

        n_before = [sh.name for sh in Sheet.all(wkb=wb_2nd)]
        Sheet.add(name="default", wkb=wb_2nd)
        Sheet.add(name="after1", after=1, wkb=wb_2nd)
        Sheet.add(name="before1", before=1, wkb=wb_2nd)
        n_after = [sh.name for sh in Sheet.all(wkb=wb_2nd)]
        
        n_before.append("default")
        n_before.insert(1, "after1")
        n_before.insert(0, "before1")
        
        assert_equal(n_before, n_after)
        wb_2nd.close()

    def test_add_named(self):
        Sheet.add('test', before=1)
        assert_equal(Sheet(1).name, 'test')

    @raises(Exception)
    def test_add_name_already_taken(self):
        Sheet.add('Sheet1')

    def test_count(self):
        count = Sheet.count()
        assert_equal(count, 3)

    def test_all(self):
        all_names = [i.name for i in Sheet.all()]
        assert_equal(all_names, ['Sheet1', 'Sheet2', 'Sheet3'])

    def test_delete(self):
        assert_true('Sheet1' in [i.name for i in Sheet.all()])
        Sheet('Sheet1').delete()
        assert_false('Sheet1' in [i.name for i in Sheet.all()])


class TestRange(TestBase):
    def setUp(self):
        super(TestRange, self).setUp('test_range_1.xlsx')

    def test_cell(self):
        params = [('A1', 22),
                  ((1, 1), 22),
                  ('A1', 22.2222),
                  ((1, 1), 22.2222),
                  ('A1', 'Test String'),
                  ((1, 1), 'Test String'),
                  ('A1', 'éöà'),
                  ((1, 1), 'éöà'),
                  ('A2', test_date_1),
                  ((2, 1), test_date_1),
                  ('A3', test_date_2),
                  ((3, 1), test_date_2)]
        for param in params:
            yield self.check_cell, param[0], param[1]

    def check_cell(self, address, value):
        # Active Sheet
        Range(address).value = value
        cell = Range(address).value
        assert_equal(cell, value)

        # SheetName
        Range('Sheet2', address).value = value
        cell = Range('Sheet2', address).value
        assert_equal(cell, value)

        # SheetIndex
        Range(3, address).value = value
        cell = Range(3, address).value
        assert_equal(cell, value)

    def test_range_address(self):
        """ Style: Range('A1:C3') """
        address = 'C1:E3'

        # Active Sheet
        Range(address[:2]).value = data  # assign to starting cell only
        cells = Range(address).value
        assert_equal(cells, data)

        # Sheetname
        Range('Sheet2', address).value = data
        cells = Range('Sheet2', address).value
        assert_equal(cells, data)

        # Sheetindex
        Range(3, address).value = data
        cells = Range(3, address).value
        assert_equal(cells, data)

    def test_range_index(self):
        """ Style: Range((1,1), (3,3)) """
        index1 = (1, 3)
        index2 = (3, 5)

        # Active Sheet
        Range(index1, index2).value = data
        cells = Range(index1, index2).value
        assert_equal(cells, data)

        # Sheetname
        Range('Sheet2', index1, index2).value = data
        cells = Range('Sheet2', index1, index2).value
        assert_equal(cells, data)

        # Sheetindex
        Range(3, index1, index2).value = data
        cells = Range(3, index1, index2).value
        assert_equal(cells, data)

    def test_named_range_value(self):
        value = 22.222
        # Active Sheet
        Range('F1').name = 'cell_sheet1'
        Range('cell_sheet1').value = value
        cells = Range('cell_sheet1').value
        assert_equal(cells, value)

        Range('A1:C3').name = 'range_sheet1'
        Range('range_sheet1').value = data
        cells = Range('range_sheet1').value
        assert_equal(cells, data)

        # Sheetname
        Range('Sheet2', 'F1').name = 'cell_sheet2'
        Range('Sheet2', 'cell_sheet2').value = value
        cells = Range('Sheet2', 'cell_sheet2').value
        assert_equal(cells, value)

        Range('Sheet2', 'A1:C3').name = 'range_sheet2'
        Range('Sheet2', 'range_sheet2').value = data
        cells = Range('Sheet2', 'range_sheet2').value
        assert_equal(cells, data)

        # Sheetindex
        Range(3, 'F3').name = 'cell_sheet3'
        Range(3, 'cell_sheet3').value = value
        cells = Range(3, 'cell_sheet3').value
        assert_equal(cells, value)

        Range(3, 'A1:C3').name = 'range_sheet3'
        Range(3, 'range_sheet3').value = data
        cells = Range(3, 'range_sheet3').value
        assert_equal(cells, data)

    def test_array(self):
        _skip_if_no_numpy()

        # 1d array
        Range('Sheet6', 'A1').value = array_1d
        cells = Range('Sheet6', 'A1:D1').options(np.array).value
        assert_array_equal(cells, array_1d)

        # 2d array
        Range('Sheet6', 'A4').value = array_2d
        cells = Range('Sheet6', 'A4').options(np.array, expand='table').value
        assert_array_equal(cells, array_2d)

        # 1d array (atleast_2d)
        Range('Sheet6', 'A10').value = array_1d
        cells = Range('Sheet6', 'A10:D10').options(np.array, ndim=2).value
        assert_array_equal(cells, np.atleast_2d(array_1d))

        # 2d array (atleast_2d)
        Range('Sheet6', 'A12').value = array_2d
        cells = Range('Sheet6', 'A12').options(np.array, ndim=2, expand='table').value
        assert_array_equal(cells, array_2d)

    def sheet_ref(self):
        Range(Sheet(1), 'A20').value = 123
        assert_equal(Range(1, 'A20').value, 123)

        Range(Sheet(1), (2, 2), (4, 4)).value = 321
        assert_equal(Range(1, (2, 2)).value, 321)

    def test_vertical(self):
        Range('Sheet4', 'A10').value = data
        if sys.platform.startswith('win') and self.wb.application.version == '14.0':
            Range('Sheet4', 'A12:B12').xl_range.NumberFormat = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = Range('Sheet4', 'A10').vertical.value
        assert_equal(cells, [row[0] for row in data])

    def test_horizontal(self):
        Range('Sheet4', 'A20').value = data
        cells = Range('Sheet4', 'A20').horizontal.value
        assert_equal(cells, data[0])

    def test_table(self):
        Range('Sheet4', 'A1').value = data
        if sys.platform.startswith('win') and self.wb.application.version == '14.0':
            Range('Sheet4', 'A3:B3').xl_range.NumberFormat = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = Range('Sheet4', 'A1').table.value
        assert_equal(cells, data)

    def test_list(self):

        # 1d List Row
        Range('Sheet4', 'A27').value = list_row_1d
        cells = Range('Sheet4', 'A27:C27').value
        assert_equal(list_row_1d, cells)

        # 2d List Row
        Range('Sheet4', 'A29').value = list_row_2d
        cells = Range('Sheet4', 'A29:C29', ndim=2).value
        assert_equal(list_row_2d, cells)

        # 1d List Col
        Range('Sheet4', 'A31').value = list_col
        cells = Range('Sheet4', 'A31:A33').value
        assert_equal([i[0] for i in list_col], cells)
        # 2d List Col
        cells = Range('Sheet4', 'A31:A33', ndim=2).value
        assert_equal(list_col, cells)

    def test_is_cell(self):
        assert_equal(Range('A1').is_cell(), True)
        assert_equal(Range('A1:B1').is_cell(), False)
        assert_equal(Range('A1:A2').is_cell(), False)
        assert_equal(Range('A1:B2').is_cell(), False)

    def test_is_row(self):
        assert_equal(Range('A1').is_row(), False)
        assert_equal(Range('A1:B1').is_row(), True)
        assert_equal(Range('A1:A2').is_row(), False)
        assert_equal(Range('A1:B2').is_row(), False)

    def test_is_column(self):
        assert_equal(Range('A1').is_column(), False)
        assert_equal(Range('A1:B1').is_column(), False)
        assert_equal(Range('A1:A2').is_column(), True)
        assert_equal(Range('A1:B2').is_column(), False)

    def test_is_table(self):
        assert_equal(Range('A1').is_table(), False)
        assert_equal(Range('A1:B1').is_table(), False)
        assert_equal(Range('A1:A2').is_table(), False)
        assert_equal(Range('A1:B2').is_table(), True)

    def test_formula(self):
        Range('A1').formula = '=SUM(A2:A10)'
        assert_equal(Range('A1').formula, '=SUM(A2:A10)')

    def test_formula_array(self):
        Range('A1').value = [[1, 4], [2, 5], [3, 6]]
        Range('D1').formula_array = '=SUM(A1:A3*B1:B3)'
        assert_equal(Range('D1').value, 32.)

    def test_current_region(self):
        values = [[1., 2.], [3., 4.]]
        Range('A20').value = values
        assert_equal(Range('B21').current_region.value, values)

    def test_clear_content(self):
        Range('Sheet4', 'G1').value = 22
        Range('Sheet4', 'G1').clear_contents()
        cell = Range('Sheet4', 'G1').value
        assert_equal(cell, None)

    def test_clear(self):
        Range('Sheet4', 'G1').value = 22
        Range('Sheet4', 'G1').clear()
        cell = Range('Sheet4', 'G1').value
        assert_equal(cell, None)

    def test_dataframe_1(self):
        _skip_if_no_pandas()

        df_expected = df_1
        Range('Sheet5', 'A1').value = df_expected
        df_result = Range('Sheet5', 'A1:C5').options(pd.DataFrame).value
        df_result.index = pd.Int64Index(df_result.index)
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_2(self):
        """ Covers GH Issue #31"""
        _skip_if_no_pandas()

        df_expected = df_2
        Range('Sheet5', 'A9').value = df_expected
        cells = Range('Sheet5', 'B9:B15').value
        df_result = DataFrame(cells[1:], columns=[cells[0]])
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiindex(self):
        _skip_if_no_pandas()

        df_expected = df_multiindex
        Range('Sheet5', 'A20').value = df_expected
        cells = Range('Sheet5', 'D20').table.value
        multiindex = Range('Sheet5', 'A20:C28').value
        ix = pd.MultiIndex.from_tuples(multiindex[1:], names=multiindex[0])
        df_result = DataFrame(cells[1:], columns=cells[0], index=ix)
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_multiheader(self):
        _skip_if_no_pandas()

        df_expected = df_multiheader
        Range('Sheet5', 'A52').value = df_expected
        cells = Range('Sheet5', 'B52').table.value
        df_result = DataFrame(cells[2:], columns=pd.MultiIndex.from_arrays(cells[:2]))
        assert_frame_equal(df_expected, df_result)

    def test_dataframe_dateindex(self):
        _skip_if_no_pandas()

        df_expected = df_dateindex
        Range('Sheet5', 'A100').value = df_expected
        if sys.platform.startswith('win') and self.wb.application.version == '14.0':
            Range('Sheet5', 'A100').vertical.number_format = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        cells = Range('Sheet5', 'B100').table.value
        index = Range('Sheet5', 'A101').vertical.value
        df_result = DataFrame(cells[1:], index=index, columns=cells[0])
        assert_frame_equal(df_expected, df_result)

    def test_read_df_0header_0index(self):
        _skip_if_no_pandas()

        Range('A1').value = [[1, 2, 3],
                             [4, 5, 6],
                             [7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]])

        df2 = Range('A1:C3').options(pd.DataFrame, header=0, index=False).value
        assert_frame_equal(df1, df2)

    def test_df_1header_0index(self):
        _skip_if_no_pandas()

        Range('A1').options(pd.DataFrame, index=False, header=True).value = pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b'])
        df = Range('A1').options(pd.DataFrame, index=False, header=True,
                                 expand='table').value
        assert_frame_equal(df, pd.DataFrame([[1., 2.], [3., 4.]], columns=['a', 'b']))

    def test_df_0header_1index(self):
        _skip_if_no_pandas()

        Range('A1').options(pd.DataFrame, index=True, header=False).value = pd.DataFrame([[1., 2.], [3., 4.]], index=[10., 20.])
        df = Range('A1').options(pd.DataFrame, index=True, header=False,
                                 expand='table').value
        assert_frame_equal(df, pd.DataFrame([[1., 2.], [3., 4.]], index=[10., 20.]))

    def test_read_df_1header_1namedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [['ix1', 'c', 'd', 'c'],
                             [1, 1, 2, 3],
                             [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=[1., 2.],
                           columns=['c', 'd', 'c'])
        df1.index.name = 'ix1'

        df2 = Range('A1:D3').options(pd.DataFrame).value
        assert_frame_equal(df1, df2)

    def test_read_df_1header_1unnamedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [[None, 'c', 'd', 'c'],
                             [1, 1, 2, 3],
                             [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=pd.Index([1., 2.]),
                           columns=['c', 'd', 'c'])

        df2 = Range('A1:D3').options(pd.DataFrame).value

        assert_frame_equal(df1, df2)

    def test_read_df_2header_1namedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [[None, 'a', 'a', 'b'],
                             ['ix1', 'c', 'd', 'c'],
                             [1, 1, 2, 3],
                             [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=[1., 2.],
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))
        df1.index.name = 'ix1'

        df2 = Range('A1:D4').options(pd.DataFrame, header=2).value

        assert_frame_equal(df1, df2)

    def test_read_df_2header_1unnamedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [[None, 'a', 'a', 'b'],
                             [None, 'c', 'd', 'c'],
                             [1, 1, 2, 3],
                             [2, 4, 5, 6]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.]],
                           index=pd.Index([1, 2]),
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))

        df2 = Range('A1:D4').options(pd.DataFrame, header=2).value
        df2.index = pd.Int64Index(df2.index)

        assert_frame_equal(df1, df2)

    def test_read_df_2header_2namedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [[None, None, 'a', 'a', 'b'],
                             ['x1', 'x2', 'c', 'd', 'c'],
                             ['a', 1, 1, 2, 3],
                             ['a', 2, 4, 5, 6],
                             ['b', 1, 7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))

        df2 = Range('A1:E5').options(pd.DataFrame, header=2, index=2).value
        assert_frame_equal(df1, df2)

    def test_read_df_2header_2unnamedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [[None, None, 'a', 'a', 'b'],
                             [None, None, 'c', 'd', 'c'],
                             ['a', 1, 1, 2, 3],
                             ['a', 2, 4, 5, 6],
                             ['b', 1, 7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]]),
                           columns=pd.MultiIndex.from_arrays([['a', 'a', 'b'], ['c', 'd', 'c']]))

        df2 = Range('A1:E5').options(pd.DataFrame, header=2, index=2).value
        assert_frame_equal(df1, df2)

    def test_read_df_1header_2namedindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [['x1', 'x2', 'a', 'd', 'c'],
                             ['a', 1, 1, 2, 3],
                             ['a', 2, 4, 5, 6],
                             ['b', 1, 7, 8, 9]]

        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           index=pd.MultiIndex.from_arrays([['a', 'a', 'b'], [1., 2., 1.]], names=['x1', 'x2']),
                           columns=['a', 'd', 'c'])

        df2 = Range('A1:E4').options(pd.DataFrame, header=1, index=2).value
        assert_frame_equal(df1, df2)

    def test_timeseries_1(self):
        _skip_if_no_pandas()

        series_expected = timeseries_1
        Range('Sheet5', 'A40').options(header=False).value = series_expected
        if sys.platform.startswith('win') and self.wb.application.version == '14.0':
            Range('Sheet5', 'A40').vertical.xl_range.NumberFormat = 'dd/mm/yyyy'  # Hack for Excel 2010 bug, see GH #43
        series_result = Range('Sheet5', 'A40:B49').options(pd.Series, header=False).value
        assert_series_equal(series_expected, series_result)

    def test_read_series_noheader_noindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [[1.],
                             [2.],
                             [3.]]
        s = Range('A1:A3').options(pd.Series, index=False, header=False).value
        assert_series_equal(s, pd.Series([1., 2., 3.]))

    def test_read_series_noheader_index(self):
        _skip_if_no_pandas()

        Range('A1').value = [[10., 1.],
                             [20., 2.],
                             [30., 3.]]
        s = Range('A1:B3').options(pd.Series, index=True, header=False).value
        assert_series_equal(s, pd.Series([1., 2., 3.], index=[10., 20., 30.]))

    def test_read_series_header_noindex(self):
        _skip_if_no_pandas()

        Range('A1').value = [['name'],
                             [1.],
                             [2.],
                             [3.]]
        s = Range('A1:A4').options(pd.Series, index=False, header=True).value
        assert_series_equal(s, pd.Series([1., 2., 3.], name='name'))

    def test_read_series_header_index(self):
        _skip_if_no_pandas()

        # Named index
        Range('A1').value = [['ix', 'name'],
                             [10., 1.],
                             [20., 2.],
                             [30., 3.]]
        s = Range('A1:B4').options(pd.Series, index=True, header=True).value
        assert_series_equal(s, pd.Series([1., 2., 3.], name='name', index=pd.Index([10., 20., 30.], name='ix')))

        # Nameless index
        Range('A1').value = [[None, 'name'],
                             [10., 1.],
                             [20., 2.],
                             [30., 3.]]
        s = Range('A1:B4').options(pd.Series, index=True, header=True).value
        assert_series_equal(s, pd.Series([1., 2., 3.], name='name', index=[10., 20., 30.]))

    def test_write_series_noheader_noindex(self):
        _skip_if_no_pandas()

        Range('A1').options(index=False).value = pd.Series([1., 2., 3.])
        assert_equal([[1.],[2.],[3.]], Range('A1').options(ndim=2, expand='table').value)

    def test_write_series_noheader_index(self):
        _skip_if_no_pandas()

        Range('A1').options(index=True).value = pd.Series([1., 2., 3.], index=[10., 20., 30.])
        assert_equal([[10., 1.],[20., 2.],[30., 3.]], Range('A1').options(ndim=2, expand='table').value)

    def test_write_series_header_noindex(self):
        _skip_if_no_pandas()

        Range('A1').options(index=False).value = pd.Series([1., 2., 3.], name='name')
        assert_equal([['name'],[1.],[2.],[3.]], Range('A1').options(ndim=2, expand='table').value)

    def test_write_series_header_index(self):
        _skip_if_no_pandas()

        # Named index
        Range('A1').value = pd.Series([1., 2., 3.], name='name', index=pd.Index([10., 20., 30.], name='ix'))
        assert_equal([['ix', 'name'],[10., 1.],[20., 2.],[30., 3.]], Range('A1').options(ndim=2, expand='table').value)

        # Nameless index
        Range('A1').value = pd.Series([1., 2., 3.], name='name', index=[10., 20., 30.])
        assert_equal([[None, 'name'],[10., 1.],[20., 2.],[30., 3.]], Range('A1:B4').options(ndim=2).value)

    def test_none(self):
        """ Covers GH Issue #16"""
        # None
        Range('Sheet1', 'A7').value = None
        assert_equal(None, Range('Sheet1', 'A7').value)
        # List
        Range('Sheet1', 'A7').value = [None, None]
        assert_equal(None, Range('Sheet1', 'A7').horizontal.value)

    def test_scalar_nan(self):
        """Covers GH Issue #15"""
        _skip_if_no_numpy()

        Range('Sheet1', 'A20').value = np.nan
        assert_equal(None, Range('Sheet1', 'A20').value)

    def test_atleast_2d_scalar(self):
        """Covers GH Issue #53a"""
        Range('Sheet1', 'A50').value = 23
        result = Range('Sheet1', 'A50').options(ndim=2).value
        assert_equal([[23]], result)

    def test_atleast_2d_scalar_as_array(self):
        """Covers GH Issue #53b"""
        _skip_if_no_numpy()

        Range('Sheet1', 'A50').value = 23
        result = Range('Sheet1', 'A50').options(np.array, ndim=2).value
        assert_equal(np.array([[23]]), result)

    def test_column_width(self):
        Range('Sheet1', 'A1:B2').column_width = 10.0
        result = Range('Sheet1', 'A1').column_width
        assert_equal(10.0, result)

        Range('Sheet1', 'A1:B2').value = 'ensure cells are used'
        Range('Sheet1', 'B2').column_width = 20.0
        result = Range('Sheet1', 'A1:B2').column_width
        if sys.platform.startswith('win'):
            assert_equal(None, result)
        else:
            assert_equal(kw.missing_value, result)

    def test_row_height(self):
        Range('Sheet1', 'A1:B2').row_height = 15.0
        result = Range('Sheet1', 'A1').row_height
        assert_equal(15.0, result)

        Range('Sheet1', 'A1:B2').value = 'ensure cells are used'
        Range('Sheet1', 'B2').row_height = 20.0
        result = Range('Sheet1', 'A1:B2').row_height
        if sys.platform.startswith('win'):
            assert_equal(None, result)
        else:
            assert_equal(kw.missing_value, result)

    def test_width(self):
        """Width depends on default style text size, so do not test absolute widths"""
        Range('Sheet1', 'A1:D4').column_width = 10.0
        result_before = Range('Sheet1', 'A1').width
        Range('Sheet1', 'A1:D4').column_width = 12.0
        result_after = Range('Sheet1', 'A1').width
        assert_true(result_after > result_before)

    def test_height(self):
        Range('Sheet1', 'A1:D4').row_height = 60.0
        result = Range('Sheet1', 'A1:D4').height
        assert_equal(240.0, result)

    def test_left(self):
        assert_equal(Range('Sheet1','A1').left, 0.0)
        Range('Sheet1','A1').column_width = 20.0
        assert_equal(Range('Sheet1','B1').left, Range('Sheet1','A1').width)

    def test_top(self):
        assert_equal(Range('Sheet1','A1').top, 0.0)
        Range('Sheet1','A1').row_height = 20.0
        assert_equal(Range('Sheet1','A2').top, Range('Sheet1','A1').height)

    def test_autofit_range(self):
        # TODO: compare col/row widths before/after - not implemented yet
        Range('Sheet1', 'A1:D4').value = 'test_string'
        Range('Sheet1', 'A1:D4').autofit()
        Range('Sheet1', 'A1:D4').autofit('r')
        Range('Sheet1', 'A1:D4').autofit('c')
        Range('Sheet1', 'A1:D4').autofit('rows')
        Range('Sheet1', 'A1:D4').autofit('columns')

    def test_autofit_col(self):
        # TODO: compare col/row widths before/after - not implemented yet
        Range('Sheet1', 'A1:D4').value = 'test_string'
        Range('Sheet1', 'A:D').autofit()
        Range('Sheet1', 'A:D').autofit('r')
        Range('Sheet1', 'A:D').autofit('c')
        Range('Sheet1', 'A:D').autofit('rows')
        Range('Sheet1', 'A:D').autofit('columns')

    def test_autofit_row(self):
        # TODO: compare col/row widths before/after - not implemented yet
        Range('Sheet1', 'A1:D4').value = 'test_string'
        Range('Sheet1', '1:1000000').autofit()
        Range('Sheet1', '1:1000000').autofit('r')
        Range('Sheet1', '1:1000000').autofit('c')
        Range('Sheet1', '1:1000000').autofit('rows')
        Range('Sheet1', '1:1000000').autofit('columns')

    def test_number_format_cell(self):
        format_string = "mm/dd/yy;@"
        Range('Sheet1', 'A1').number_format = format_string
        result = Range('Sheet1', 'A1').number_format
        assert_equal(format_string, result)

    def test_number_format_range(self):
        format_string = "mm/dd/yy;@"
        Range('Sheet1', 'A1:D4').number_format = format_string
        result = Range('Sheet1', 'A1:D4').number_format
        assert_equal(format_string, result)

    def test_get_address(self):
        res = Range((1, 1), (3, 3)).get_address()
        assert_equal(res, '$A$1:$C$3')

        res = Range((1, 1), (3, 3)).get_address(False)
        assert_equal(res, '$A1:$C3')

        res = Range((1, 1), (3, 3)).get_address(True, False)
        assert_equal(res, 'A$1:C$3')

        res = Range((1, 1), (3, 3)).get_address(False, False)
        assert_equal(res, 'A1:C3')

        res = Range((1, 1), (3, 3)).get_address(include_sheetname=True)
        assert_equal(res, 'Sheet1!$A$1:$C$3')

        res = Range('Sheet2', (1, 1), (3, 3)).get_address(include_sheetname=True)
        assert_equal(res, 'Sheet2!$A$1:$C$3')

        res = Range((1, 1), (3, 3)).get_address(external=True)
        assert_equal(res, '[test_range_1.xlsx]Sheet1!$A$1:$C$3')

    def test_hyperlink(self):
        address = 'www.xlwings.org'
        # Naked address
        Range('A1').add_hyperlink(address)
        assert_equal(Range('A1').value, address)
        hyperlink = Range('A1').hyperlink
        if not hyperlink.endswith('/'):
            hyperlink += '/'
        assert_equal(hyperlink, 'http://' + address + '/')

        # Address + FriendlyName
        Range('A2').add_hyperlink(address, 'test_link')
        assert_equal(Range('A2').value, 'test_link')
        hyperlink = Range('A2').hyperlink
        if not hyperlink.endswith('/'):
            hyperlink += '/'
        assert_equal(hyperlink, 'http://' + address + '/')

    def test_hyperlink_formula(self):
        Range('B10').formula = '=HYPERLINK("http://xlwings.org", "xlwings")'
        assert_equal(Range('B10').hyperlink, 'http://xlwings.org')

    def test_color(self):
        rgb = (30, 100, 200)

        Range('A1').color = rgb
        assert_equal(rgb, Range('A1').color)

        Range('A2').color = RgbColor.rgbAqua
        assert_equal((0, 255, 255), Range('A2').color)

        Range('A2').color = None
        assert_equal(Range('A2').color, None)

        Range('A1:D4').color = rgb
        assert_equal(rgb, Range('A1:D4').color)

    def test_size(self):
        assert_equal(Range('A1:C4').size, 12)

    def test_shape(self):
        assert_equal(Range('A1:C4').shape, (4, 3))

    def test_len(self):
        assert_equal(len(Range('A1:C4')), 4)

    def test_iterator(self):
        Range('A20').value = [[1., 2.], [3., 4.]]
        l = []

        r = Range('A20:B21')
        for i in r:
            l.append(i.value)

        assert_equal(l, [1., 2., 3., 4.])

        # check that reiterating on same range functions properly
        assert_equal([c.value for c in r], [1., 2., 3., 4.])

        Range('Sheet2', 'A20').value = [[1., 2.], [3., 4.]]
        l = []

        for i in Range('Sheet2', 'A20:B21'):
            l.append(i.value)

        assert_equal(l, [1., 2., 3., 4.])

    def test_resize(self):
        r = Range('A1').resize(4, 5)
        assert_equal(r.shape, (4, 5))

        r = Range('A1').resize(row_size=4)
        assert_equal(r.shape, (4, 1))

        r = Range('A1:B4').resize(column_size=5)
        assert_equal(r.shape, (4, 5))

        r = Range('A1:B4').resize(row_size=5)
        assert_equal(r.shape, (5, 2))

        r = Range('A1:B4').resize()
        assert_equal(r.shape, (4, 2))

        assert_raises(AssertionError, Range('A1:B4').resize, row_size=0)
        assert_raises(AssertionError, Range('A1:B4').resize, column_size=0)

    def test_offset(self):
        o = Range('A1:B3').offset(3, 4)
        assert_equal(o.get_address(), '$E$4:$F$6')

        o = Range('A1:B3').offset(row_offset=3)
        assert_equal(o.get_address(), '$A$4:$B$6')

        o = Range('A1:B3').offset(column_offset=4)
        assert_equal(o.get_address(), '$E$1:$F$3')

    def test_date(self):
        date_1 = date(2000, 12, 3)
        Range('X1').value = date_1
        date_2 = Range('X1').value
        assert_equal(date_1, date(date_2.year, date_2.month, date_2.day))

    def test_row(self):
        assert_equal(Range('B3:F5').row, 3)

    def test_column(self):
        assert_equal(Range('B3:F5').column, 2)

    def test_last_cell(self):
        assert_equal(Range('B3:F5').last_cell.row, 5)
        assert_equal(Range('B3:F5').last_cell.column, 6)

    def test_get_set_named_range(self):
        Range('A100').name = 'test1'
        assert_equal(Range('A100').name, 'test1')

        Range('A200:B204').name = 'test2'
        assert_equal(Range('A200:B204').name, 'test2')

    def test_integers(self):
        """Covers GH 227"""
        Range('A99').value = 2147483647  # max SInt32
        assert_equal(Range('A99').value, 2147483647)

        Range('A100').value = 2147483648  # SInt32 < x < SInt64
        assert_equal(Range('A100').value, 2147483648)

        Range('A101').value = 10000000000000000000  # long
        assert_equal(Range('A101').value, 10000000000000000000)

    def test_numpy_datetime(self):
        _skip_if_no_numpy()

        Range('A55').value = np.datetime64('2005-02-25T03:30Z')
        assert_equal(Range('A55').value, datetime(2005, 2, 25, 3, 30))

    def test_dataframe_timezone(self):
        _skip_if_no_pandas()

        dt = np.datetime64(1434149887000, 'ms')
        ix = pd.DatetimeIndex(data=[dt], tz='GMT')
        df = pd.DataFrame(data=[1], index=ix, columns=['A'])
        Range('A1').value = df
        assert_equal(Range('A2').value, datetime(2015, 6, 12, 22, 58, 7))

    def test_datetime_timezone(self):
        eastern = pytz.timezone('US/Eastern')
        dt_naive = datetime(2002, 10, 27, 6, 0, 0)
        dt_tz = eastern.localize(dt_naive)
        Range('F34').value = dt_tz
        assert_equal(Range('F34').value, dt_naive)

    @raises(IndexError)
    def test_zero_based_index1(self):
        Range((0, 1)).value = 123

    @raises(IndexError)
    def test_zero_based_index2(self):
        a = Range((1, 1), (1, 0)).value

    def test_dictionary(self):
        d = {'a': 1., 'b': 2.}
        Range('A1').value = d
        assert_equal(d, Range('A1:B2').options(dict).value)

    def test_write_to_multicell_range(self):
        Range('A1:B2').value = 5
        assert_equal(Range('A1:B2').value, [[5., 5.],[5., 5.]])

    # TODO: not yet implemented in xlwings
    # def test_range_clipping(self):
    #     Range('A1').options(expand=False).value = [[1., 2.], [3., 4.]]
    #     assert_equal(Range('A1:B2').value, [[1., None], [None, None]])

    def test_transpose(self):
        Range('A1').options(transpose=True).value = [[1., 2.], [3., 4.]]
        assert_equal(Range('A1:B2').value, [[1., 3.], [2., 4.]])


class TestPicture(TestBase):
    def setUp(self):
        super(TestPicture, self).setUp('test_chart_1.xlsx')

    def test_two_wkb(self):
        wb2 = Workbook(app_visible=False, app_target=APP_TARGET)
        pic1 = Picture.add(sheet=1, name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        pic2 = Picture.add(sheet=self.wb.sheet(1), name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic1.name, 'pic1')
        assert_equal(pic2.name, 'pic1')
        wb2.close()

    def test_name(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.name, 'pic1')

        pic.name = 'pic_new'
        assert_equal(pic.name, 'pic_new')

    def test_left(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.left, 0)

        pic.left = 20
        assert_equal(pic.left, 20)

    def test_top(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.left, 0)

        pic.top = 20
        assert_equal(pic.top, 20)

    def test_width(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        if PIL:
            assert_equal(pic.width, 60)
        else:
            assert_equal(pic.width, 100)

        pic.width = 50
        assert_equal(pic.width, 50)

    def test_picture_object(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        assert_equal(pic.name, Picture('pic1').name)

    def test_height(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        if PIL:
            assert_equal(pic.height, 60)
        else:
            assert_equal(pic.height, 100)

        pic.height = 50
        assert_equal(pic.height, 50)

    @raises(Exception)
    def test_delete(self):
        pic = Picture.add(name='pic1', filename=os.path.join(this_dir, 'sample_picture.png'))
        pic.delete()
        pic.name

    @raises(ShapeAlreadyExists)
    def test_duplicate(self):
        pic1 = Picture.add(os.path.join(this_dir, 'sample_picture.png'), name='pic1')
        pic2 = Picture.add(os.path.join(this_dir, 'sample_picture.png'), name='pic1')

    def test_picture_update(self):
        pic1 = Picture.add(os.path.join(this_dir, 'sample_picture.png'), name='pic1')
        pic1.update(os.path.join(this_dir, 'sample_picture.png'))


class TestPlot(TestBase):
    def setUp(self):
        super(TestPlot, self).setUp('test_chart_1.xlsx')

    def test_add_plot(self):
        _skip_if_no_matplotlib()

        fig = Figure(figsize=(8, 6))
        ax = fig.add_subplot(111)
        ax.plot([1, 2, 3, 4, 5])

        plot = Plot(fig)
        pic = plot.show('Plot1')
        assert_equal(pic.name, 'Plot1')

        plot.show('Plot2', sheet=2)
        pic2 = Picture(2, 'Plot2')
        assert_equal(pic2.name, 'Plot2')


class TestChart(TestBase):

    def setUp(self):
        super(TestChart, self).setUp('test_chart_1.xlsx')

    def test_add_keywords(self):
        name = 'My Chart'
        chart_type = ChartType.xlLine
        Range('A1').value = chart_data
        chart = Chart.add(chart_type=chart_type, name=name, source_data=Range('A1').table)

        chart_actual = Chart(name)
        name_actual = chart_actual.name
        chart_type_actual = chart_actual.chart_type
        assert_equal(name, name_actual)
        if sys.platform.startswith('win'):
            assert_equal(chart_type, chart_type_actual)
        else:
            assert_equal(kw.line_chart, chart_type_actual)

    def test_add_properties(self):
        name = 'My Chart'
        chart_type = ChartType.xlLine
        Range('Sheet2', 'A1').value = chart_data
        chart = Chart.add('Sheet2')
        chart.chart_type = chart_type
        chart.name = name
        chart.set_source_data(Range('Sheet2', 'A1').table)

        chart_actual = Chart('Sheet2', name)
        name_actual = chart_actual.name
        chart_type_actual = chart_actual.chart_type
        assert_equal(name, name_actual)
        if sys.platform.startswith('win'):
            assert_equal(chart_type, chart_type_actual)
        else:
            assert_equal(kw.line_chart, chart_type_actual)


if __name__ == '__main__':
    nose.main()
