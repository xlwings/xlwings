"""
xlwings - Make Excel fly!

Homepage and documentation: http://xlwings.org
See also: http://zoomeranalytics.com

Copyright (C) 2014, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import sys
import numbers
from . import PY3, xlplatform
from .constants import ChartType

# Optional imports
try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None


# Python 2 and 3 compatibility
if PY3:
    string_types = str
else:
    string_types = basestring

# Platform compatibility
time_types = xlplatform.time_types


class Workbook(object):
    """
    Workbook connects an Excel Workbook with Python. You can create a new connection from Python with

    * a new workbook: ``wb = Workbook()``
    * an existing workbook: ``wb = Workbook(r'C:\\path\\to\\file.xlsx')``

    If you want to create the connection from Excel through the xlwings VBA module, use:

    ``wb = Workbook()``

    Parameters
    ----------
    fullname : string, default None
        If you want to connect to an existing Excel file from Python, use the fullname, e.g:
        ``r'C:\\path\\to\\file.xlsx'``
    """
    def __init__(self, fullname=None):
        if fullname:
            if xlplatform.is_xl_object(fullname):
                self.xl_workbook = fullname
                self.xl_app = xlplatform.get_app(self.xl_workbook)
            else:
                self.fullname = fullname.lower()
                if xlplatform.is_file_open(self.fullname):
                    # Connect to an open Workbook
                    self.xl_app, self.xl_workbook = xlplatform.get_workbook(self.fullname)
                else:
                    # Open Excel and the Workbook
                    self.xl_app, self.xl_workbook = xlplatform.open_workbook(self.fullname)
        elif len(sys.argv) > 2 and sys.argv[2] == 'from_xl':
            # Connect to the workbook from which this code has been invoked
            self.fullname = sys.argv[1].lower()
            self.xl_app, self.xl_workbook = xlplatform.get_workbook(self.fullname)
        else:
            # Open Excel if necessary and create a new workbook
            self.xl_app, self.xl_workbook = xlplatform.new_workbook()

        self.name = xlplatform.get_workbook_name(self.xl_workbook)
        self.active_sheet = ActiveSheet(xl_workbook=self.xl_workbook)

        # Make the most recently created Workbook the default when creating Range objects directly
        global xl_workbook_latest
        xl_workbook_latest = self.xl_workbook
        
    @classmethod
    def current(cls):
        """
        Returns the workbook object which is currently active.
        """
        return cls(xl_workbook_latest)

    def activate(self, sheet):
        """
        Activates the given sheet.

        Parameters
        ----------
        sheet : string or integer
            Sheet name or index.
        """
        xlplatform.activate_sheet(self.xl_workbook, sheet)

    def get_selection(self, asarray=False):
        """
        Returns the currently selected Range from Excel as xlwings Range object.

        Parameters
        ----------
        asarray : boolean, default False
            returns a NumPy array where empty cells are shown as nan

        Returns
        -------
        xlwings Range object
        """
        return self.range(xlplatform.get_selection_address(self.xl_app), asarray=asarray)

    def range(self, *args, **kwargs):
        """
        The range method gets and sets the Range object with the following arguments::

            range('A1')          range('Sheet1', 'A1')          range(1, 'A1')
            range('A1:C3')       range('Sheet1', 'A1:C3')       range(1, 'A1:C3')
            range((1,2))         range('Sheet1, (1,2))          range(1, (1,2))
            range((1,1), (3,3))  range('Sheet1', (1,1), (3,3))  range(1, (1,1), (3,3))
            range('NamedRange')  range('Sheet1', 'NamedRange')  range(1, 'NamedRange')

        If no worksheet name is provided as first argument (as name or index),
        it will take the range from the active sheet.

        Please check the available methods/properties directly under the Range object.

        Parameters
        ----------
        asarray : boolean, default False
            returns a NumPy array where empty cells are shown as nan

        index : boolean, default True
            Includes the index when setting a Pandas DataFrame

        header : boolean, default True
            Includes the column headers when setting a Pandas DataFrame

        Returns
        -------
        Range
            xlwings Range object
        """
        return Range(*args, workbook=self.xl_workbook, **kwargs)

    def chart(self, *args, **kwargs):
        """
        The chart method gives access to the chart object and can be called with the following arguments::

            chart(1)            chart('Sheet1', 1)              chart(1, 1)
            chart('Chart 1')    chart('Sheet1', 'Chart 1')      chart(1, 'Chart 1')

        If no worksheet name is provided as first argument (as name or index),
        it will take the Chart from the active sheet.

        To insert a new Chart into Excel, create it as follows:

        wb.chart().add()

        Parameters
        ----------
        *args :
            Definition of sheet (optional) and chart in the above described combinations.
        """
        return Chart(*args, workbook=self.xl_workbook, **kwargs)

    def clear_contents(self, sheet=None):
        """
        Clears the content of a whole sheet but leaves the formatting.

        Parameters
        ----------
        sheet : string or integer, default None
            Sheet name or index. If sheet is None, the active sheet is used.
        """
        if sheet is None:
            sheet = self.active_sheet.index

        xlplatform.clear_contents_worksheet(self.xl_workbook, sheet)

    def clear(self, sheet=None):
        """
        Clears the content and formatting of a whole sheet.

        Parameters
        ----------
        sheet : string or integer, default None
            Sheet name or index. If sheet is None, the active sheet is used.
        """
        if sheet is None:
            sheet = self.active_sheet.index

        xlplatform.clear_worksheet(self.xl_workbook, sheet)

    def close(self):
        """Closes the Workbook without saving it"""
        xlplatform.close_workbook(self.xl_workbook)

    def __repr__(self):
        return "<xlwings.Workbook '{0}'>".format(self.name)


class ActiveSheet(object):
    """
    Returns an object that represents the active sheet. Supposed to be used from the Workbook object like so::

        wb = Workbook()
        wb.active_sheet.name

    Parameters
    ----------
    xl_workbook : pywin32 or appscript object
        Underlying Workbook object
    """
    def __init__(self, xl_workbook=None):
        if xl_workbook is None:
            xl_workbook = xl_workbook_latest
        self.xl_active_sheet = xlplatform.get_active_sheet(xl_workbook)

    @property
    def name(self):
        return xlplatform.get_workbook_name(self.xl_active_sheet)

    @property
    def index(self):
        return xlplatform.get_worksheet_index(self.xl_active_sheet)


class Range(object):
    """
    A Range object can be created with the following arguments::

        Range('A1')          Range('Sheet1', 'A1')          Range(1, 'A1')
        Range('A1:C3')       Range('Sheet1', 'A1:C3')       Range(1, 'A1:C3')
        Range((1,2))         Range('Sheet1, (1,2))          Range(1, (1,2))
        Range((1,1), (3,3))  Range('Sheet1', (1,1), (3,3))  Range(1, (1,1), (3,3))
        Range('NamedRange')  Range('Sheet1', 'NamedRange')  Range(1, 'NamedRange')

    If no worksheet name is provided as first argument (as name or index),
    it will take the Range from the active sheet.

    You usually want to go for ``Range(...).value`` to get the values (as list of lists).

    Parameters
    ----------
    *args :
        Definition of sheet (optional) and Range in the above described combinations.
    asarray : boolean, default False
        Returns a NumPy array (atleast_1d) where empty cells are transformed into nan.

    index : boolean, default True
        Includes the index when setting a Pandas DataFrame or Series.

    header : boolean, default True
        Includes the column headers when setting a Pandas DataFrame.

    atleast_2d : boolean, default False
        Returns 2d lists/arrays even if the Range is a Row or Column.
    """
    def __init__(self, *args, **kwargs):
        # Arguments
        if len(args) == 1 and isinstance(args[0], string_types):
            sheet_name_or_index = None
            range_address = args[0]
        elif len(args) == 1 and isinstance(args[0], tuple):
            sheet_name_or_index = None
            range_address = None
            self.row1 = args[0][0]
            self.col1 = args[0][1]
            self.row2 = self.row1
            self.col2 = self.col1
        elif (len(args) == 2
              and isinstance(args[0], (numbers.Number, string_types))
              and isinstance(args[1], string_types)):
            sheet_name_or_index = args[0]
            range_address = args[1]
        elif (len(args) == 2
              and isinstance(args[0], (numbers.Number, string_types))
              and isinstance(args[1], tuple)):
            sheet_name_or_index = args[0]
            range_address = None
            self.row1 = args[1][0]
            self.col1 = args[1][1]
            self.row2 = self.row1
            self.col2 = self.col1
        elif len(args) == 2 and isinstance(args[0], tuple):
            sheet_name_or_index = None
            range_address = None
            self.row1 = args[0][0]
            self.col1 = args[0][1]
            self.row2 = args[1][0]
            self.col2 = args[1][1]
        elif len(args) == 3:
            sheet_name_or_index = args[0]
            range_address = None
            self.row1 = args[1][0]
            self.col1 = args[1][1]
            self.row2 = args[2][0]
            self.col2 = args[2][1]

        # Keyword Arguments
        self.kwargs = kwargs
        self.index = kwargs.get('index', True)  # Set DataFrame with index
        self.header = kwargs.get('header', True)  # Set DataFrame with header
        self.asarray = kwargs.get('asarray', False)  # Return Data as NumPy Array
        self.strict = kwargs.get('strict', False)  # Stop table/horizontal/vertical at empty cells that contain formulas
        self.atleast_2d = kwargs.get('atleast_2d', False)  # Force data to be list of list or a 2d numpy array
        self.xl_workbook = kwargs.get('workbook', xl_workbook_latest)

        # Get sheet
        if sheet_name_or_index:
            self.xl_sheet = xlplatform.get_worksheet(self.xl_workbook, sheet_name_or_index)
        else:
            self.xl_sheet = xlplatform.get_active_sheet(self.xl_workbook)

        # Get xl_range object
        if range_address:
            self.row1 = xlplatform.get_first_row(self.xl_sheet, range_address)
            self.col1 = xlplatform.get_first_column(self.xl_sheet, range_address)
            self.row2 = self.row1 + xlplatform.count_rows(self.xl_sheet, range_address) - 1
            self.col2 = self.col1 + xlplatform.count_columns(self.xl_sheet, range_address) - 1

        self.xl_range = xlplatform.get_range_from_indices(self.xl_sheet, self.row1, self.col1, self.row2, self.col2)

    def is_cell(self):
        """
        Returns True if the Range consists of a single Cell otherwise False
        """
        if self.row1 == self.row2 and self.col1 == self.col2:
            return True
        else:
            return False

    def is_row(self):
        """
        Returns True if the Range consists of a single Row otherwise False
        """
        if self.row1 == self.row2 and self.col1 != self.col2:
            return True
        else:
            return False

    def is_column(self):
        """
        Returns True if the Range consists of a single Column otherwise False
        """
        if self.row1 != self.row2 and self.col1 == self.col2:
            return True
        else:
            return False

    def is_table(self):
        """
        Returns True if the Range consists of a 2d array otherwise False
        """
        if self.row1 != self.row2 and self.col1 != self.col2:
            return True
        else:
            return False

    @property
    def value(self):
        """
        Gets and sets the values for the given Range.

        Returns
        -------
        list or numpy array
            Empty cells are set to None. If ``asarray=True``,
            a numpy array is returned where empty cells are set to nan.
        """
        # TODO: refactor
        if self.is_cell():
            # Clean_xl_data requires and returns a list of list
            data = xlplatform.clean_xl_data([[xlplatform.get_value_from_range(self.xl_range)]])[0][0]
        elif self.is_row():
            data = xlplatform.clean_xl_data(xlplatform.get_value_from_range(self.xl_range))
            if not self.atleast_2d:
                data = data[0]
        elif self.is_column():
            data = xlplatform.clean_xl_data(xlplatform.get_value_from_range(self.xl_range))
            if not self.atleast_2d:
                data = [item for sublist in data for item in sublist]
        else:  # 2d Range, leave as list of list
            data = xlplatform.clean_xl_data(xlplatform.get_value_from_range(self.xl_range))

        # Return as NumPy Array
        if self.asarray:
            # replace None (empty cells) with nan as None produces arrays with dtype=object
            # TODO: easier like this: np.array(my_list, dtype=np.float)
            if data is None:
                data = np.nan
            if (self.is_column() or self.is_row()) and not self.atleast_2d:
                data = [np.nan if x is None else x for x in data]
            elif self.is_table() or self.atleast_2d:
                data = [[np.nan if x is None else x for x in i] for i in data]
            return np.atleast_1d(np.array(data))
        return data

    @value.setter
    def value(self, data):
        # Pandas DataFrame: Turn into NumPy object array with or without Index and Headers
        if hasattr(pd, 'DataFrame') and isinstance(data, pd.DataFrame):
            if self.index:
                data = data.reset_index()

            if self.header:
                if isinstance(data.columns, pd.MultiIndex):
                    # Ensure dtype=object because otherwise it may get assigned a string type which sometimes makes
                    # vstacking return a string array. This would cause values to be truncated and we can't easily
                    # transform np.nan in string form.
                    # Python 3 requires zip wrapped in list
                    columns = np.array(list(zip(*data.columns.tolist())), dtype=object)
                else:
                    columns = np.empty((data.columns.shape[0],), dtype=object)
                    columns[:] = np.array([data.columns.tolist()])
                data = np.vstack((columns, data.values))
            else:
                data = data.values

        # Pandas Series
        if hasattr(pd, 'Series') and isinstance(data, pd.Series):
            if self.index:
                data = data.reset_index().values
            else:
                data = data.values[:,np.newaxis]

        # NumPy array: nan have to be transformed to None, otherwise Excel shows them as 65535. This seems to be an
        # Excel bug, see: http://visualstudiomagazine.com/articles/2008/07/01/return-double-values-in-excel.aspx
        # Also, turn into list (Python 3 can't handle arrays directly)
        if hasattr(np, 'ndarray') and isinstance(data, np.ndarray):
            try:
                data = np.where(np.isnan(data), None, data)
                data = data.tolist()
            except TypeError:
                # isnan doesn't work on arrays of dtype=object
                if hasattr(pd, 'isnull'):
                    data[pd.isnull(data)] = None
                    data = data.tolist()
                else:
                    # expensive way of replacing nan with None in object arrays in case Pandas is not available
                    data = [[None if isinstance(c, float) and np.isnan(c) else c for c in row] for row in data]

        # Simple Lists: Turn into list of lists (np.nan is part of numbers.Number)
        if isinstance(data, list) and (isinstance(data[0], (numbers.Number, string_types, time_types))
                                       or data[0] is None):
            data = [data]

        # Get dimensions and prepare data for Excel
        # TODO: refactor
        if isinstance(data, (numbers.Number, string_types, time_types)) or data is None:
            # Single cells
            row2 = self.row2
            col2 = self.col2
            data = xlplatform.prepare_xl_data(data)
            try:
                # scalar np.nan need to be turned into None, otherwise Excel shows it as 65535 (same as for NumPy array)
                if hasattr(np, 'ndarray') and np.isnan(data):
                    data = None
            except TypeError:
                pass  # raised if data is not a np.nan

        else:
            # List of List
            row2 = self.row1 + len(data) - 1
            col2 = self.col1 + len(data[0]) - 1
            data = [[xlplatform.prepare_xl_data(c) for c in row] for row in data]

        xlplatform.set_value(xlplatform.get_range_from_indices(self.xl_sheet,
                                                               self.row1, self.col1, row2, col2), data)

    @property
    def formula(self):
        """
        Gets or sets the formula for the given Range.
        """
        return xlplatform.get_formula(self.xl_range)

    @formula.setter
    def formula(self, value):
        xlplatform.set_formula(self.xl_range, value)

    @property
    def table(self):
        """
        Returns a contiguous Range starting with the indicated cell as top-left corner and going down and right as
        long as no empty cell is hit.

        Parameters
        ----------
        strict : boolean, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        xlwings Range object

        Examples
        --------
        To get the values of a contiguous range or clear its contents use::

            Range('A1').table.value
            Range('A1').table.clear_contents()

        """
        row2 = Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), **self.kwargs).vertical.row2
        col2 = Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), **self.kwargs).horizontal.col2

        return Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def vertical(self):
        """
        Returns a contiguous Range starting with the indicated cell and going down as long as no empty cell is hit.
        This corresponds to ``Ctrl + Shift + Down Arrow`` in Excel.

        Parameters
        ----------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        xlwings Range object

        Examples
        --------
        To get the values of a contiguous range or clear its contents use::

            Range('A1').vertical.value
            Range('A1').vertical.clear_contents()

        """
        # A single cell is a special case as End(xlDown) jumps over adjacent empty cells
        if xlplatform.get_value_from_index(self.xl_sheet, self.row1 + 1, self.col1) in [None, ""]:
            row2 = self.row1
        else:
            row2 = xlplatform.get_row_index_end_down(self.xl_sheet, self.row1, self.col1)

        # Strict stops at cells that contain a formula but show an empty value
        if self.strict:
            row2 = self.row1
            while xlplatform.get_value_from_index(self.xl_sheet, row2 + 1, self.col1) not in [None, ""]:
                row2 += 1

        col2 = self.col2

        return Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def horizontal(self):
        """
        Returns a contiguous Range starting with the indicated cell and going right as long as no empty cell is hit.

        Parameters
        ----------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        xlwings Range object

        Examples
        --------
        To get the values of a contiguous range or clear its contents use::

            Range('A1').horizontal.value
            Range('A1').horizontal.clear_contents()

        """
        # A single cell is a special case as End(xlToRight) jumps over adjacent empty cells
        if xlplatform.get_value_from_index(self.xl_sheet, self.row1, self.col1 + 1) in [None, ""]:
            col2 = self.col1
        else:
            col2 = xlplatform.get_column_index_end_right(self.xl_sheet, self.row1, self.col1)

        # Strict: stops at cells that contain a formula but show an empty value
        if self.strict:
            col2 = self.col1
            while xlplatform.get_value_from_index(self.xl_sheet, self.row1, col2 + 1) not in [None, ""]:
                col2 += 1

        row2 = self.row2

        return Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), (row2, col2), **self.kwargs)

    @property
    def current_region(self):
        """
        The current_region property returns a Range object representing a range bounded by (but not including) any
        combination of blank rows and blank columns or the edges of the worksheet. It corresponds to ``Ctrl + *``.

        Returns
        -------
        xlwings Range object

        """
        address = xlplatform.get_current_region_address(self.xl_sheet, self.row1, self.col1)
        return Range(xlplatform.get_worksheet_name(self.xl_sheet), address, **self.kwargs)

    def clear(self):
        """
        Clears the content and the formatting of a Range.
        """
        xlplatform.clear_range(self.xl_range)

    def clear_contents(self):
        """
        Clears the content of a Range but leaves the formatting.
        """
        xlplatform.clear_contents_range(self.xl_range)

    def autofit(self, axis=None):
        """
        Autofits the width of either columns, rows or both.

        Parameters
        ----------
        axis : string or integer, default None
            - To autofit rows, use one of the following: 0 or 'rows' or 'r'
            - To autofit columns, use one of the following: 1 or 'columns' or 'c'
            - To autofit rows and columns, provide no arguments

        Examples
        --------
        ::

            # Autofit column A
            Range('A:A').autofit()
            # Autofit row 1
            Range('1:1').autofit()
            # Autofit columns and rows, taking into account Range('A1:E4')
            Range('A1:E4').autofit()
            # AutoFit columns, taking into account Range('A1:E4')
            Range('A1:E4').autofit(axis=1)
            # AutoFit rows, taking into account Range('A1:E4')
            Range('A1:E4').autofit('rows')

        """
        xlplatform.autofit(self, axis)

    def __repr__(self):
        return "<xlwings.Range of Workbook '{0}'>".format(xlplatform.get_workbook_name(self.xl_workbook))


class Chart(object):
    """
    A chart object that represents an existing Excel chart can be created with the following arguments::

        Chart(1)            Chart('Sheet1', 1)              Chart(1, 1)
        Chart('Chart 1')    Chart('Sheet1', 'Chart 1')      Chart(1, 'Chart 1')

    If no worksheet name is provided as first argument (as name or index),
    it will take the chart from the active sheet.

    To insert a new chart into Excel, create it as follows::

        Chart().add()

    Parameters
    ----------
    *args
        Definition of sheet (optional) and chart in the above described combinations.

    chart_type : Member of ChartType, default xlColumnClustered
        Chart type, can also be set using the ``chart_type`` property

    """
    def __init__(self, *args, **kwargs):
        # Use global Workbook if none provided
        self.xl_workbook = kwargs.get('workbook', xl_workbook_latest)

        # Arguments
        if len(args) == 0:
            pass
        elif len(args) > 0:
            if len(args) == 1:
                _sheet_name_or_index = xlplatform.get_worksheet_name(xlplatform.get_active_sheet(self.xl_workbook))
                _name_or_index = args[0]
            elif len(args) == 2:
                _sheet_name_or_index = args[0]
                _name_or_index = args[1]

            # Get xl_chart object
            self.xl_chart = xlplatform.get_chart_object(self.xl_workbook, _sheet_name_or_index, _name_or_index)
            self.index = xlplatform.get_chart_index(self.xl_chart)
            self.name = xlplatform.get_chart_name(self.xl_chart)

        # Chart Type
        chart_type = kwargs.get('chart_type')
        if chart_type:
            self.chart_type = chart_type

        # Source Data
        source_data = kwargs.get('source_data')
        if source_data:
            self.set_source_data(source_data)

    def add(self, sheet=None, left=168, top=217, width=355, height=211, **kwargs):
        """
        Inserts a new chart in Excel.

        Arguments
        ---------
        sheet : string or integer, default None
            Name or Index of the sheet, defaults to the active sheet

        left : float, default 100
            left position in points

        top : float, default 75
            top position in points

        width : float, default 375
            width in points

        height : float, default 225
            height in points

        chart_type : xlwings.ChartType member, default xlColumnClustered
            Excel chart type. E.g. xlwings.ChartType.xlLine

        name : str, default None
            Excel chart name. Defaults to Excel standard name if not provided, e.g. 'Chart 1'

        source_data : xlwings Range
            e.g. Range('A1').table

        """
        chart_type = kwargs.get('chart_type', ChartType.xlColumnClustered)
        name = kwargs.get('name')
        source_data = kwargs.get('source_data')

        if sheet is None:
            sheet = xlplatform.get_worksheet_index(xlplatform.get_active_sheet(self.xl_workbook))

        xl_chart = xlplatform.add_chart(self.xl_workbook, sheet, left, top, width, height)

        if name:
            xlplatform.set_chart_name(xl_chart, name)
        else:
            name = xlplatform.get_chart_name(xl_chart)

        return Chart(sheet, name, workbook=self.xl_workbook, chart_type=chart_type, source_data=source_data)

    @property
    def name(self):
        """
        Gets and sets the name of a chart
        """
        return xlplatform.get_chart_name(self.xl_chart)

    @name.setter
    def name(self, value):
        xlplatform.set_chart_name(self.xl_chart, value)

    @property
    def chart_type(self):
        """
        Gets and sets the chart type of a chart
        """
        return xlplatform.get_chart_type(self.xl_chart)

    @chart_type.setter
    def chart_type(self, value):
        xlplatform.set_chart_type(self.xl_chart, value)

    def activate(self):
        """
        Makes the chart the active chart.
        """
        xlplatform.activate_chart(self.xl_chart)

    def set_source_data(self, source):
        """
        Sets the source for the chart

        Arguments
        ---------
        source : Range
            xlwings Range object, e.g. ``Range('A1')``
        """
        xlplatform.set_source_data_chart(self.xl_chart, source.xl_range)

    def __repr__(self):
        return "<xlwings.Chart '{0}'>".format(self.name)

