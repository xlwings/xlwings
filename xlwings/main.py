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
from . import xlplatform, string_types, time_types
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


class Workbook(object):
    """
    Workbook connects an Excel Workbook with Python. You can create a new connection from Python with

    * a new workbook: ``wb = Workbook()``
    * an unsaved workbook: ``wb = Workbook('Book1')``
    * a saved workbook (open or closed): ``wb = Workbook(r'C:\\path\\to\\file.xlsx')``

    If you want to create the connection from Excel through the xlwings VBA module, use:

    ``wb = Workbook()``

    Arguments
    ---------
    fullname : string, default None
        If you want to connect to an existing Excel file from Python, use the fullname, e.g:
        ``r'C:\\path\\to\\file.xlsx'``. If the file is unsaved, then use the name only, e.g.: ``Book1``
    """
    def __init__(self, fullname=None, xl_workbook=None):
        if xl_workbook:
            self.xl_workbook = xl_workbook
            self.xl_app = xlplatform.get_app(self.xl_workbook)
        elif fullname:
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
        self.active_sheet = Sheet.active(workbook=self)

        # Make the most recently created Workbook the default when creating Range objects directly
        xlplatform.set_xl_workbook_latest(self.xl_workbook)
        
    @classmethod
    def current(cls):
        """
        Returns the current workbook object, i.e. the default workbook used by Range, Sheet and Chart if not specified
        otherwise. On Windows, it also means that Workbook.new() Workbook.open() are acting on the same instance of
        Excel as this Workbook.
        """
        return cls(xl_workbook=xlplatform.get_xl_workbook_latest())

    def set_current(self):
        """
        This makes the workbook the default Workbook that Range, Sheet and Chart use if not specified
        otherwise. On Windows, it also means that Workbook.new() Workbook.open() are acting on the same instance of
        Excel as this Workbook.
        """
        xlplatform.set_xl_workbook_latest(self.xl_workbook)

    def get_selection(self, asarray=False):
        """
        Returns the currently selected Range from Excel as Range object.

        Keyword Arguments
        -----------------
        asarray : boolean, default False
            returns a NumPy array where empty cells are shown as nan

        Returns
        -------
        Range object
        """
        return Range(xlplatform.get_selection_address(self.xl_app), workbook=self, asarray=asarray)

    def close(self):
        """Closes the Workbook without saving it"""
        xlplatform.close_workbook(self.xl_workbook)

    def __repr__(self):
        return "<Workbook '{0}'>".format(self.name)


class Sheet(object):
    """
    Represents a Sheet of the current Workbook. Either call it with the Sheet name or index::

        Sheet('Sheet1')
        Sheet(1)

    Arguments
    ---------
    sheet : str or int
        Sheet name or index

    Keyword Arguments
    -----------------
    workbook : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via Workbook.set_current().
    """

    def __init__(self, sheet, workbook=None):
        if workbook is None:
            self.xl_workbook = xlplatform.get_xl_workbook_latest()
        else:
            self.xl_workbook = workbook.xl_workbook
        self.sheet = sheet
        self.xl_sheet = xlplatform.get_xl_sheet(self.xl_workbook, self.sheet)

    def activate(self):
        """
        Activates the given sheet.

        Arguments
        ---------
        sheet : string or integer
            Sheet name or index.
        """
        xlplatform.activate_sheet(self.xl_workbook, self.sheet)

    def clear_contents(self):
        """Clears the content of the whole sheet but leaves the formatting."""
        xlplatform.clear_contents_worksheet(self.xl_workbook, self.sheet)

    def clear(self):
        """Clears the content and formatting of the whole sheet."""
        xlplatform.clear_worksheet(self.xl_workbook, self.sheet)

    @property
    def name(self):
        """Get or set the name of the Sheet."""
        return xlplatform.get_worksheet_name(self.xl_sheet)

    @name.setter
    def name(self, value):
        xlplatform.set_worksheet_name(self.xl_sheet, value)

    @property
    def index(self):
        """Returns the index of the Sheet."""
        return xlplatform.get_worksheet_index(self.xl_sheet)

    @classmethod
    def active(cls, workbook=None):
        """Returns the workbook object which is currently active."""
        if workbook is None:
            xl_workbook = xlplatform.get_xl_workbook_latest()
        else:
            xl_workbook = workbook.xl_workbook
        return cls(xlplatform.get_worksheet_name(xlplatform.get_active_sheet(xl_workbook)), workbook)

    def __repr__(self):
        return "<Sheet '{0}'>".format(self.name)


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

    Arguments
    ---------
    *args :
        Definition of sheet (optional) and Range in the above described combinations.

    Keyword Arguments
    -----------------
    asarray : boolean, default False
        Returns a NumPy array (atleast_1d) where empty cells are transformed into nan.

    index : boolean, default True
        Includes the index when setting a Pandas DataFrame or Series.

    header : boolean, default True
        Includes the column headers when setting a Pandas DataFrame.

    atleast_2d : boolean, default False
        Returns 2d lists/arrays even if the Range is a Row or Column.

    workbook : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via Workbook.set_current().
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
        self.workbook = kwargs.get('workbook', None)
        if self.workbook is None:
            self.xl_workbook = xlplatform.get_xl_workbook_latest()
        else:
            self.xl_workbook = self.workbook.xl_workbook
        self.index = kwargs.get('index', True)  # Set DataFrame with index
        self.header = kwargs.get('header', True)  # Set DataFrame with header
        self.asarray = kwargs.get('asarray', False)  # Return Data as NumPy Array
        self.strict = kwargs.get('strict', False)  # Stop table/horizontal/vertical at empty cells that contain formulas
        self.atleast_2d = kwargs.get('atleast_2d', False)  # Force data to be list of list or a 2d numpy array

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
            data = xlplatform.clean_xl_data([[xlplatform.get_value_from_range(self.xl_range)]])
            if not self.atleast_2d:
                data = data[0][0]
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

        Keyword Arguments
        -----------------
        strict : boolean, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        Range object

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

        Arguments
        ---------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        Range object

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

        Keyword Arguments
        -----------------
        strict : bool, default False
            strict stops the table at empty cells even if they contain a formula. Less efficient than if set to False.

        Returns
        -------
        Range object

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
        combination of blank rows and blank columns or the edges of the worksheet. It corresponds to ``Ctrl + *`` on
        Windows and ``Shift-Ctrl-Space``on Mac.

        Returns
        -------
        Range object

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

        Arguments
        ---------
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
        return "<Range of Workbook '{0}'>".format(xlplatform.get_workbook_name(self.xl_workbook))


class Chart(object):
    """
    A chart object that represents an existing Excel chart can be created with the following arguments::

        Chart(1)            Chart('Sheet1', 1)              Chart(1, 1)
        Chart('Chart 1')    Chart('Sheet1', 'Chart 1')      Chart(1, 'Chart 1')

    If no worksheet name is provided as first argument (as name or index),
    it will take the chart from the active sheet.

    To insert a new chart into Excel, create it as follows::

        Chart.add()

    Arguments
    ---------
    *args
        Definition of sheet (optional) and chart in the above described combinations.

    Keyword Arguments
    -----------------
    chart_type : Member of ChartType, default xlColumnClustered
        Chart type, can also be set using the ``chart_type`` property

    workbook : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via Workbook.set_current().
    """
    def __init__(self, *args, **kwargs):
        # Use global Workbook if none provided
        self.workbook = kwargs.get('workbook', None)
        if self.workbook is None:
            self.xl_workbook = xlplatform.get_xl_workbook_latest()
        else:
            self.xl_workbook = self.workbook.xl_workbook

        # Arguments
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

    @classmethod
    def add(cls, sheet=None, left=168, top=217, width=355, height=211, **kwargs):
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

        Keyword Arguments
        -----------------
        chart_type : xlwings.ChartType member, default xlColumnClustered
            Excel chart type. E.g. xlwings.ChartType.xlLine

        name : str, default None
            Excel chart name. Defaults to Excel standard name if not provided, e.g. 'Chart 1'

        source_data : Range
            e.g. Range('A1').table

        workbook : Workbook object, default Workbook.current()
            Defaults to the Workbook that was instantiated last or set via Workbook.set_current().
        """
        workbook = kwargs.get('workbook', None)
        if workbook is None:
            xl_workbook = xlplatform.get_xl_workbook_latest()
        else:
            xl_workbook = workbook.xl_workbook

        chart_type = kwargs.get('chart_type', ChartType.xlColumnClustered)
        name = kwargs.get('name')
        source_data = kwargs.get('source_data')

        if sheet is None:
            sheet = xlplatform.get_worksheet_index(xlplatform.get_active_sheet(xl_workbook))

        xl_chart = xlplatform.add_chart(xl_workbook, sheet, left, top, width, height)

        if name:
            xlplatform.set_chart_name(xl_chart, name)
        else:
            name = xlplatform.get_chart_name(xl_chart)

        return cls(sheet, name, workbook=workbook, chart_type=chart_type, source_data=source_data)

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
            Range object, e.g. ``Range('A1')``
        """
        xlplatform.set_source_data_chart(self.xl_chart, source.xl_range)

    def __repr__(self):
        return "<Chart '{0}' on Workbook '{1}'>".format(self.name, self.workbook.name)

