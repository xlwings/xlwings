"""
xlwings - Make Excel fly with Python!

Homepage and documentation: http://xlwings.org
See also: http://zoomeranalytics.com

Copyright (C) 2014-2016, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import os
import sys
import re
import numbers
import itertools
import inspect
import collections
import tempfile
import shutil

from . import xlplatform, string_types, xrange, map, ShapeAlreadyExists
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

try:
    from matplotlib.backends.backend_agg import FigureCanvas
except ImportError:
    FigureCanvas = None

try:
    from PIL import Image
except ImportError:
    Image = None


class Application(object):
    """
    Application is dependent on the Workbook since there might be different application instances on Windows.
    """

    def __init__(self, wkb):
        self.wkb = wkb
        self.xl_app = wkb.xl_app

    @property
    def version(self):
        """
        Returns Excel's version string.

        .. versionadded:: 0.5.0
        """
        return xlplatform.get_app_version_string(self.wkb.xl_workbook)

    def quit(self):
        """
        Quits the application without saving any workbooks.

        .. versionadded:: 0.3.3
        """
        xlplatform.quit_app(self.xl_app)

    @property
    def screen_updating(self):
        """
        True if screen updating is turned on. Read/write Boolean.

        .. versionadded:: 0.3.3
        """
        return xlplatform.get_screen_updating(self.xl_app)

    @screen_updating.setter
    def screen_updating(self, value):
        xlplatform.set_screen_updating(self.xl_app, value)

    @property
    def visible(self):
        """
        Gets or sets the visibility of Excel to ``True`` or  ``False``. This property can also be
        conveniently set during instantiation of a new Workbook: ``Workbook(app_visible=False)``

        .. versionadded:: 0.3.3
        """
        return xlplatform.get_visible(self.xl_app)

    @visible.setter
    def visible(self, value):
        xlplatform.set_visible(self.xl_app, value)

    @property
    def calculation(self):
        """
        Returns or sets a Calculation value that represents the calculation mode.

        Example
        -------
        >>> from xlwings import Workbook, Application
        >>> from xlwings.constants import Calculation
        >>> wb = Workbook()
        >>> Application(wkb=wb).calculation = Calculation.xlCalculationManual

        .. versionadded:: 0.3.3
        """
        return xlplatform.get_calculation(self.xl_app)

    @calculation.setter
    def calculation(self, value):
        xlplatform.set_calculation(self.xl_app, value)

    def calculate(self):
        """
        Calculates all open Workbooks

        .. versionadded:: 0.3.6
        """
        xlplatform.calculate(self.xl_app)


class Workbook(object):
    """
    ``Workbook`` connects an Excel Workbook with Python. You can create a new connection from Python with

    * a new workbook: ``wb = Workbook()``
    * the active workbook: ``wb = Workbook.active()``
    * an unsaved workbook: ``wb = Workbook('Book1')``
    * a saved (open) workbook by name (incl. xlsx etc): ``wb = Workbook('MyWorkbook.xlsx')``
    * a saved (open or closed) workbook by path: ``wb = Workbook(r'C:\\path\\to\\file.xlsx')``

    Keyword Arguments
    -----------------
    fullname : str, default None
        Full path or name (incl. xlsx, xlsm etc.) of existing workbook or name of an unsaved workbook.

    xl_workbook : pywin32 or appscript Workbook object, default None
        This enables to turn existing Workbook objects of the underlying libraries into xlwings objects

    app_visible : boolean, default True
        The resulting Workbook will be visible by default. To open it without showing a window,
        set ``app_visible=False``. Or, to not alter the visibility (e.g., if Excel is already running),
        set ``app_visible=None``. Note that this property acts on the whole Excel instance, not just the
        specific Workbook.

    app_target : str, default None
        Mac-only, use the full path to the Excel application,
        e.g. ``/Applications/Microsoft Office 2011/Microsoft Excel`` or ``/Applications/Microsoft Excel``

        On Windows, if you want to change the version of Excel that xlwings talks to, go to ``Control Panel >
        Programs and Features`` and ``Repair`` the Office version that you want as default.


    To create a connection when the Python function is called from Excel, use:

    ``wb = Workbook.caller()``

    """

    def __init__(self, fullname=None, xl_workbook=None, app_visible=True, app_target=None):
        if xl_workbook:
            self.xl_workbook = xl_workbook
            self.xl_app = xlplatform.get_app(self.xl_workbook, app_target)
        elif fullname:
            self.fullname = fullname
            if not os.path.isfile(fullname) or xlplatform.is_file_open(self.fullname):
                # Connect to unsaved Workbook (e.g. 'Workbook1') or to an opened Workbook
                self.xl_app, self.xl_workbook = xlplatform.get_open_workbook(self.fullname, app_target)
            else:
                # Open Excel and the Workbook
                self.xl_app, self.xl_workbook = xlplatform.open_workbook(self.fullname, app_target)
        else:
            # Open Excel if necessary and create a new workbook
            self.xl_app, self.xl_workbook = xlplatform.new_workbook(app_target)

        self.name = xlplatform.get_workbook_name(self.xl_workbook)

        if fullname is None:
            self.fullname = xlplatform.get_fullname(self.xl_workbook)

        # Make the most recently created Workbook the default when creating Range objects directly
        xlplatform.set_xl_workbook_current(self.xl_workbook)

        if app_visible is not None:
            xlplatform.set_visible(self.xl_app, app_visible)

    @property
    def active_sheet(self):
        return Sheet.active(wkb=self)

    @classmethod
    def active(cls, app_target=None):
        """
        Returns the Workbook that is currently active or has been active last. On Windows,
        this works across all instances.

        .. versionadded:: 0.4.1
        """
        xl_workbook = xlplatform.get_active_workbook(app_target=app_target)
        return cls(xl_workbook=xl_workbook, app_target=app_target)

    @classmethod
    def caller(cls):
        """
        Creates a connection when the Python function is called from Excel:

        ``wb = Workbook.caller()``

        Always pack the ``Workbook`` call into the function being called from Excel, e.g.:

        .. code-block:: python

             def my_macro():
                wb = Workbook.caller()
                Range('A1').value = 1

        To be able to easily invoke such code from Python for debugging, use ``Workbook.set_mock_caller()``.

        .. versionadded:: 0.3.0
        """
        if hasattr(Workbook, '_mock_file'):
            # Use mocking Workbook, see Workbook.set_mock_caller()
            _, xl_workbook = xlplatform.get_open_workbook(Workbook._mock_file)
            return cls(xl_workbook=xl_workbook)
        elif len(sys.argv) > 2 and sys.argv[2] == 'from_xl':
            # Connect to the workbook from which this code has been invoked
            fullname = sys.argv[1].lower()
            if sys.platform.startswith('win'):
                xl_app, xl_workbook = xlplatform.get_open_workbook(fullname, hwnd=sys.argv[4])
                return cls(xl_workbook=xl_workbook)
            else:
                xl_app, xl_workbook = xlplatform.get_open_workbook(fullname, app_target=sys.argv[3])
                return cls(xl_workbook=xl_workbook, app_target=sys.argv[3])
        elif xlplatform.get_xl_workbook_current():
            # Called through ExcelPython connection
            return cls(xl_workbook=xlplatform.get_xl_workbook_current())
        else:
            raise Exception('Workbook.caller() must not be called directly. Call through Excel or set a mock caller '
                            'first with Workbook.set_mock_caller().')

    @staticmethod
    def set_mock_caller(fullpath):
        """
        Sets the Excel file which is used to mock ``Workbook.caller()`` when the code is called from within Python.

        Examples
        --------
        ::

            # This code runs unchanged from Excel and Python directly
            import os
            from xlwings import Workbook, Range

            def my_macro():
                wb = Workbook.caller()
                Range('A1').value = 'Hello xlwings!'

            if __name__ == '__main__':
                # Mock the calling Excel file
                Workbook.set_mock_caller(r'C:\\path\\to\\file.xlsx')
                my_macro()

        .. versionadded:: 0.3.1
        """
        Workbook._mock_file = fullpath

    @classmethod
    def current(cls):
        """
        Returns the current Workbook object, i.e. the default Workbook used by ``Sheet``, ``Range`` and ``Chart`` if not
        specified otherwise. On Windows, in case there are various instances of Excel running, opening an existing or
        creating a new Workbook through ``Workbook()`` is acting on the same instance of Excel as this Workbook. Use
        like this: ``Workbook.current()``.

        .. versionadded:: 0.2.2
        """
        return cls(xl_workbook=xlplatform.get_xl_workbook_current(), app_visible=None)

    def set_current(self):
        """
        This makes the Workbook the default that ``Sheet``, ``Range`` and ``Chart`` use if not specified
        otherwise. On Windows, in case there are various instances of Excel running, opening an existing or creating a
        new Workbook through ``Workbook()`` is acting on the same instance of Excel as this Workbook.

        .. versionadded:: 0.2.2
        """
        xlplatform.set_xl_workbook_current(self.xl_workbook)

    def get_selection(self):
        """
        Returns the currently selected cells from Excel as ``Range`` object.

        Example
        -------

        >>> import xlwings as xw
        >>> wb = xw.Workbook.active()
        >>> wb.get_selection()
        <Range on Sheet 'Sheet1' of Workbook 'Workbook1'>
        >>> wb.get_selection.value
        [[1.0, 2.0], [3.0, 4.0]]
        >>> wb.get_selection().options(transpose=True).value
        [[1.0, 3.0], [2.0, 4.0]]

        Returns
        -------
        Range object
        """
        return Range(xlplatform.get_selection_address(self.xl_app), wkb=self)

    def close(self):
        """
        Closes the Workbook without saving it.

        .. versionadded:: 0.1.1
        """
        xlplatform.close_workbook(self.xl_workbook)

    def save(self, path=None):
        """
        Saves the Workbook. If a path is being provided, this works like SaveAs() in Excel. If no path is specified and
        if the file hasn't been saved previously, it's being saved in the current working directory with the current
        filename. Existing files are overwritten without prompting.

        Arguments
        ---------
        path : str, default None
            Full path to the workbook

        Example
        -------
        >>> from xlwings import Workbook
        >>> wb = Workbook()
        >>> wb.save()
        >>> wb.save(r'C:\\path\\to\\new_file_name.xlsx')

        .. versionadded:: 0.3.1
        """
        xlplatform.save_workbook(self.xl_workbook, path)

    @staticmethod
    def get_xl_workbook(wkb):
        """
        Returns the ``xl_workbook_current`` if ``wkb`` is ``None``, otherwise the ``xl_workbook`` of ``wkb``. On Windows,
        ``xl_workbook`` is a pywin32 COM object, on Mac it's an appscript object.

        Arguments
        ---------
        wkb : Workbook or None
            Workbook object
        """
        if wkb is None and xlplatform.get_xl_workbook_current() is None:
            raise NameError('You must first instantiate a Workbook object.')
        elif wkb is None:
            xl_workbook = xlplatform.get_xl_workbook_current()
        else:
            xl_workbook = wkb.xl_workbook
        return xl_workbook

    @staticmethod
    def open_template():
        """
        Creates a new Excel file with the xlwings VBA module already included. This method must be called from an
        interactive Python shell::

        >>> Workbook.open_template()

        .. versionadded:: 0.3.3
        """
        this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))
        template_file = 'xlwings_template.xltm'
        try:
            os.remove(os.path.join(this_dir, '~$' + template_file))
        except OSError:
            pass

        xlplatform.open_template(os.path.realpath(os.path.join(this_dir, template_file)))

    @property
    def names(self):
        """
        A collection of all the (platform-specific) name objects in the application or workbook.
        Each name object represents a defined name for a range of cells (built-in or custom ones).

        .. versionadded:: 0.4.0
        """
        names = NamesDict(self.xl_workbook)
        xlplatform.set_names(self.xl_workbook, names)
        return names

    def macro(self, name):
        """
        Runs a Sub or Function in Excel VBA.

        Arguments:
        ----------
        name : Name of Sub or Function with or without module name, e.g. ``'Module1.MyMacro'`` or ``'MyMacro'``

        Examples:
        ---------
        This VBA function:

        .. code-block:: vb

            Function MySum(x, y)
                MySum = x + y
            End Function

        can be accessed like this:

        >>> wb = xw.Workbook.active()
        >>> my_sum = wb.macro('MySum')
        >>> my_sum(1, 2)
        3


        .. versionadded:: 0.7.1
        """
        return Macro(name, self)

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
    wkb : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.


    .. versionadded:: 0.2.3
    """

    def __init__(self, sheet, wkb=None):
        self.xl_workbook = Workbook.get_xl_workbook(wkb)
        self.sheet = sheet
        self.xl_sheet = xlplatform.get_xl_sheet(self.xl_workbook, self.sheet)

    def activate(self):
        """Activates the sheet."""
        xlplatform.activate_sheet(self.xl_workbook, self.sheet)

    def autofit(self, axis=None):
        """
        Autofits the width of either columns, rows or both on a whole Sheet.

        Arguments
        ---------
        axis : string, default None
            - To autofit rows, use one of the following: ``rows`` or ``r``
            - To autofit columns, use one of the following: ``columns`` or ``c``
            - To autofit rows and columns, provide no arguments

        Examples
        --------
        ::

            # Autofit columns
            Sheet('Sheet1').autofit('c')
            # Autofit rows
            Sheet('Sheet1').autofit('r')
            # Autofit columns and rows
            Range('Sheet1').autofit()

        .. versionadded:: 0.2.3
        """
        xlplatform.autofit_sheet(self, axis)

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
    def active(cls, wkb=None):
        """Returns the active Sheet. Use like so: ``Sheet.active()``"""
        xl_workbook = Workbook.get_xl_workbook(wkb)
        return cls(xlplatform.get_worksheet_name(xlplatform.get_active_sheet(xl_workbook)), wkb)

    @classmethod
    def add(cls, name=None, before=None, after=None, wkb=None):
        """
        Creates a new worksheet: the new worksheet becomes the active sheet. If neither ``before`` nor
        ``after`` is specified, the new Sheet will be placed at the end.

        Arguments
        ---------
        name : str, default None
            Sheet name, defaults to Excel standard name

        before : str or int, default None
            Sheet name or index

        after : str or int, default None
            Sheet name or index

        Returns
        -------
        Sheet object

        Examples
        --------

        >>> Sheet.add()  # Place at end with default name
        >>> Sheet.add('NewSheet', before='Sheet1')  # Include name and position
        >>> new_sheet = Sheet.add(after=3)
        >>> new_sheet.index
        4

        .. versionadded:: 0.2.3
        """
        xl_workbook = Workbook.get_xl_workbook(wkb)

        if before is None and after is None:
            after = Sheet(Sheet.count(wkb=wkb), wkb=wkb)
        elif before:
            before = Sheet(before, wkb=wkb)
        elif after:
            after = Sheet(after, wkb=wkb)

        if name:
            if name.lower() in [i.name.lower() for i in Sheet.all(wkb=wkb)]:
                raise Exception('That sheet name is already taken.')
            else:
                xl_sheet = xlplatform.add_sheet(xl_workbook, before, after)
                xlplatform.set_worksheet_name(xl_sheet, name)
                return cls(name, wkb)
        else:
            xl_sheet = xlplatform.add_sheet(xl_workbook, before, after)
            return cls(xlplatform.get_worksheet_name(xl_sheet), wkb)

    @staticmethod
    def count(wkb=None):
        """
        Counts the number of Sheets.

        Keyword Arguments
        -----------------
        wkb : Workbook object, default Workbook.current()
            Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.

        Examples
        --------
        >>> Sheet.count()
        3

        .. versionadded:: 0.2.3
        """
        xl_workbook = Workbook.get_xl_workbook(wkb)
        return xlplatform.count_worksheets(xl_workbook)

    @staticmethod
    def all(wkb=None):
        """
        Returns a list with all Sheet objects.

        Keyword Arguments
        -----------------
        wkb : Workbook object, default Workbook.current()
            Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.

        Examples
        --------
        >>> Sheet.all()
        [<Sheet 'Sheet1' of Workbook 'Book1'>, <Sheet 'Sheet2' of Workbook 'Book1'>]
        >>> [i.name.lower() for i in Sheet.all()]
        ['sheet1', 'sheet2']
        >>> [i.autofit() for i in Sheet.all()]

        .. versionadded:: 0.2.3
        """
        xl_workbook = Workbook.get_xl_workbook(wkb)
        sheet_list = []
        for i in range(1, xlplatform.count_worksheets(xl_workbook) + 1):
            sheet_list.append(Sheet(i, wkb=wkb))

        return sheet_list

    def delete(self):
        """
        Deletes the Sheet.

        .. versionadded: 0.6.0
        """
        xlplatform.delete_sheet(self)

    def __repr__(self):
        return "<Sheet '{0}' of Workbook '{1}'>".format(self.name, xlplatform.get_workbook_name(self.xl_workbook))


class Range(object):
    """
    Range(*args, wkb=None)

    A Range object can be instantiated with the following arguments::

        Range('A1')          Range('Sheet1', 'A1')          Range(1, 'A1')
        Range('A1:C3')       Range('Sheet1', 'A1:C3')       Range(1, 'A1:C3')
        Range((1,2))         Range('Sheet1, (1,2))          Range(1, (1,2))
        Range((1,1), (3,3))  Range('Sheet1', (1,1), (3,3))  Range(1, (1,1), (3,3))
        Range('NamedRange')  Range('Sheet1', 'NamedRange')  Range(1, 'NamedRange')

    The Sheet can also be provided as Sheet object::

        sh = Sheet(1)
        Range(sh, 'A1')

    If no worksheet name is provided as first argument, it will take the Range from the active sheet.

    You usually want to go for ``Range(...).value`` to get the values (as list of lists). You can
    influence the reading/writing behavior by making use of :ref:`converters` and their options:
    ``Range(...).options(...).value``

    Arguments
    ---------
    *args :
        Definition of sheet (optional) and Range in the above described combinations.

    Keyword Arguments
    -----------------
    wkb : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via `Workbook.set_current()``.
    """

    def __init__(self, *args, **options):

        self.workbook = options.pop('wkb', None)

        # Arguments
        if xlplatform.is_range_instance(args[0]):
            self.xl_range = args[0]
            self.xl_sheet = xlplatform.get_range_sheet(self.xl_range)
            self.xl_workbook = xlplatform.get_sheet_workbook(self.xl_sheet)
            self.workbook = Workbook(xl_workbook=self.xl_workbook)
            self.row1, self.col1, self.row2, self.col2 = xlplatform.get_range_coordinates(self.xl_range)
            range_address = None
        elif len(args) == 1 and isinstance(args[0], string_types):
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
              and isinstance(args[0], (numbers.Number, string_types, Sheet))
              and isinstance(args[1], string_types)):
            if isinstance(args[0], Sheet):
                sheet_name_or_index = args[0].index
            else:
                sheet_name_or_index = args[0]
            range_address = args[1]
        elif (len(args) == 2
              and isinstance(args[0], (numbers.Number, string_types, Sheet))
              and isinstance(args[1], tuple)):
            if isinstance(args[0], Sheet):
                sheet_name_or_index = args[0].index
            else:
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
            if isinstance(args[0], Sheet):
                sheet_name_or_index = args[0].index
            else:
                sheet_name_or_index = args[0]
            range_address = None
            self.row1 = args[1][0]
            self.col1 = args[1][1]
            self.row2 = args[2][0]
            self.col2 = args[2][1]

        # Keyword Arguments
        self._options = options
        if self.workbook is None and xlplatform.get_xl_workbook_current() is None:
            raise NameError('You must first instantiate a Workbook object.')
        elif self.workbook is None:
            self.xl_workbook = xlplatform.get_xl_workbook_current()
        else:
            self.xl_workbook = self.workbook.xl_workbook

        # Get sheet
        if not hasattr(self, 'xl_sheet'):
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

        if 0 in (self.row1, self.col1, self.row2, self.col2):
            raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")

        if not hasattr(self, 'xl_range'):
            self.xl_range = xlplatform.get_range_from_indices(self.xl_sheet, self.row1, self.col1, self.row2, self.col2)

    def __iter__(self):
        # Iterator object that returns cell Ranges: (1, 1), (1, 2) etc.
        return map(lambda cell: Range(xlplatform.get_worksheet_name(self.xl_sheet), cell, wkb=self.workbook, **self._options),
                   itertools.product(xrange(self.row1, self.row2 + 1), xrange(self.col1, self.col2 + 1)))

    def options(self, convert=None, **options):
        """
        Allows you to set a converter and their options. Converters define how Excel Ranges and their values are
        being converted both during reading and writing operations. If no explicit converter is specified, the
        base converter is being applied, see :ref:`converters`.

        Arguments
        ---------
        ``convert`` : object, default None
            A converter, e.g. ``dict``, ``np.array``, ``pd.DataFrame``, ``pd.Series``, defaults to default converter

        Keyword Arguments
        -----------------
        ndim : int, default None
            number of dimensions

        numbers : type, default None
            type of numbers, e.g. ``int``

        dates : type, default None
            e.g. ``datetime.date`` defaults to ``datetime.datetime``

        empty : object, default None
            transformation of empty cells

        transpose : Boolean, default False
            transpose values

        expand : str, default None
            One of ``'table'``, ``'vertical'``, ``'horizontal'``, see also ``Range.table`` etc

         => For converter-specific options, see :ref:`converters`.

        Returns
        -------
        Range object


        .. versionadded:: 0.7.0
        """
        options['convert'] = convert
        return Range(
            xlplatform.get_worksheet_name(self.xl_sheet),
            (self.row1, self.col1),
            (self.row2, self.col2),
            **options
        )

    def is_cell(self):
        """
        Returns ``True`` if the Range consists of a single Cell otherwise ``False``.

        .. versionadded:: 0.1.1
        """
        if self.row1 == self.row2 and self.col1 == self.col2:
            return True
        else:
            return False

    def is_row(self):
        """
        Returns ``True`` if the Range consists of a single Row otherwise ``False``.

        .. versionadded:: 0.1.1
        """
        if self.row1 == self.row2 and self.col1 != self.col2:
            return True
        else:
            return False

    def is_column(self):
        """
        Returns ``True`` if the Range consists of a single Column otherwise ``False``.

        .. versionadded:: 0.1.1
        """
        if self.row1 != self.row2 and self.col1 == self.col2:
            return True
        else:
            return False

    def is_table(self):
        """
        Returns ``True`` if the Range consists of a 2d array otherwise ``False``.

        .. versionadded:: 0.1.1
        """
        if self.row1 != self.row2 and self.col1 != self.col2:
            return True
        else:
            return False

    @property
    def shape(self):
        """
        Tuple of Range dimensions.

        .. versionadded:: 0.3.0
        """
        return self.row2 - self.row1 + 1, self.col2 - self.col1 + 1

    @property
    def size(self):
        """
        Number of elements in the Range.

        .. versionadded:: 0.3.0
        """
        return self.shape[0] * self.shape[1]

    def __len__(self):
        return self.row2 - self.row1 + 1

    @property
    def value(self):
        """
        Gets and sets the values for the given Range.

        Returns
        -------
        object
            Empty cells are set to ``None``.
        """
        return conversion.read(self, None, self._options)

    @value.setter
    def value(self, data):
        conversion.write(data, self, self._options)

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
    def formula_array(self):
        """
        Gets or sets an  array formula for the given Range.

        .. versionadded:: 0.7.1
        """
        return xlplatform.get_formula_array(self.xl_range)

    @formula_array.setter
    def formula_array(self, value):
        xlplatform.set_formula_array(self.xl_range, value)

    @property
    def table(self):
        """
        Returns a contiguous Range starting with the indicated cell as top-left corner and going down and right as
        long as no empty cell is hit.

        Keyword Arguments
        -----------------
        strict : boolean, default False
            ``True`` stops the table at empty cells even if they contain a formula. Less efficient than if set to
            ``False``.

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
                     (self.row1, self.col1), wkb=self.workbook, **self._options).vertical.row2
        col2 = Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), wkb=self.workbook, **self._options).horizontal.col2

        return Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row1, self.col1), (row2, col2), wkb=self.workbook, **self._options)

    @property
    def vertical(self):
        """
        Returns a contiguous Range starting with the indicated cell and going down as long as no empty cell is hit.
        This corresponds to ``Ctrl-Shift-DownArrow`` in Excel.

        Arguments
        ---------
        strict : bool, default False
            ``True`` stops the table at empty cells even if they contain a formula. Less efficient than if set to
            ``False``.

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
                     (self.row1, self.col1), (row2, col2), wkb=self.workbook, **self._options)

    @property
    def strict(self):
        return self._options.get('strict', False)

    @property
    def horizontal(self):
        """
        Returns a contiguous Range starting with the indicated cell and going right as long as no empty cell is hit.

        Keyword Arguments
        -----------------
        strict : bool, default False
            ``True`` stops the table at empty cells even if they contain a formula. Less efficient than if set to
            ``False``.

        Returns
        -------
        Range object

        Examples
        --------
        To get the values of a contiguous Range or clear its contents use::

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
                     (self.row1, self.col1), (row2, col2), wbk=self.workbook, **self._options)

    @property
    def current_region(self):
        """
        This property returns a Range object representing a range bounded by (but not including) any
        combination of blank rows and blank columns or the edges of the worksheet. It corresponds to ``Ctrl-*`` on
        Windows and ``Shift-Ctrl-Space`` on Mac.

        Returns
        -------
        Range object

        """
        address = xlplatform.get_current_region_address(self.xl_sheet, self.row1, self.col1)
        return Range(xlplatform.get_worksheet_name(self.xl_sheet), address, wbk=self.workbook, **self._options)

    @property
    def number_format(self):
        """
        Gets and sets the number_format of a Range.

        Examples
        --------

        >>> Range('A1').number_format
        'General'
        >>> Range('A1:C3').number_format = '0.00%'
        >>> Range('A1:C3').number_format
        '0.00%'

        .. versionadded:: 0.2.3
        """
        return xlplatform.get_number_format(self)

    @number_format.setter
    def number_format(self, value):
        xlplatform.set_number_format(self, value)

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

    @property
    def column_width(self):
        """
        Gets or sets the width, in characters, of a Range.
        One unit of column width is equal to the width of one character in the Normal style.
        For proportional fonts, the width of the character 0 (zero) is used.

        If all columns in the Range have the same width, returns the width.
        If columns in the Range have different widths, returns None.

        column_width must be in the range:
        0 <= column_width <= 255

        Note: If the Range is outside the used range of the Worksheet, and columns in the Range have different widths,
        returns the width of the first column.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return xlplatform.get_column_width(self.xl_range)

    @column_width.setter
    def column_width(self, value):
        xlplatform.set_column_width(self.xl_range, value)

    @property
    def row_height(self):
        """
        Gets or sets the height, in points, of a Range.
        If all rows in the Range have the same height, returns the height.
        If rows in the Range have different heights, returns None.

        row_height must be in the range:
        0 <= row_height <= 409.5

        Note: If the Range is outside the used range of the Worksheet, and rows in the Range have different heights,
        returns the height of the first row.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return xlplatform.get_row_height(self.xl_range)

    @row_height.setter
    def row_height(self, value):
        xlplatform.set_row_height(self.xl_range, value)

    @property
    def width(self):
        """
        Returns the width, in points, of a Range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return xlplatform.get_width(self.xl_range)

    @property
    def height(self):
        """
        Returns the height, in points, of a Range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return xlplatform.get_height(self.xl_range)

    @property
    def left(self):
        """
        Returns the distance, in points, from the left edge of column A to the left edge of the range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.6.0
        """
        return xlplatform.get_left(self.xl_range)

    @property
    def top(self):
        """
        Returns the distance, in points, from the top edge of row 1 to the top edge of the range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.6.0
        """
        return xlplatform.get_top(self.xl_range)

    def autofit(self, axis=None):
        """
        Autofits the width of either columns, rows or both.

        Arguments
        ---------
        axis : string or integer, default None
            - To autofit rows, use one of the following: ``rows`` or ``r``
            - To autofit columns, use one of the following: ``columns`` or ``c``
            - To autofit rows and columns, provide no arguments

        Examples
        --------
        ::

            # Autofit column A
            Range('A:A').autofit('c')
            # Autofit row 1
            Range('1:1').autofit('r')
            # Autofit columns and rows, taking into account Range('A1:E4')
            Range('A1:E4').autofit()
            # AutoFit rows, taking into account Range('A1:E4')
            Range('A1:E4').autofit('rows')

        .. versionadded:: 0.2.2
        """
        xlplatform.autofit(self, axis)

    def get_address(self, row_absolute=True, column_absolute=True, include_sheetname=False, external=False):
        """
        Returns the address of the range in the specified format.

        Arguments
        ---------
        row_absolute : bool, default True
            Set to True to return the row part of the reference as an absolute reference.

        column_absolute : bool, default True
            Set to True to return the column part of the reference as an absolute reference.

        include_sheetname : bool, default False
            Set to True to include the Sheet name in the address. Ignored if external=True.

        external : bool, default False
            Set to True to return an external reference with workbook and worksheet name.

        Returns
        -------
        str

        Examples
        --------
        ::

            >>> Range((1,1)).get_address()
            '$A$1'
            >>> Range((1,1)).get_address(False, False)
            'A1'
            >>> Range('Sheet1', (1,1), (3,3)).get_address(True, False, True)
            'Sheet1!A$1:C$3'
            >>> Range('Sheet1', (1,1), (3,3)).get_address(True, False, external=True)
            '[Workbook1]Sheet1!A$1:C$3'

        .. versionadded:: 0.2.3
        """

        if include_sheetname and not external:
            # TODO: when the Workbook name contains spaces but not the Worksheet name, it will still be surrounded
            # by '' when include_sheetname=True. Also, should probably changed to regex
            temp_str = xlplatform.get_address(self.xl_range, row_absolute, column_absolute, True)

            if temp_str.find("[") > -1:
                results_address = temp_str[temp_str.rfind("]") + 1:]
                if results_address.find("'") > -1:
                    results_address = "'" + results_address
                return results_address
            else:
                return temp_str

        else:
            return xlplatform.get_address(self.xl_range, row_absolute, column_absolute, external)

    def __repr__(self):
        return "<Range on Sheet '{0}' of Workbook '{1}'>".format(xlplatform.get_worksheet_name(self.xl_sheet),
                                                                 xlplatform.get_workbook_name(self.xl_workbook))

    @property
    def hyperlink(self):
        """
        Returns the hyperlink address of the specified Range (single Cell only)

        Examples
        --------
        >>> Range('A1').value
        'www.xlwings.org'
        >>> Range('A1').hyperlink
        'http://www.xlwings.org'

        .. versionadded:: 0.3.0
        """
        if self.formula.lower().startswith('='):
            # If it's a formula, extract the URL from the formula string
            formula = self.formula
            try:
                return re.compile(r'\"(.+?)\"').search(formula).group(1)
            except AttributeError:
                raise Exception("The cell doesn't seem to contain a hyperlink!")
        else:
            # If it has been set pragmatically
            return xlplatform.get_hyperlink_address(self.xl_range)

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        """
        Adds a hyperlink to the specified Range (single Cell)

        Arguments
        ---------
        address : str
            The address of the hyperlink.
        text_to_display : str, default None
            The text to be displayed for the hyperlink. Defaults to the hyperlink address.
        screen_tip: str, default None
            The screen tip to be displayed when the mouse pointer is paused over the hyperlink.
            Default is set to '<address> - Click once to follow. Click and hold to select this cell.'


        .. versionadded:: 0.3.0
        """
        if text_to_display is None:
            text_to_display = address
        if address[:4] == 'www.':
            address = 'http://' + address
        if screen_tip is None:
            screen_tip = address + ' - Click once to follow. Click and hold to select this cell.'
        xlplatform.set_hyperlink(self.xl_range, address, text_to_display, screen_tip)

    @property
    def color(self):
        """
        Gets and sets the background color of the specified Range.

        To set the color, either use an RGB tuple ``(0, 0, 0)`` or a color constant.
        To remove the background, set the color to ``None``, see Examples.

        Returns
        -------
        RGB : tuple

        Examples
        --------
        >>> Range('A1').color = (255,255,255)
        >>> from xlwings import RgbColor
        >>> Range('A2').color = RgbColor.rgbAqua
        >>> Range('A2').color
        (0, 255, 255)
        >>> Range('A2').color = None
        >>> Range('A2').color is None
        True

        .. versionadded:: 0.3.0
        """
        return xlplatform.get_color(self.xl_range)

    @color.setter
    def color(self, color_or_rgb):
        xlplatform.set_color(self.xl_range, color_or_rgb)

    def resize(self, row_size=None, column_size=None):
        """
        Resizes the specified Range

        Arguments
        ---------
        row_size: int > 0
            The number of rows in the new range (if None, the number of rows in the range is unchanged).
        column_size: int > 0
            The number of columns in the new range (if None, the number of columns in the range is unchanged).

        Returns
        -------
        Range : Range object


        .. versionadded:: 0.3.0
        """
        if row_size is not None:
            assert row_size > 0
            row2 = self.row1 + row_size - 1
        else:
            row2 = self.row2
        if column_size is not None:
            assert column_size > 0
            col2 = self.col1 + column_size - 1
        else:
            col2 = self.col2

        return Range(xlplatform.get_worksheet_name(self.xl_sheet), (self.row1, self.col1), (row2, col2), wkb=self.workbook, **self._options)

    def offset(self, row_offset=None, column_offset=None):
        """
        Returns a Range object that represents a Range that's offset from the specified range.

        Returns
        -------
        Range : Range object


        .. versionadded:: 0.3.0
        """

        if row_offset:
            row1 = self.row1 + row_offset
            row2 = self.row2 + row_offset
        else:
            row1, row2 = self.row1, self.row2

        if column_offset:
            col1 = self.col1 + column_offset
            col2 = self.col2 + column_offset
        else:
            col1, col2 = self.col1, self.col2

        return Range(xlplatform.get_worksheet_name(self.xl_sheet), (row1, col1), (row2, col2), wkb=self.workbook, **self._options)

    @property
    def column(self):
        """
        Returns the number of the first column in the in the specified range. Read-only.

        Returns
        -------
        Integer


        .. versionadded:: 0.3.5
        """
        return self.col1

    @property
    def row(self):
        """
        Returns the number of the first row in the in the specified range. Read-only.

        Returns
        -------
        Integer


        .. versionadded:: 0.3.5
        """
        return self.row1

    @property
    def last_cell(self):
        """
        Returns the bottom right cell of the specified range. Read-only.

        Returns
        -------
        Range object

        Example
        -------
        >>> rng = Range('A1').table
        >>> rng.last_cell.row, rng.last_cell.column
        (4, 5)

        .. versionadded:: 0.3.5
        """
        return Range(xlplatform.get_worksheet_name(self.xl_sheet),
                     (self.row2, self.col2), **self._options)

    @property
    def name(self):
        """
        Sets or gets the name of a Range.

        To delete a named Range, use ``del wb.names['NamedRange']`` if ``wb`` is
        your Workbook object.

        .. versionadded:: 0.4.0
        """
        return xlplatform.get_named_range(self)

    @name.setter
    def name(self, value):
        xlplatform.set_named_range(self, value)


# This has to be after definition of Range to resolve circular reference
from . import conversion


class Shape(object):
    """
    A Shape object represents an existing Excel shape and can be instantiated with the following arguments::

        Shape(1)            Shape('Sheet1', 1)              Shape(1, 1)
        Shape('Shape 1')    Shape('Sheet1', 'Shape 1')      Shape(1, 'Shape 1')

    The Sheet can also be provided as Sheet object::

        sh = Sheet(1)
        Shape(sh, 'Shape 1')

    If no Worksheet is provided as first argument, it will take the Shape from the active Sheet.

    Arguments
    ---------
    *args
        Definition of Sheet (optional) and shape in the above described combinations.

    Keyword Arguments
    -----------------
    wkb : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.


    .. versionadded:: 0.5.0
    """
    def __init__(self, *args, **kwargs):
        # Use current Workbook if none provided
        self.wkb = kwargs.get('wkb', None)
        self.xl_workbook = Workbook.get_xl_workbook(self.wkb)

        # Arguments
        if len(args) == 1:
            self.sheet_name_or_index = xlplatform.get_worksheet_name(xlplatform.get_active_sheet(self.xl_workbook))
            self.name_or_index = args[0]
        elif len(args) == 2:
            if isinstance(args[0], Sheet):
                self.sheet_name_or_index = args[0].index
            else:
                self.sheet_name_or_index = args[0]
            self.name_or_index = args[1]

        self.xl_shape = xlplatform.get_shape(self)
        self.name = xlplatform.get_shape_name(self)

    @property
    def name(self):
        """
        Returns or sets a String value representing the name of the object.

        .. versionadded:: 0.5.0
        """
        return xlplatform.get_shape_name(self)

    @name.setter
    def name(self, value):
        self.xl_shape = xlplatform.set_shape_name(self.xl_workbook, self.sheet_name_or_index, self.xl_shape, value)

    @property
    def left(self):
        """
        Returns or sets a value that represents the distance, in points, from the left edge of the object to the
        left edge of column A.

        .. versionadded:: 0.5.0
        """
        return xlplatform.get_shape_left(self)

    @left.setter
    def left(self, value):
        xlplatform.set_shape_left(self, value)

    @property
    def top(self):
        """
        Returns or sets a value that represents the distance, in points, from the top edge of the topmost shape
        in the shape range to the top edge of the worksheet.

        .. versionadded:: 0.5.0
        """
        return xlplatform.get_shape_top(self)

    @top.setter
    def top(self, value):
        xlplatform.set_shape_top(self, value)

    @property
    def width(self):
        """
        Returns or sets a value that represents the width, in points, of the object.

        .. versionadded:: 0.5.0
        """
        return xlplatform.get_shape_width(self)

    @width.setter
    def width(self, value):
        xlplatform.set_shape_width(self, value)

    @property
    def height(self):
        """
        Returns or sets a value that represents the height, in points, of the object.

        .. versionadded:: 0.5.0
        """
        return xlplatform.get_shape_height(self)

    @height.setter
    def height(self, value):
        xlplatform.set_shape_height(self, value)

    def delete(self):
        """
        Deletes the object.

        .. versionadded:: 0.5.0
        """
        xlplatform.delete_shape(self)

    def activate(self):
        """
        Activates the object.

        .. versionadded:: 0.5.0
        """
        xlplatform.activate_shape(self.xl_shape)


class Chart(Shape):
    """
    A Chart object represents an existing Excel chart and can be instantiated with the following arguments::

        Chart(1)            Chart('Sheet1', 1)              Chart(1, 1)
        Chart('Chart 1')    Chart('Sheet1', 'Chart 1')      Chart(1, 'Chart 1')

    The Sheet can also be provided as Sheet object::

        sh = Sheet(1)
        Chart(sh, 'Chart 1')

    If no Worksheet is provided as first argument, it will take the Chart from the active Sheet.

    To insert a new Chart into Excel, create it as follows::

        Chart.add()

    Arguments
    ---------
    *args
        Definition of Sheet (optional) and chart in the above described combinations.

    Keyword Arguments
    -----------------
    wkb : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.

    Example
    -------
    >>> from xlwings import Workbook, Range, Chart, ChartType
    >>> wb = Workbook()
    >>> Range('A1').value = [['Foo1', 'Foo2'], [1, 2]]
    >>> chart = Chart.add(source_data=Range('A1').table, chart_type=ChartType.xlLine)
    >>> chart.name
    'Chart1'
    >>> chart.chart_type = ChartType.xl3DArea

    """

    def __init__(self, *args, **kwargs):
        super(Chart, self).__init__(*args, **kwargs)

        # Get xl_chart object
        self.xl_chart = xlplatform.get_chart_object(self.xl_workbook, self.sheet_name_or_index, self.name_or_index)
        self.index = xlplatform.get_chart_index(self.xl_chart)

        # Chart Type
        chart_type = kwargs.get('chart_type')
        if chart_type:
            self.chart_type = chart_type

        # Source Data
        source_data = kwargs.get('source_data')
        if source_data:
            self.set_source_data(source_data)

    @classmethod
    def add(cls, sheet=None, left=0, top=0, width=355, height=211, **kwargs):
        """
        Inserts a new Chart into Excel.

        Arguments
        ---------
        sheet : str or int or xlwings.Sheet, default None
            Name or index of the Sheet or Sheet object, defaults to the active Sheet

        left : float, default 0
            left position in points

        top : float, default 0
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

        wkb : Workbook object, default Workbook.current()
            Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.

        Returns
        -------

        xlwings Chart object
        """
        wkb = kwargs.get('wkb', None)
        xl_workbook = Workbook.get_xl_workbook(wkb)

        chart_type = kwargs.get('chart_type', ChartType.xlColumnClustered)
        name = kwargs.get('name')
        source_data = kwargs.get('source_data')

        if isinstance(sheet, Sheet):
                sheet = sheet.index
        if sheet is None:
            sheet = xlplatform.get_worksheet_index(xlplatform.get_active_sheet(xl_workbook))

        xl_chart = xlplatform.add_chart(xl_workbook, sheet, left, top, width, height)

        if name:
            xlplatform.set_chart_name(xl_chart, name)
        else:
            name = xlplatform.get_chart_name(xl_chart)

        return cls(sheet, name, wkb=wkb, chart_type=chart_type, source_data=source_data)

    @property
    def chart_type(self):
        """
        Gets and sets the chart type of a chart.

        .. versionadded:: 0.1.1
        """
        return xlplatform.get_chart_type(self.xl_chart)

    @chart_type.setter
    def chart_type(self, value):
        xlplatform.set_chart_type(self.xl_chart, value)

    def set_source_data(self, source):
        """
        Sets the source for the chart.

        Arguments
        ---------
        source : Range
            Range object, e.g. ``Range('A1')``
        """
        xlplatform.set_source_data_chart(self.xl_chart, source.xl_range)

    def __repr__(self):
        return "<Chart '{0}' on Sheet '{1}' of Workbook '{2}'>".format(self.name,
                                                                       Sheet(self.sheet_name_or_index).name,
                                                                       xlplatform.get_workbook_name(self.xl_workbook))


class Picture(Shape):
    """
    A Picture object represents an existing Excel Picture and can be instantiated with the following arguments::

        Picture(1)              Picture('Sheet1', 1)                Picture(1, 1)
        Picture('Picture 1')    Picture('Sheet1', 'Picture 1')      Picture(1, 'Picture 1')

    The Sheet can also be provided as Sheet object::

        sh = Sheet(1)
        Shape(sh, 'Picture 1')

    If no Worksheet is provided as first argument, it will take the Picture from the active Sheet.

    Arguments
    ---------
    *args
        Definition of Sheet (optional) and picture in the above described combinations.

    Keyword Arguments
    -----------------
    wkb : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.


    .. versionadded:: 0.5.0
    """
    def __init__(self, *args, **kwargs):
        super(Picture, self).__init__(*args, **kwargs)
        self.xl_picture = xlplatform.get_picture(self)
        self.index = xlplatform.get_picture_index(self)

    @classmethod
    def add(cls, filename, sheet=None, name=None, link_to_file=False, save_with_document=True,
            left=0, top=0, width=None, height=None, wkb=None):
        """
        Inserts a picture into Excel.

        Arguments
        ---------

        filename : str
            The full path to the file.

        Keyword Arguments
        -----------------
        sheet : str or int or xlwings.Sheet, default None
            Name or index of the Sheet or ``xlwings.Sheet`` object, defaults to the active Sheet

        name : str, default None
            Excel picture name. Defaults to Excel standard name if not provided, e.g. 'Picture 1'

        left : float, default 0
            Left position in points.

        top : float, default 0
            Top position in points.

        width : float, default None
            Width in points. If PIL/Pillow is installed, it defaults to the width of the picture.
            Otherwise it defaults to 100 points.

        height : float, default None
            Height in points. If PIL/Pillow is installed, it defaults to the height of the picture.
            Otherwise it defaults to 100 points.

        wkb : Workbook object, default Workbook.current()
            Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.

        Returns
        -------
        xlwings Picture object


        .. versionadded:: 0.5.0
        """

        xl_workbook = Workbook.get_xl_workbook(wkb)

        if isinstance(sheet, Sheet):
                sheet = sheet.index
        if sheet is None:
            sheet = xlplatform.get_worksheet_index(xlplatform.get_active_sheet(xl_workbook))

        if name:
            if name in xlplatform.get_shapes_names(xl_workbook, sheet):
                raise ShapeAlreadyExists('A shape with this name already exists.')

        if sys.platform.startswith('darwin') and xlplatform.get_major_app_version_number(xl_workbook) >= 15:
            # Office 2016 for Mac is sandboxed. This path seems to work without the need of granting access explicitly
            xlwings_picture = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/xlwings_picture.png'
            shutil.copy2(filename, xlwings_picture)
            filename = xlwings_picture

        # Image dimensions
        im_width, im_height = None, None
        if width is None or height is None:
            if Image:
                im = Image.open(filename)
                im_width, im_height = im.size

        if width is None:
            if im_width is not None:
                width = im_width
            else:
                width = 100

        if height is None:
            if im_height is not None:
                height = im_height
            else:
                height = 100

        xl_picture = xlplatform.add_picture(xl_workbook, sheet, filename, link_to_file, save_with_document,
                                            left, top, width, height)

        if sys.platform.startswith('darwin') and xlplatform.get_major_app_version_number(xl_workbook) >= 15:
            os.remove(xlwings_picture)

        if name is None:
            name = xlplatform.get_picture_name(xl_picture)
        else:
            xlplatform.set_shape_name(xl_workbook, sheet, xl_picture, name)

        return cls(sheet, name, wkb=wkb)

    def update(self, filename):
        """
        Replaces an existing picture with a new one, taking over the attributes of the existing picture.

        Arguments
        ---------

        filename : str
            Path to the picture.


        .. versionadded:: 0.5.0
        """
        wkb = self.wkb
        name = self.name
        left, top, width, height = self.left, self.top, self.width, self.height
        sheet_name_or_index = self.sheet_name_or_index
        xlplatform.delete_shape(self)
        # TODO: link_to_file, save_with_document
        Picture.add(filename, sheet=sheet_name_or_index, left=left, top=top, width=width, height=height,
                    name=name, wkb=wkb)


class Plot(object):
    """
    Plot allows to easily display Matplotlib figures as pictures in Excel.

    Arguments
    ---------
    figure : matplotlib.figure.Figure
        Matplotlib figure

    Example
    -------
    Get a matplotlib ``figure`` object:

    * via PyPlot interface::

        import matplotlib.pyplot as plt
        fig = plt.figure()
        plt.plot([1, 2, 3, 4, 5])

    * via object oriented interface::

        from matplotlib.figure import Figure
        fig = Figure(figsize=(8, 6))
        ax = fig.add_subplot(111)
        ax.plot([1, 2, 3, 4, 5])

    * via Pandas::

        import pandas as pd
        import numpy as np

        df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
        ax = df.plot(kind='bar')
        fig = ax.get_figure()

    Then show it in Excel as picture::

        plot = Plot(fig)
        plot.show('Plot1')


    .. versionadded:: 0.5.0
    """
    def __init__(self, figure):
        self.figure = figure

    def show(self, name, sheet=None, left=0, top=0, width=None, height=None, wkb=None):
        """
        Inserts the matplotlib figure as picture into Excel if a picture with that name doesn't exist yet.
        Otherwise it replaces the picture, taking over its position and size.

        Arguments
        ---------

        name : str
            Name of the picture in Excel

        Keyword Arguments
        -----------------
        sheet : str or int or xlwings.Sheet, default None
            Name or index of the Sheet or ``xlwings.Sheet`` object, defaults to the active Sheet

        left : float, default 0
            Left position in points. Only has an effect if the picture doesn't exist yet in Excel.

        top : float, default 0
            Top position in points. Only has an effect if the picture doesn't exist yet in Excel.

        width : float, default None
            Width in points, defaults to the width of the matplotlib figure.
            Only has an effect if the picture doesn't exist yet in Excel.

        height : float, default None
            Height in points, defaults to the height of the matplotlib figure.
            Only has an effect if the picture doesn't exist yet in Excel.

        wkb : Workbook object, default Workbook.current()
            Defaults to the Workbook that was instantiated last or set via ``Workbook.set_current()``.

        Returns
        -------
        xlwings Picture object

        .. versionadded:: 0.5.0
        """
        xl_workbook = Workbook.get_xl_workbook(wkb)

        if isinstance(sheet, Sheet):
                sheet = sheet.index
        if sheet is None:
            sheet = xlplatform.get_worksheet_index(xlplatform.get_active_sheet(xl_workbook))

        if sys.platform.startswith('darwin') and xlplatform.get_major_app_version_number(xl_workbook) >= 15:
            # Office 2016 for Mac is sandboxed. This path seems to work without the need of granting access explicitly
            filename = os.path.expanduser("~") + '/Library/Containers/com.microsoft.Excel/Data/xlwings_plot.png'
        else:
            temp_dir = os.path.realpath(tempfile.gettempdir())
            filename = os.path.join(temp_dir, 'xlwings_plot.png')
        canvas = FigureCanvas(self.figure)
        canvas.draw()
        self.figure.savefig(filename, format='png', bbox_inches='tight')

        if width is None:
            width = self.figure.bbox.bounds[2:][0]

        if height is None:
            height = self.figure.bbox.bounds[2:][1]

        try:
            return Picture.add(sheet=sheet, filename=filename, left=left, top=top, width=width,
                               height=height, name=name, wkb=wkb)
        except ShapeAlreadyExists:
            pic = Picture(sheet, name, wkb=wkb)
            pic.update(filename)
            return pic
        finally:
            os.remove(filename)


class NamesDict(collections.MutableMapping):
    """
    Implements the Workbook.Names collection.
    Currently only used to be able to do ``del wb.names['NamedRange']``
    """

    def __init__(self, xl_workbook, *args, **kwargs):
        self.xl_workbook = xl_workbook
        self.store = dict()
        self.update(dict(*args, **kwargs))

    def __getitem__(self, key):
        return self.store[self.__keytransform__(key)]

    def __setitem__(self, key, value):
        self.store[self.__keytransform__(key)] = value

    def __delitem__(self, key):
        xlplatform.delete_name(self.xl_workbook, key)

    def __iter__(self):
        return iter(self.store)

    def __len__(self):
        return len(self.store)

    def __keytransform__(self, key):
        return key


def view(obj):
    """
    Opens a new workbook and displays an object on its first sheet.

    Parameters
    ----------
    obj : any type with built-in converter
        the object to display

        >>> import xlwings as xw
        >>> import pandas as pd
        >>> import numpy as np
        >>> df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
        >>> xw.view(df)


    .. versionadded:: 0.7.1
    """
    sht = Workbook().active_sheet
    Range(sht, 'A1').value = obj
    sht.autofit()


class Macro(object):
    def __init__(self, name, wb=None, app=None):
        self.name = name
        self.wb = wb
        self.app = app

    def run(self, *args):
        return xlplatform.run(self.wb, self.name, self.app or Application(self.wb), args)

    __call__ = run
