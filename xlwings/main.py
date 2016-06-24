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

from . import xlplatform, string_types, xrange, map, ShapeAlreadyExists, PY3
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


class Apps(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def active(self):
        for app in self.impl:
            return App(impl=app)
        return None

    def __repr__(self):
        return '{}({})'.format(
            self.__class__.__name__,
            repr(list(self))
        )

    def __getitem__(self, item):
        return App(impl=self.impl[item])

    def __len__(self):
        return len(self.impl)

    def __iter__(self):
        for app in self.impl:
            yield App(impl=app)


apps = Apps(impl=xlplatform.Apps())


class App(object):
    """
    Application is dependent on the Workbook since there might be different application instances on Windows.
    """

    def __init__(self, spec=None, impl=None, visible=None):
        if impl is None:
            self.impl = xlplatform.App(spec=spec)
            if visible or visible is None:
                self.visible = True
        else:
            self.impl = impl
            if visible:
                self.visible = True

    @property
    def api(self):
        return self.impl.api

    @property
    def version(self):
        return self.impl.version

    @property
    def active_book(self):
        impl = self.impl.active_book
        return impl and Book(impl=impl)

    @property
    def active_sheet(self):
        return Sheet(impl=self.impl.active_sheet)

    @property
    def selection(self):
        return Range(impl=self.impl.selection)

    def activate(self, steal_focus=False):
        return self.impl.activate(steal_focus)

    @property
    def visible(self):
        return self.impl.visible

    @visible.setter
    def visible(self, value):
        self.impl.visible = value

    def quit(self):
        return self.impl.quit()

    def kill(self):
        return self.impl.kill()

    @property
    def screen_updating(self):
        return self.impl.screen_updating

    @screen_updating.setter
    def screen_updating(self, value):
        self.impl.screen_updating = value

    @property
    def calculation(self):
        return self.impl.calculation

    @calculation.setter
    def calculation(self, value):
        self.impl.calculation = value

    def calculate(self):
        self.impl.calculate()

    @property
    def books(self):
        return Books(impl=self.impl.books)

    @property
    def hwnd(self):
        return self.impl.hwnd

    @property
    def pid(self):
        return self.impl.pid

    def range(self, arg1, arg2=None):
        return Range(impl=self.impl.range(arg1, arg2))

    @property
    def major_version(self):
        return int(self.version.split('.')[0])

    def __repr__(self):
        return "<Excel App %s>" % self.pid

    def __eq__(self, other):
        return type(other) is App and other.pid == self.pid

    def __hash__(self):
        return hash(self.pid)

    def book(self, fullname=None):
        wbs = self.books

        if fullname:
            if not PY3 and isinstance(fullname, str):
                fullname = unicode(fullname, 'mbcs')  # noqa
            fullname = fullname.lower()

            for wb in wbs:
                if wb.fullname.lower() == fullname or wb.name.lower() == fullname:
                    return wb

            if os.path.isfile(fullname):
                return wbs.open(fullname)
            else:
                raise Exception("Could not connect to workbook '%s'" % fullname)

        else:
            # create a new workbook
            return wbs.add()

    def macro(self, macro):
        return Macro(self, macro)


class Book(object):
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

    def __init__(self, fullname=None, impl=None):
        if not impl:
            if fullname:
                if not PY3 and isinstance(fullname, str):
                    fullname = unicode(fullname, 'mbcs')  # noqa
                fullname = fullname.lower()

                candidates = []
                for app in apps:
                    for wb in app.books:
                        if wb.fullname.lower() == fullname or wb.name.lower() == fullname:
                            candidates.append((app, wb))

                if len(candidates) == 0:
                    if os.path.isfile(fullname):
                        impl = active.app.books.open(fullname).impl
                    else:
                        raise Exception("Could not connect to workbook '%s'" % fullname)
                elif len(candidates) > 1:
                    raise Exception("Workbook '%s' is open in more than one Excel instance." % fullname)
                else:
                    impl = candidates[0][1].impl
            else:
                # Open Excel if necessary and create a new workbook
                if active.app:
                    impl = active.app.books.add().impl
                else:
                    app = App()
                    impl = app.books[0].impl

        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    @classmethod
    def active(cls):
        """
        Returns the Workbook that is currently active or has been active last. On Windows,
        this works across all instances.

        .. versionadded:: 0.4.1
        """
        return apps.active.active_book

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
        if hasattr(Book, '_mock_file'):
            # Use mocking Book, see Book.set_mock_caller()
            impl = Book(Book._mock_file).impl
            return cls(impl=impl)
        elif len(sys.argv) > 2 and sys.argv[2] == 'from_xl':
            fullname = sys.argv[1].lower()
            if sys.platform.startswith('win'):
                app = App(impl=xlplatform.App(xl=int(sys.argv[4])))  # hwnd
                return cls(impl=app.book(fullname).impl)
            else:
                return cls(impl=Book(fullname).impl)  # This raises an exception if the same file is open in 2 instances
        else:
            # TODO
            # Called via OPTIMIZED_CONNECTION = True
            return cls(impl=active.book.impl)
        # raise Exception('Workbook.caller() must not be called directly. Call through Excel or set a mock caller '
        #                 'first with Book.set_mock_caller().')

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
        Book._mock_file = fullpath

    @staticmethod
    def open_template():
        """
        Creates a new Excel file with the xlwings VBA module already included. This method must be called from an
        interactive Python shell::

        >>> Book.open_template()

        .. versionadded:: 0.3.3
        """
        this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))
        template_file = 'xlwings_template.xltm'
        try:
            os.remove(os.path.join(this_dir, '~$' + template_file))
        except OSError:
            pass

        xlplatform.open_template(os.path.realpath(os.path.join(this_dir, template_file)))

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
        return self.app.macro("'{0}'!{1}".format(self.name, name))

    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def sheets(self):
        return Sheets(impl=self.impl.sheets)

    @property
    def app(self):
        return App(impl=self.impl.app)

    def close(self):
        self.impl.close()

    @property
    def active_sheet(self):
        return Sheet(impl=self.impl.active_sheet)

    def save(self, path=None):
        return self.impl.save(path)

    @property
    def fullname(self):
        return self.impl.fullname

    @property
    def names(self):
        return Names(impl=self.impl.names)

    def activate(self):
        self.app.activate()
        self.impl.activate()

    @property
    def selection(self):
        return Range(impl=self.impl.active_sheet.selection)

    def sheet(self, name_or_index=None):
        if name_or_index is None:
            return self.sheets.add()
        else:
            return self.sheets(name_or_index)

    def __repr__(self):
        return "<Book [{0}]>".format(self.name)


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

    def __init__(self, sheet=None, impl=None):
        if impl is None:
            self.impl = Book.active().sheet(sheet).impl
        else:
            self.impl = impl

    @property
    def api(self):
        return self.impl.api

    @classmethod
    def active(cls):
        """Returns the active Sheet in the current application. Use like so: ``Sheet.active()``"""
        return apps.active.active_sheet

    @classmethod
    def add(cls, name=None, before=None, after=None):
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
        return active.book.sheets.add(name, before, after)

    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def names(self):
        return Names(impl=self.impl.names)

    @property
    def book(self):
        return Book(impl=self.impl.book)

    @property
    def index(self):
        return self.impl.index

    def range(self, arg1, arg2=None):
        if isinstance(arg1, Range):
            arg1 = arg1.impl
        if isinstance(arg2, Range):
            arg2 = arg2.impl
        return Range(impl=self.impl.range(arg1, arg2))

    @property
    def cells(self):
        return Range(impl=self.impl.cells)

    def activate(self):
        return self.impl.activate()

    def get_value_from_index(self, row_index, column_index):
        return self.impl.get_value_from_index(row_index, column_index)

    def clear_contents(self):
        return self.impl.clear_contents()

    def clear(self):
        return self.impl.clear()

    def autofit(self, axis=None):
        return self.impl.autofit(axis)

    def delete(self):
        return self.impl.delete()

    def add_picture(self, filename, link_to_file, save_with_document, left, top, width, height):
        return Shape(impl=self.impl.add_picture(filename, link_to_file, save_with_document, left, top, width, height))

    def get_shape_object(self, shape_name_or_index):
        return Shape(impl=self.impl.get_shape_object(shape_name_or_index))

    def get_chart_object(self, chart_name_or_index):
        return Chart(impl=self.impl.get_chart_object(chart_name_or_index))

    def get_shapes_names(self):
        shapes = self.xl.Shapes
        if shapes is not None:
            return [i.Name for i in shapes]
        else:
            return []

    def add_chart(self, left, top, width, height):
        return Chart(xl=self.xl.ChartObjects().Add(left, top, width, height))

    def __repr__(self):
        return "<Sheet [{1}]{0}>".format(self.name, self.book.name)

    @property
    def charts(self):
        return Charts(impl=self.impl.charts)

    @property
    def shapes(self):
        return Shapes(impl=self.impl.shapes)

    @property
    def pictures(self):
        return Shapes(impl=self.impl.pictures)


class Range(object):
    """
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

    wkb : Workbook object, default Workbook.current()
        Defaults to the Workbook that was instantiated last or set via `Workbook.set_current()``.
    """

    def __init__(self, *args, impl=None, **options):

        # Arguments
        if impl is None:
            if len(args) == 2 and isinstance(args[0], Range) and isinstance(args[1], Range):
                #if args[0].sheet.impl != args[1].sheet.impl:
                #    raise ValueError("Ranges are not on the same sheet")
                impl = args[0].sheet.range(args[0], args[1]).impl
            elif len(args) == 1 and isinstance(args[0], string_types):
                impl = active.app.range(args[0]).impl
            elif 0 < len(args) <= 3:
                if isinstance(args[-1], tuple):
                    if len(args) > 1 and isinstance(args[-2], tuple):
                        spec = (args[-2], args[-1])
                    else:
                        spec = (args[-1],)
                    if any(0 in s for s in spec):
                        raise IndexError("Attempted to access 0-based Range. xlwings/Excel Ranges are 1-based.")
                elif isinstance(args[-1], string_types):
                    spec = (args[-1],)
                residual_args = args[:-len(spec)]
                if residual_args:
                    if len(residual_args) > 1:
                        raise ValueError("Invalid arguments")
                    else:
                        sheet = residual_args[0]
                        if not isinstance(sheet, Sheet):
                            sheet = Sheet(sheet)
                else:
                    sheet = Sheet.active()
                impl = sheet.range(*spec).impl
            else:
                raise ValueError("Invalid arguments")

        self.impl = impl

        # Keyword Arguments
        self._options = options

    @property
    def api(self):
        return self.impl.api

    def __iter__(self):
        # Iterator object that returns cell Ranges: (1, 1), (1, 2) etc.
        for i in range(len(self)):
            yield self(i+1)

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
            impl=self.impl,
            **options
        )

    @property
    def sheet(self):
        return Sheet(impl=self.impl.sheet)

    def __len__(self):
        return len(self.impl)

    @property
    def row(self):
        return self.impl.row

    @property
    def column(self):
        return self.impl.column

    @property
    def row_count(self):
        return self.impl.row_count

    @property
    def column_count(self):
        return self.impl.column_count

    @property
    def raw_value(self):
        return self.impl.raw_value

    @raw_value.setter
    def raw_value(self, data):
        self.impl.raw_value = data

    def clear_contents(self):
        return self.impl.clear_contents()

    def get_cell(self, row, col):
        return Range(impl=self.impl.get_cell(row, col))

    def clear(self):
        return self.impl.clear()

    def end(self, direction):
        return Range(impl=self.impl.end(direction))

    @property
    def formula(self):
        return self.impl.formula

    @formula.setter
    def formula(self, value):
        self.impl.formula = value

    @property
    def formula_array(self):
        return self.impl.formula_array

    @formula_array.setter
    def formula_array(self, value):
        self.impl.formula_array = value

    @property
    def column_width(self):
        return self.impl.column_width

    @column_width.setter
    def column_width(self, value):
        self.impl.column_width = value

    @property
    def row_height(self):
        return self.impl.row_height

    @row_height.setter
    def row_height(self, value):
        self.impl.row_height = value

    @property
    def width(self):
        return self.impl.width

    @property
    def height(self):
        return self.impl.height

    @property
    def left(self):
        return self.impl.left

    @property
    def top(self):
        return self.impl.top

    @property
    def number_format(self):
        return self.impl.number_format

    @number_format.setter
    def number_format(self, value):
        self.impl.number_format = value

    @property
    def address(self):
        return self.impl.address

    @property
    def current_region(self):
        return Range(impl=self.impl.current_region)

    def autofit(self, axis=None):
        return self.impl.autofit(axis)

    def set_hyperlink(self, address, text_to_display=None, screen_tip=None):
        return self.impl.set_hyperlink(address, text_to_display, screen_tip)

    @property
    def color(self):
        return self.impl.color

    @color.setter
    def color(self, color_or_rgb):
        self.impl.color = color_or_rgb

    @property
    def name(self):
        impl = self.impl.name
        return impl and Name(impl=impl)

    @name.setter
    def name(self, value):
        self.impl.name = value

    def __call__(self, *args):
        return Range(impl=self.impl(*args))

    @property
    def rows(self):
        return Range(impl=self.impl.rows)

    @property
    def columns(self):
        return Range(impl=self.impl.columns)

    @property
    def shape(self):
        """
        Tuple of Range dimensions.

        .. versionadded:: 0.3.0
        """
        return self.row_count, self.column_count

    @property
    def size(self):
        """
        Number of elements in the Range.

        .. versionadded:: 0.3.0
        """
        a, b = self.shape
        return a * b

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
        origin = self(1, 1)
        if origin(2, 1).raw_value in [None, ""]:
            bottom_left = origin
        elif origin(3, 1).raw_value in [None, ""]:
            bottom_left = origin(2, 1)
        else:
            bottom_left = origin(2, 1).end('down')

        if origin(1, 2).raw_value in [None, ""]:
            top_right = origin
        elif origin(1, 3).raw_value in [None, ""]:
            top_right = origin(1, 2)
        else:
            top_right = origin(1, 2).end('right')

        return Range(top_right, bottom_left)

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
        if self(2, 1).raw_value in [None, ""]:
            return Range(self(1, 1), self(1, self.column_count))
        elif self(3, 1).raw_value in [None, ""]:
            return Range(self(1, 1), self(2, self.column_count))
        else:
            end_row = self(2, 1).end('down').row - self.row + 1
            return Range(self(1, 1), self(end_row, self.column_count))

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
        if self(1, 2).raw_value in [None, ""]:
            return Range(self(1, 1), self(self.row_count, 1))
        elif self(1, 3).raw_value in [None, ""]:
            return Range(self(1, 1), self(self.row_count, 2))
        else:
            end_column = self(1, 2).end('right').column - self.column + 1
            return Range(self(1, 1), self(self.row_count, end_column))

    def __getitem__(self, key):
        if type(key) is tuple:
            row, col = key
            if isinstance(row, slice):
                if row.step is not None:
                    raise ValueError("Slice steps not supported.")
                row1 = 0 if row.start is None else row.start
                row2 = self.row_count - 1 if row.stop is None else row.stop - 1
            else:
                row1 = row2 = row
            if isinstance(col, slice):
                if col.step is not None:
                    raise ValueError("Slice steps not supported.")
                col1 = 0 if col.start is None else col.start
                col2 = self.column_count - 1 if col.stop is None else col.stop - 1
            else:
                col1 = col2 = col
            if col1 == col2 and row1 == row2:
                return self(row1 + 1, col1 + 1)
            else:
                return self.sheet.range(
                    self(row1 + 1, col1 + 1),
                    self(row2 + 1, col2 + 1)
                )
        elif isinstance(key, slice):
            if key.step is not None:
                raise ValueError("Slice steps not supported.")
            l = len(self)
            start = key.start
            if start is None:
                start = 0
            elif start >= l:
                raise IndexError("Start index %s out of range (%s elements)." % (start, l))
            elif start < 0:
                if start < -l:
                    raise IndexError("Start index %s out of range (%s elements)." % (start, l))
                else:
                    start = l + start
            stop = key.stop
            if stop is None:
                stop = l
            elif stop > l:
                raise IndexError("Stop index %s out of range (%s elements)." % (stop, l))
            elif stop < 0:
                if stop <= -l:
                    raise IndexError("Stop index %s out of range (%s elements)." % (stop, l))
                else:
                    stop = l + stop
            return self._cls.Range(self(start + 1), self(stop))
        else:
            l = len(self)
            if key >= l:
                raise IndexError("Index %s out of range (%s elements)." % (key, l))
            elif key < 0:
                if key < -l:
                    raise IndexError("Index %s out of range (%s elements)." % (key, l))
                else:
                    return self(l + key + 1)
            else:
                return self(key + 1)

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
            temp_str = self.impl.get_address(row_absolute, column_absolute, True)

            if temp_str.find("[") > -1:
                results_address = temp_str[temp_str.rfind("]") + 1:]
                if results_address.find("'") > -1:
                    results_address = "'" + results_address
                return results_address
            else:
                return temp_str

        else:
            return self.impl.get_address(row_absolute, column_absolute, external)

    def __repr__(self):
        return "<Range [{1}]{0}!{2}>".format(
            self.sheet.name,
            self.sheet.book.name,
            self.address
        )

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
            return self.impl.get_hyperlink_address()

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
        self.impl.set_hyperlink(address, text_to_display, screen_tip)

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
        else:
            row_size = self.row_count
        if column_size is not None:
            assert column_size > 0
        else:
            column_size = self.column_count

        return Range(self(1, 1), self(row_size, column_size)).options(**self._options)

    def offset(self, row_offset=0, column_offset=0):
        """
        Returns a Range object that represents a Range that's offset from the specified range.

        Returns
        -------
        Range : Range object


        .. versionadded:: 0.3.0
        """
        return Range(
            self(
                row_offset + 1,
                column_offset + 1
            ),
            self(
                row_offset + self.row_count,
                column_offset + self.column_count
            )
        ).options(**self._options)

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
        return self(self.row_count, self.column_count).options(**self._options)


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
    def __init__(self, *args, impl=None, **kwargs):

        if impl is None:
            if len(args) == 1:
                impl = Sheet.active().get_shape_object(args[0])

            elif len(args) == 2:
                sheet = args[0]
                if not isinstance(sheet, Sheet):
                    sheet = Sheet(sheet)
                impl = sheet.get_shape_object(args[1])

            else:
                raise ValueError("Invalid arguments")

        self.impl = impl

    @property
    def name(self):
        return self.impl.name

    @property
    def contents(self):
        impl = self.impl.contents
        if isinstance(impl, xlplatform.Chart):
            return Chart(impl=impl)
        elif isinstance(impl, xlplatform.Picture):
            return Picture(impl=impl)
        else:
            raise Exception("Unsupported shape content type")

    @property
    def parent(self):
        return Sheet(impl=self.impl.parent)

    def __repr__(self):
        return "<Shape '{0}' in {1}>".format(
            self.name,
            self.parent
        )


class Collection(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    def __call__(self, name_or_index):
        return self._wrap(impl=self.impl(name_or_index))

    def __len__(self):
        return len(self.impl)

    def __iter__(self):
        for impl in self.impl:
            yield self._wrap(impl=impl)

    def __getitem__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            l = len(self)
            if name_or_index >= l:
                raise IndexError("Workbook index %s out of range (%s workbooks)" % (name_or_index, l))
            if name_or_index < 0:
                if name_or_index < -l:
                    raise IndexError("Workbook index %s out of range (%s workbooks)" % (name_or_index, l))
                name_or_index += l
            return self(name_or_index + 1)
        else:
            return self(name_or_index)

    def __repr__(self):
        r = []
        for i, x in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(x))

        return '{}({})'.format(
            self.__class__.__name__,
            "[" + ", ".join(r) + "]"
        )


class Shapes(Collection):
    _wrap = Shape

    def add_picture(self, filename, link_to_file, save_with_document, left, top, width, height):
        return Shape(impl=self.impl.add_picture(filename, link_to_file, save_with_document, left, top, width, height))


class Chart(object):
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
    >>> from xlwings import Book, Range, Chart, ChartType
    >>> wb = Book()
    >>> Range('A1').value = [['Foo1', 'Foo2'], [1, 2]]
    >>> chart = Chart.add(source_data=Range('A1').table, chart_type=ChartType.xlLine)
    >>> chart.name
    'Chart1'
    >>> chart.chart_type = ChartType.xl3DArea

    """

    def __init__(self, *args, impl=None, **kwargs):
        if impl is not None:
            self.impl = impl
        elif len(args) == 1:
            # Get xl_chart object
            self.impl = Sheet.active().charts(args[0]).impl
        elif len(args) == 2:
            sheet = args[0]
            if not isinstance(sheet, Sheet):
                sheet = Sheet(sheet)
            self.impl = sheet.charts(args[1]).impl

        # Chart Type
        chart_type = kwargs.get('chart_type')
        if chart_type:
            self.chart_type = chart_type

        # Source Data
        source_data = kwargs.get('source_data')
        if source_data:
            self.set_source_data(source_data)

    @property
    def api(self):
        return self.impl.api

    @property
    def name(self):
        return self.impl.name

    @property
    def container(self):
        impl = self.impl.container
        if isinstance(impl, xlplatform.Shape):
            return Shape(impl=impl)
        else:
            raise Exception("Container type not supported")

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

        if sheet is None:
            sheet = active.sheet
        elif not isinstance(sheet, Sheet):
            sheet = active.book.sheet(sheet)

        xl_chart = sheet.xl_sheet.add_chart(left, top, width, height)

        chart_type = kwargs.get('chart_type', ChartType.xlColumnClustered)
        name = kwargs.get('name')
        source_data = kwargs.get('source_data')

        if name:
            xl_chart.set_name(name)

        return cls(xl_chart=xl_chart, chart_type=chart_type, source_data=source_data)

    @property
    def chart_type(self):
        """
        Gets and sets the chart type of a chart.

        .. versionadded:: 0.1.1
        """
        return self.impl.chart_type

    @chart_type.setter
    def chart_type(self, value):
        self.impl.chart_type = value

    def set_source_data(self, source):
        """
        Sets the source for the chart.

        Arguments
        ---------
        source : Range
            Range object, e.g. ``Range('A1')``
        """
        self.impl.set_source_data(source.impl)

    def __repr__(self):
        return "<Chart '{0}' in {1}>".format(
            self.name,
            self.container
        )


class Charts(Collection):
    _wrap = Chart


class Picture(object):
    """
    A Picture object represents an existing Excel Picture and can be instantiated with the following arguments::

        Picture(1)              Picture('Sheet1', 1)                Picture(1, 1)
        Picture('Picture 1')    Picture('Sheet1', 'Picture 1')      Picture(1, 'Picture 1')

    The Sheet can also be provided as Sheet object::

        sh = Sheet(1)
        Picture(sh, 'Picture 1')

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
    def __init__(self, *args, impl=None, **kwargs):
        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    @property
    def name(self):
        return self.impl.name

    @property
    def container(self):
        return Shape(impl=self.impl.container)

    @classmethod
    def add(cls, filename, sheet=None, name=None, link_to_file=False, save_with_document=True,
            left=0, top=0, width=None, height=None):
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

        if sheet is None:
            sheet = active.sheet
        elif not isinstance(sheet, Sheet):
            sheet = active.book.sheet(sheet)

        if name:
            if name in sheet.xl_sheet.get_shapes_names():
                raise ShapeAlreadyExists('A shape with this name already exists.')

        if sys.platform.startswith('darwin') and sheet.book.app.major_version >= 15:
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

        if sys.platform.startswith('darwin') and sheet.book.app.major_version >= 15:
            os.remove(xlwings_picture)

        xl_shape = sheet.xl_sheet.add_picture(
            filename, link_to_file, save_with_document,
            left, top, width, height
        )

        if name:
            xl_shape.set_name(name)

        return cls(xl_shape=xl_shape)

    def update(self, filename):
        """
        Replaces an existing picture with a new one, taking over the attributes of the existing picture.

        Arguments
        ---------

        filename : str
            Path to the picture.


        .. versionadded:: 0.5.0
        """
        left, top, width, height = self.left, self.top, self.width, self.height
        name = self.name
        self.xl_shape.delete()
        # TODO: link_to_file, save_with_document
        self.xl_shape = Picture.add(filename, left=left, top=top, width=width, height=height, name=name).xl_shape


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
        xl_workbook = Book.get_xl_workbook(wkb)

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


class Names(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    def __call__(self, name_or_index):
        return Name(impl=self.impl(name_or_index))

    def contains(self, name_or_index):
        return self.impl.contains(name_or_index)

    def __len__(self):
        return len(self.impl)

    def add(self, name, refers_to):
        return Name(impl=self.impl.add(name, refers_to))

    def __getitem__(self, item):
        if isinstance(item, numbers.Number):
            return self(item + 1)
        else:
            return self(item)

    def __setitem__(self, key, value):
        if isinstance(value, Range):
            value.name = key
        elif key in self:
            self[key].refers_to = value
        else:
            self.add(key, value)

    def __contains__(self, item):
        if isinstance(item, numbers.Number):
            return 0 <= item < len(self)
        else:
            return self.contains(item)

    def __delitem__(self, key):
        if key in self:
            self[key].delete()
        else:
            raise KeyError(key)

    def __iter__(self):
        for i in range(len(self)):
            yield self(i+1)

    def __repr__(self):
        r = []
        for i, n in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(n))
        return "[" + ", ".join(r) + "]"


class Name(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    def delete(self):
        self.impl.delete()

    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def refers_to(self):
        return self.impl.refers_to

    @refers_to.setter
    def refers_to(self, value):
        self.impl.refers_to = value

    @property
    def refers_to_range(self):
        return Range(impl=self.impl.refers_to_range)

    def __repr__(self):
        return "<Name '%s': %s>" % (self.name, self.refers_to)
    

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
    sht = Book().active_sheet
    Range(sht, 'A1').value = obj
    sht.autofit()


class Macro(object):
    def __init__(self, app, macro):
        self.app = app
        self.macro = macro

    def run(self, *args):
        return self.app.impl.run(self.macro, args)

    __call__ = run


class Books(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    def __call__(self, name_or_index):
        return Book(impl=self.impl(name_or_index))

    def __len__(self):
        return len(self.impl)

    def add(self):
        return Book(impl=self.impl.add())

    def open(self, fullname):
        return Book(impl=self.impl.open(fullname))

    def __iter__(self):
        for impl in self.impl:
            yield Book(impl=impl)

    def __getitem__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            l = len(self)
            if name_or_index >= l:
                raise IndexError("Workbook index %s out of range (%s workbooks)" % (name_or_index, l))
            if name_or_index < 0:
                if name_or_index < -l:
                    raise IndexError("Workbook index %s out of range (%s workbooks)" % (name_or_index, l))
                name_or_index += l
            return self(name_or_index + 1)
        else:
            return self(name_or_index)

    def __repr__(self):
        r = []
        for i, wb in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(wb))
        return "["+", ".join(r)+"]"


class Sheets(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        return self.impl.api

    def __call__(self, name_or_index):
        if isinstance(name_or_index, Sheet):
            return name_or_index
        else:
            return Sheet(impl=self.impl(name_or_index))

    def __len__(self):
        return len(self.impl)

    def __repr__(self):
        r = []
        for i, sht in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(sht))
        return "["+", ".join(r)+"]"

    def __getitem__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            l = len(self)
            if name_or_index >= l:
                raise IndexError("Sheet index %s out of range (%s sheets)" % (name_or_index, l))
            if name_or_index < 0:
                if name_or_index < -l:
                    raise IndexError("Sheet index %s out of range (%s sheets)" % (name_or_index, l))
                name_or_index += l
            return self(name_or_index + 1)
        else:
            return self(name_or_index)

    def __delitem__(self, name_or_index):
        self[name_or_index].delete()

    def __iter__(self):
        for i in range(len(self)):
            yield self(i+1)

    def add(self, name=None, before=None, after=None):
        if name is not None:
            if name.lower() in (s.name.lower() for s in self):
                raise ValueError("Sheet named '%s' already present in workbook" % name)
        if before is not None and not isinstance(before, Sheet):
            before = self(before)
        if after is not None and not isinstance(after, Sheet):
            after = self(after)
        s = self.impl.add(before and before.impl, after and after.impl)
        if name is not None:
            s.name = name
        return s


class ActiveObjects(object):

    @property
    def app(self):
        return apps.active

    @property
    def book(self):
        return Book.active()

    @property
    def sheet(self):
        return Sheet.active()


active = ActiveObjects()