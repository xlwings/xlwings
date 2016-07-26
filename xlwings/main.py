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
import inspect

from . import xlplatform, string_types, ShapeAlreadyExists, PY3
from .utils import VersionNumber
from . import utils

# Optional imports
try:
    import matplotlib as mpl
    from matplotlib.backends.backend_agg import FigureCanvas
except ImportError:
    mpl = None
    FigureCanvas = None

try:
    from PIL import Image
except ImportError:
    Image = None


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

    count = property(__len__)

    def __iter__(self):
        for impl in self.impl:
            yield self._wrap(impl=impl)

    def __getitem__(self, key):
        if isinstance(key, numbers.Number):
            l = len(self)
            if key >= l:
                raise IndexError("Index %s out of range (%s elements)" % (key, l))
            if key < 0:
                if key < -l:
                    raise IndexError("Index %s out of range (%s elements)" % (key, l))
                key += l
            return self(key + 1)
        elif isinstance(key, slice):
            raise ValueError(self.impl.__class__.__name__ + " object does not support slicing")
        else:
            return self(key)

    def __contains__(self, key):
        return key in self.impl

    # used by repr - by default the name of the collection class, but can be overridden
    @property
    def _name(self):
        return self.__class__.__name__

    def __repr__(self):
        r = []
        for i, x in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(x))

        return '{}({})'.format(
            self._name,
            "[" + ", ".join(r) + "]"
        )


class Apps(object):
    """
    Use ``xw.apps`` to get access to the :class:`App` objects that are contained in the apps collection.

    Examples
    --------

    >>> xw.apps
    Apps([<Excel App 1668>, <Excel App 1644>])
    >>> xw.apps[0]
    <Excel App 1668>
    >>> xw.apps.active
    <Excel App 1668>

    """

    def __init__(self, impl):
        self.impl = impl

    @property
    def active(self):
        """
        Returns the active app.
        """
        for app in self.impl:
            return App(impl=app)
        return None

    def __call__(self, i):
        return self[i-1]

    def __repr__(self):
        return '{}({})'.format(
            self.__class__.__name__,
            repr(list(self))
        )

    def __getitem__(self, item):
        return App(impl=self.impl[item])

    def __len__(self):
        return len(self.impl)

    @property
    def count(self):
        """
        Returns the number of apps.

        """
        return len(self)

    def __iter__(self):
        for app in self.impl:
            yield App(impl=app)


apps = Apps(impl=xlplatform.Apps())


class App(object):
    """
    Excel application instance.

    Parameters
    ----------

    spec : str, default None
        Mac-only, use the full path to the Excel application,
        e.g. ``/Applications/Microsoft Office 2011/Microsoft Excel`` or ``/Applications/Microsoft Excel``

        On Windows, if you want to change the version of Excel that xlwings talks to, go to ``Control Panel >
        Programs and Features`` and ``Repair`` the Office version that you want as default.

    visible : bool, default None
        Returns or sets a boolean value that determines whether the app is visible.

    Examples
    --------

    >>> app1 = xw.App()
    >>> app2 = xw.App()
    >>> xw.apps
    Apps([<Excel App 1668>, <Excel App 1644>])

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
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    @property
    def version(self):
        """
        Returns the Excel version number object.

        Examples
        --------

        >>> app.version
        VersionNumber('15.24')
        >>> app.version.major
        15


        .. versionchanged:: 0.9.0
        """
        return VersionNumber(self.impl.version)

    @property
    def selection(self):
        """
        Returns the selected cells as Range.


        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.selection)

    def activate(self, steal_focus=False):
        """
        Activates the Excel app.

        Parameters
        ----------
        steal_focus : bool, default False
            If True, make frontmost application and hand over focus from Python to Excel.


        .. versionadded:: 0.9.0
        """
        # Win Excel >= 2013 fails if visible=False...we may somehow not be using the correct HWND
        self.impl.activate(steal_focus)
        if self != apps.active:
            raise Exception('Could not activate App! Try to instantiate the App with visible=True.')

    @property
    def visible(self):
        """
        Gets or sets the visibility of Excel to ``True`` or  ``False``.


        .. versionadded:: 0.3.3
        """
        return self.impl.visible

    @visible.setter
    def visible(self, value):
        self.impl.visible = value

    def quit(self):
        """
        Quits the application without saving any workbooks.


        .. versionadded:: 0.3.3

        """
        return self.impl.quit()

    def kill(self):
        return self.impl.kill()

    @property
    def screen_updating(self):
        """
        Gets or sets screen updating to ``True`` or ``False``.


        .. versionadded:: 0.3.3
        """
        return self.impl.screen_updating

    @screen_updating.setter
    def screen_updating(self, value):
        self.impl.screen_updating = value

    @property
    def calculation(self):
        """
        Returns or sets a calculation value that represents the calculation mode.
        Modes: ``'manual'``, ``'automatic'``, ``'semiautomatic'``

        Examples
        --------
        >>> wb = xw.Book()
        >>> wb.app.calculation = 'manual'


        .. versionchanged:: 0.9.0
        """
        return self.impl.calculation

    @calculation.setter
    def calculation(self, value):
        self.impl.calculation = value

    def calculate(self):
        """
        Calculates all open books.


        .. versionadded:: 0.3.6

        """
        self.impl.calculate()

    @property
    def books(self):
        """
        A collection of all Book objects that are currently open.


        .. versionadded:: 0.9.0
        """
        return Books(impl=self.impl.books)

    @property
    def hwnd(self):
        """
        Returns the Window handle (Windows-only).


        .. versionadded:: 0.9.0
        """
        return self.impl.hwnd

    @property
    def pid(self):
        """
        Returns the PID of the app.


        .. versionadded:: 0.9.0
        """
        return self.impl.pid

    def range(self, arg1, arg2=None):
        """
        Returns a Range object from the active sheet of the active book.

        See Also
        --------
        Range


        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.range(arg1, arg2))

    def __repr__(self):
        return "<Excel App %s>" % self.pid

    def __eq__(self, other):
        return type(other) is App and other.pid == self.pid

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.pid)

    def macro(self, name):
        """
        Runs a Sub or Function in Excel VBA that are not part of a specific workbook but e.g. are part of an add-in.

        Arguments
        ---------
        name : Name of Sub or Function with or without module name, e.g. ``'Module1.MyMacro'`` or ``'MyMacro'``

        Examples
        --------
        This VBA function:

        .. code-block:: vb.net

            Function MySum(x, y)
                MySum = x + y
            End Function

        can be accessed like this:

        >>> app = xw.apps[0]
        >>> my_sum = app.macro('MySum')
        >>> my_sum(1, 2)
        3

        See Also
        --------

        Book.macro


        .. versionadded:: 0.9.0
        """
        return Macro(self, name)


class Book(object):
    """
    Represents an Excel Workbook in Python. The following syntax is an alternative for using
    the ``books`` collection to instantiate Book objects in the active app:

    +--------------------+--------------------------------------+--------------------------------------------+
    |                    | xw.Book                              | xw.books                                   |
    +====================+======================================+============================================+
    | New book           | ``xw.Book()``                        | ``xw.books.add()``                         |
    +--------------------+--------------------------------------+--------------------------------------------+
    | Unsaved book       | ``xw.Book('Book1')``                 | ``xw.books['Book1']``                      |
    +--------------------+--------------------------------------+--------------------------------------------+
    | Book by (full)name | ``xw.Book(r'C:/path/to/file.xlsx')`` | ``xw.books.open(r'C:/path/to/file.xlsx')`` |
    +--------------------+--------------------------------------+--------------------------------------------+

    Parameters
    ----------
    fullname : str, default None
        Full path or name (incl. xlsx, xlsm etc.) of existing workbook or name of an unsaved workbook. Without a full
        path, it looks for the file in the current working directory.

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
                        if not apps.active:
                            app = App()
                        impl = apps.active.books.open(fullname).impl
                    else:
                        raise Exception("Could not connect to workbook '%s'" % fullname)
                elif len(candidates) > 1:
                    raise Exception("Workbook '%s' is open in more than one Excel instance." % fullname)
                else:
                    impl = candidates[0][1].impl
            else:
                # Open Excel if necessary and create a new workbook
                if apps.active:
                    impl = apps.active.books.add().impl
                else:
                    app = App()
                    impl = app.books[0].impl

        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __eq__(self, other):
        return isinstance(other, Book) and self.app == other.app and self.name == other.name

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash((self.app, self.name))

    @classmethod
    def caller(cls):
        """
        References the calling book when the Python function is called from Excel via ``RunPython``.
        Pack it into the function being called from Excel, e.g.:

        .. code-block:: python

             def my_macro():
                wb = xw.Book.caller()
                wb.sheets[0].range('A1').value = 1

        To be able to easily invoke such code from Python for debugging, use ``xw.Book.set_mock_caller()``.

        .. versionadded:: 0.3.0
        """
        if hasattr(Book, '_mock_caller'):
            # Use mocking Book, see Book.set_mock_caller()
            return cls(impl=Book._mock_caller.impl)
        elif len(sys.argv) > 2 and sys.argv[2] == 'from_xl':
            fullname = sys.argv[1].lower()
            if sys.platform.startswith('win'):
                app = App(impl=xlplatform.App(xl=int(sys.argv[4])))  # hwnd
                return cls(impl=app.books.open(fullname).impl)
            else:
                # On Mac, the same file open in two instances is not supported
                return cls(impl=Book(fullname).impl)
        elif xlplatform.BOOK_CALLER:
            # Called via OPTIMIZED_CONNECTION = True
            return cls(impl=xlplatform.Book(xlplatform.BOOK_CALLER))
        else:
            raise Exception('Workbook.caller() must not be called directly. Call through Excel or set a mock caller '
                            'first with Book.set_mock_caller().')

    def set_mock_caller(self):
        """
        Sets the Excel file which is used to mock ``xw.Book.caller()`` when the code is called from Python and not from
        Excel via ``RunPython``.

        Examples
        --------
        ::

            # This code runs unchanged from Excel via RunPython and from Python directly
            import os
            import xlwings as xw

            def my_macro():
                wb = xw.Book.caller()
                wb.sheets[0].range('A1').value = 'Hello xlwings!'

            if __name__ == '__main__':
                xw.Book(r'C:\\path\\to\\file.xlsx').set_mock_caller()
                my_macro()

        .. versionadded:: 0.3.1
        """
        Book._mock_caller = self

    @staticmethod
    def open_template():
        """
        Creates a new Excel file with the xlwings VBA module already included. This method must be called from an
        interactive Python shell::

        >>> Book.open_template()

        See Also
        --------

        :ref:`command_line`


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

        Arguments
        ---------
        name : Name of Sub or Function with or without module name, e.g. ``'Module1.MyMacro'`` or ``'MyMacro'``

        Examples
        --------
        This VBA function:

        .. code-block:: vb.net

            Function MySum(x, y)
                MySum = x + y
            End Function

        can be accessed like this:

        >>> wb = xw.books.active
        >>> my_sum = wb.macro('MySum')
        >>> my_sum(1, 2)
        3

        See Also
        --------

        App.macro


        .. versionadded:: 0.7.1
        """
        return self.app.macro("'{0}'!{1}".format(self.name, name))

    @property
    def name(self):
        """
        Returns the name of the book as str.
        """
        return self.impl.name

    @property
    def sheets(self):
        """
        Returns a sheets collection that represents all the sheets in the book.


        .. versionadded:: 0.9.0
        """
        return Sheets(impl=self.impl.sheets)

    @property
    def app(self):
        """
        Returns an app object that represents the creator of the book.


        .. versionadded:: 0.9.0
        """
        return App(impl=self.impl.app)

    def close(self):
        """
        Closes the book without saving it.


        .. versionadded:: 0.1.1
        """
        self.impl.close()

    def save(self, path=None):
        return self.impl.save(path)

    @property
    def fullname(self):
        return self.impl.fullname

    @property
    def names(self):
        """
        Returns a names collection that represents all the names in the specified book (including all sheet-specific names).

        .. versionchanged:: 0.9.0

        """
        return Names(impl=self.impl.names)

    def activate(self, steal_focus=False):
        """
        Activates the book.

        Parameters
        ----------
        steal_focus : bool, default False
            If True, make frontmost window and hand over focus from Python to Excel.

        """
        self.app.activate(steal_focus=steal_focus)
        self.impl.activate()

    @property
    def selection(self):
        """
        Returns the selected cells as Range.


        .. versionadded:: 0.9.0
        """
        return Range(impl=self.app.selection.impl)

    def __repr__(self):
        return "<Book [{0}]>".format(self.name)


class Sheet(object):
    def __init__(self, sheet=None, impl=None):
        if impl is None:
            self.impl = books.active.sheets(sheet).impl
        else:
            self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __eq__(self, other):
        return isinstance(other, Sheet) and self.book == other.book and self.name == other.name

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash((self.book, self.name))

    @property
    def name(self):
        """Gets or sets the name of the Sheet."""
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def names(self):
        """
        Returns a names collection that represents all the sheet-specific names
        (names defined with the "SheetName!" prefix).

        .. versionadded:: 0.9.0

        """
        return Names(impl=self.impl.names)

    @property
    def book(self):
        """Returns the Book of the specified Sheet. Read-only."""
        return Book(impl=self.impl.book)

    @property
    def index(self):
        """Returns the index of the Sheet (1-based as in Excel)."""
        return self.impl.index

    def range(self, arg1, arg2=None):
        """
        Gets or sets a Range object from the active sheet of the active book.

        See Also
        --------
        Range


        .. versionadded:: 0.9.0
        """
        if isinstance(arg1, Range):
            if arg1.sheet != self:
                raise ValueError("First range is not on this sheet")
            arg1 = arg1.impl
        if isinstance(arg2, Range):
            if arg2.sheet != self:
                raise ValueError("Second range is not on this sheet")
            arg2 = arg2.impl
        return Range(impl=self.impl.range(arg1, arg2))

    @property
    def cells(self):
        """
        Returns a Range object that represents all the cells on the Sheet
        (not just the cells that are currently in use).


        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.cells)

    def activate(self):
        """Activates the Sheet and returns it."""
        self.book.activate()
        return self.impl.activate()

    def select(self):
        """
        Selects the Sheet. Select only works on the active book.


        .. versionadded:: 0.9.0
        """
        return self.impl.select()

    def clear_contents(self):
        """Clears the content of the whole sheet but leaves the formatting."""
        return self.impl.clear_contents()

    def clear(self):
        """Clears the content and formatting of the whole sheet."""
        return self.impl.clear()

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
        >>> Sheet('Sheet1').autofit('c')
        >>> Sheet('Sheet1').autofit('r')
        >>> Range('Sheet1').autofit()


        .. versionadded:: 0.2.3
        """
        return self.impl.autofit(axis)

    def delete(self):
        """
        Deletes the Sheet.


        .. versionadded: 0.6.0
        """
        return self.impl.delete()

    def __repr__(self):
        return "<Sheet [{1}]{0}>".format(self.name, self.book.name)

    @property
    def charts(self):
        """
        Returns a collection of all the charts on the Sheet.

        See Also
        --------

        Charts


        .. versionadded:: 0.9.0
        """
        return Charts(impl=self.impl.charts)

    @property
    def shapes(self):
        """
        Returns a collection of all the shapes on the Sheet.

        See Also
        --------

        Shapes


        .. versionadded:: 0.9.0
        """
        return Shapes(impl=self.impl.shapes)

    @property
    def pictures(self):
        """
        Returns a collection of all the pictures on the Sheet.

        See Also
        --------

        Pictures


        .. versionadded:: 0.9.0
        """
        return Pictures(impl=self.impl.pictures)


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

    def __init__(self, *args, **options):

        # Arguments
        impl = options.pop('impl', None)
        if impl is None:
            if len(args) == 2 and isinstance(args[0], Range) and isinstance(args[1], Range):
                if args[0].sheet != args[1].sheet:
                    raise ValueError("Ranges are not on the same sheet")
                impl = args[0].sheet.range(args[0], args[1]).impl
            elif len(args) == 1 and isinstance(args[0], string_types):
                impl = apps.active.range(args[0]).impl
            elif len(args) == 1 and isinstance(args[0], tuple):
                impl = sheets.active.range(*args).impl
            elif len(args) == 2 and isinstance(args[0], tuple) and isinstance(args[1], tuple):
                impl = sheets.active.range(*args).impl
            else:
                raise ValueError("Invalid arguments")

        self.impl = impl

        # Keyword Arguments
        self._options = options

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __eq__(self, other):
        return (
           isinstance(other, Range)
           and self.sheet == other.sheet
           and self.row == other.row
           and self.column == other.column
           and self.shape == other.shape
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash((self.sheet, self.row, self.column, self.shape))

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

    count = property(__len__)

    @property
    def row(self):
        return self.impl.row

    @property
    def column(self):
        return self.impl.column

    @property
    def raw_value(self):
        return self.impl.raw_value

    @raw_value.setter
    def raw_value(self, data):
        self.impl.raw_value = data

    def clear_contents(self):
        return self.impl.clear_contents()

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

    @property
    def address(self):
        return self.impl.address

    @property
    def current_region(self):
        return Range(impl=self.impl.current_region)

    def autofit(self):
        return self.impl.autofit()

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
        return RangeRows(self)

    @property
    def columns(self):
        return RangeColumns(self)

    @property
    def shape(self):
        """
        Tuple of Range dimensions.

        .. versionadded:: 0.3.0
        """
        return self.impl.shape

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

    def expand(self, mode):
        if mode == 'table':
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

        elif mode in ['vertical', 'd', 'down']:
            if self(2, 1).raw_value in [None, ""]:
                return Range(self(1, 1), self(1, self.shape[1]))
            elif self(3, 1).raw_value in [None, ""]:
                return Range(self(1, 1), self(2, self.shape[1]))
            else:
                end_row = self(2, 1).end('down').row - self.row + 1
                return Range(self(1, 1), self(end_row, self.shape[1]))

        elif mode in ['horizontal', 'r', 'right']:
            if self(1, 2).raw_value in [None, ""]:
                return Range(self(1, 1), self(self.shape[0], 1))
            elif self(1, 3).raw_value in [None, ""]:
                return Range(self(1, 1), self(self.shape[0], 2))
            else:
                end_column = self(1, 2).end('right').column - self.column + 1
                return Range(self(1, 1), self(self.shape[0], end_column))

    def __getitem__(self, key):
        if type(key) is tuple:
            row, col = key

            n = self.shape[0]
            if isinstance(row, slice):
                row1, row2, step = row.indices(n)
                if step != 1:
                    raise ValueError("Slice steps not supported.")
                row2 -= 1
            elif isinstance(row, int):
                if row < 0:
                    row += n
                if row < 0 or row >= n:
                    raise IndexError("Row index %s out of range (%s rows)." % (row, n))
                row1 = row2 = row
            else:
                raise TypeError("Row indices must be integers or slices, not %s" % type(row).__name__)

            n = self.shape[1]
            if isinstance(col, slice):
                col1, col2, step = col.indices(n)
                if step != 1:
                    raise ValueError("Slice steps not supported.")
                col2 -= 1
            elif isinstance(col, int):
                if col < 0:
                    col += n
                if col < 0 or col >= n:
                    raise IndexError("Column index %s out of range (%s columns)." % (col, n))
                col1 = col2 = col
            else:
                raise TypeError("Column indices must be integers or slices, not %s" % type(col).__name__)

            return self.sheet.range((
                self.row + row1,
                self.column + col1,
                max(0, row2 - row1 + 1),
                max(0, col2 - col1 + 1)
            ))

        elif isinstance(key, slice):
            if self.shape[0] > 1 and self.shape[1] > 1:
                raise IndexError("One-dimensional slicing is not allowed on two-dimensional ranges")

            if self.shape[0] > 1:
                return self[key, :]
            else:
                return self[:, key]

        elif isinstance(key, int):
            n = len(self)
            k = key + n if key < 0 else key
            if k < 0 or k >= n:
                raise IndexError("Index %s out of range (%s elements)." % (key, n))
            else:
                return self(k + 1)

        else:
            raise TypeError("Cell indices must be integers or slices, not %s" % type(key).__name__)

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
            return self.impl.hyperlink

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
        self.impl.add_hyperlink(address, text_to_display, screen_tip)

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
            row_size = self.shape[0]
        if column_size is not None:
            assert column_size > 0
        else:
            column_size = self.shape[1]

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
                row_offset + self.shape[0],
                column_offset + self.shape[1]
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
        return self(self.shape[0], self.shape[1]).options(**self._options)

    def select(self):
        # Select only works on the active sheet
        self.impl.select()


# This has to be after definition of Range to resolve circular reference
from . import conversion


class Ranges(object):
    pass


class RangeRows(Ranges):

    def __init__(self, rng):
        self.rng = rng

    def __len__(self):
        return self.rng.shape[0]

    count = property(__len__)

    def autofit(self):
        self.rng.impl.autofit(axis='r')

    def __iter__(self):
        for i in range(0, self.rng.shape[0]):
            yield self.rng[i, :]

    def __call__(self, key):
        return self.rng[key-1, :]

    def __getitem__(self, key):
        if isinstance(key, slice):
            return RangeRows(rng=self.rng[key, :])
        elif isinstance(key, int):
            return self.rng[key, :]
        else:
            raise TypeError("Indices must be integers or slices, not %s" % type(key).__name__)

    def __repr__(self):
        return '{}({})'.format(
            self.__class__.__name__,
            repr(self.rng)
        )


class RangeColumns(Ranges):

    def __init__(self, rng):
        self.rng = rng

    def __len__(self):
        return self.rng.shape[1]

    count = property(__len__)

    def autofit(self):
        self.rng.impl.autofit(axis='c')

    def __iter__(self):
        for j in range(0, self.rng.shape[1]):
            yield self.rng[:, j]

    def __call__(self, key):
        return self.rng[:, key-1]

    def __getitem__(self, key):
        if isinstance(key, slice):
            return RangeRows(rng=self.rng[:, key])
        elif isinstance(key, int):
            return self.rng[:, key]
        else:
            raise TypeError("Indices must be integers or slices, not %s" % type(key).__name__)

    def __repr__(self):
        return '{}({})'.format(
            self.__class__.__name__,
            repr(self.rng)
        )


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
    def __init__(self, *args, **options):
        impl = options.pop('impl', None)
        if impl is None:
            if len(args) == 1:
                impl = sheets.active.shapes(args[0]).impl

            else:
                raise ValueError("Invalid arguments")

        self.impl = impl

    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def type(self):
        return self.impl.type

    @property
    def left(self):
        return self.impl.left

    @left.setter
    def left(self, value):
        self.impl.left = value

    @property
    def top(self):
        return self.impl.top

    @top.setter
    def top(self, value):
        self.impl.top = value

    @property
    def width(self):
        return self.impl.width

    @width.setter
    def width(self, value):
        self.impl.width = value

    @property
    def height(self):
        return self.impl.height

    @height.setter
    def height(self, value):
        self.impl.height = value

    def delete(self):
        self.impl.delete()

    def activate(self):
        self.impl.activate()

    @property
    def parent(self):
        return Sheet(impl=self.impl.parent)

    def __eq__(self, other):
        return (
            isinstance(other, Shape) and
            other.parent == self.parent and
            other.name == self.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __repr__(self):
        return "<Shape '{0}' in {1}>".format(
            self.name,
            self.parent
        )


class Shapes(Collection):
    _wrap = Shape


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

    def __init__(self, name_or_index=None, impl=None):
        if impl is not None:
            self.impl = impl
        elif name_or_index is not None:
            self.impl = sheets.active.charts(name_or_index).impl
        else:
            self.impl = sheets.active.charts.add().impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def parent(self):
        impl = self.impl.parent
        if isinstance(impl, xlplatform.Book):
            return Book(impl=self.impl.parent)
        else:
            return Sheet(impl=self.impl.parent)

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

    source_data = property(None, set_source_data)

    @property
    def left(self):
        return self.impl.left

    @left.setter
    def left(self, value):
        self.impl.left = value

    @property
    def top(self):
        return self.impl.top

    @top.setter
    def top(self, value):
        self.impl.top = value

    @property
    def width(self):
        return self.impl.width

    @width.setter
    def width(self, value):
        self.impl.width = value

    @property
    def height(self):
        return self.impl.height

    @height.setter
    def height(self, value):
        self.impl.height = value

    def delete(self):
        self.impl.delete()

    def __repr__(self):
        return "<Chart '{0}' in {1}>".format(
            self.name,
            self.parent
        )


class Charts(Collection):
    _wrap = Chart

    def add(self, left=0, top=0, width=355, height=211):
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

        impl = self.impl.add(
            left,
            top,
            width,
            height
        )

        return Chart(impl=impl)


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
    def __init__(self, impl=None):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    @property
    def parent(self):
        return Sheet(impl=self.impl.parent)

    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        if value in self.parent.pictures:
            if value == self.name:
                return
            else:
                raise ShapeAlreadyExists()

        self.impl.name = value

    @property
    def left(self):
        return self.impl.left

    @left.setter
    def left(self, value):
        self.impl.left = value

    @property
    def top(self):
        return self.impl.top

    @top.setter
    def top(self, value):
        self.impl.top = value

    @property
    def width(self):
        return self.impl.width

    @width.setter
    def width(self, value):
        self.impl.width = value

    @property
    def height(self):
        return self.impl.height

    @height.setter
    def height(self, value):
        self.impl.height = value

    def delete(self):
        self.impl.delete()

    def __eq__(self, other):
        return (
            isinstance(other, Picture) and
            other.parent == self.parent and
            other.name == self.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __repr__(self):
        return "<Picture '{0}' in {1}>".format(
            self.name,
            self.parent
        )

    def update(self, image):
        """
        Replaces an existing picture with a new one, taking over the attributes of the existing picture.

        Arguments
        ---------

        filename : str
            Path to the picture.


        .. versionadded:: 0.5.0
        """

        filename, width, height = utils.process_image(image, self.width, self.height)

        left, top = self.left, self.top
        name = self.name

        # todo: link_to_file, save_with_document
        picture = self.parent.pictures.add(filename, left=left, top=top, width=width, height=height)
        self.delete()

        picture.name = name
        self.impl = picture.impl

        return picture


class Pictures(Collection):
    _wrap = Picture

    @property
    def parent(self):
        return Sheet(impl=self.impl.parent)

    def add(self, image, link_to_file=False, save_with_document=True, left=0, top=0, width=None, height=None, name=None, update=False):

        if update:
            if name is None:
                raise ValueError("If update is true then name must be specified")
            else:
                try:
                    pic = self[name]
                    pic.update(image)
                    return pic
                except KeyError:
                    pass

        filename, width, height = utils.process_image(image, width, height)

        if not (link_to_file or save_with_document):
            raise Exception("Arguments link_to_file and save_with_document cannot both be false")

        # Image dimensions
        im_width, im_height = None, None
        if width is None or height is None:
            if Image:
                with Image.open(filename) as im:
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

        picture = Picture(impl=self.impl.add(
            filename, link_to_file, save_with_document, left, top, width, height
        ))

        if name is not None:
            picture.name = name
        return picture


class Names(object):

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __call__(self, name_or_index):
        return Name(impl=self.impl(name_or_index))

    def contains(self, name_or_index):
        return self.impl.contains(name_or_index)

    def __len__(self):
        return len(self.impl)

    count = property(__len__)

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
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being used.


        .. versionadded:: 0.9.0
        """
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
    

def view(obj, name=None, sheet=None):
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
    if sheet is None:
        sheet = Book().sheets.active
    else:
        sheet.clear()

    if mpl and isinstance(obj, mpl.figure.Figure):
        return sheet.pictures.add(obj, name=name, update=name is not None)
    else:
        sheet.range('A1').value = obj
        sheet.autofit()


class Macro(object):
    def __init__(self, app, macro):
        self.app = app
        self.macro = macro

    def run(self, *args):
        return self.app.impl.run(self.macro, args)

    __call__ = run


class Books(Collection):
    """
    ``xw.books`` returns all books in the active app. Use ``app.books`` to return all books of a specific app.
    """
    _wrap = Book

    @property
    def active(self):
        """
        Returns the active Book.
        """
        return Book(impl=self.impl.active)

    def add(self):
        """
        Creates a new Book. The new Book becomes the active Book. Returns a Book object.
        """
        return Book(impl=self.impl.add())

    def open(self, fullname):
        """
        Opens a Book if it is not open yet and returns it. If it is already open, it doesn't raise an exception but
        simply returns the Book object.

        Parameters
        ----------
        fullname : str
            filename or fully qualified filename, e.g. ``r'C:\\path\\to\\file.xlsx'`` or ``'file.xlsm'``. Without a full
            path, it looks for the file in the current working directory.

        Returns
        -------
        Book : Book that has been opened.

        """
        if not os.path.exists(fullname):
            raise FileNotFoundError("No such file: '%s'" % fullname)
        fullname = os.path.realpath(fullname)
        _, name = os.path.split(fullname)
        try:
            impl = self.impl(name)
            # on windows, samefile only available on Py>=3.2
            if hasattr(os.path, 'samefile'):
                throw = not os.path.samefile(impl.fullname, fullname)
            else:
                throw = (os.path.normpath(os.path.realpath(impl.fullname)) != os.path.normpath(fullname))
            if throw:
                raise ValueError(
                    "Cannot open two workbooks named '%s', even if they are saved in different locations." % name
                )
        except KeyError:
            impl = self.impl.open(fullname)
        return Book(impl=impl)


class Sheets(Collection):

    _wrap = Sheet

    @property
    def active(self):
        """
        Returns the active Sheet.
        """
        return Sheet(impl=self.impl.active)

    def __call__(self, name_or_index):
        if isinstance(name_or_index, Sheet):
            return name_or_index
        else:
            return Sheet(impl=self.impl(name_or_index))

    def __delitem__(self, name_or_index):
        self[name_or_index].delete()

    def add(self, name=None, before=None, after=None):
        """
        Creates a new Sheet and makes it the active sheet.

        Parameters
        ----------
        name : str, default None
            Name of the new sheet. If None, will default to Excel's default name.
        before : Sheet, default None
            An object that specifies the sheet before which the new sheet is added.
        after : Sheet, default None
            An object that specifies the sheet after which the new sheet is added.

        Returns
        -------

        """
        if name is not None:
            if name.lower() in (s.name.lower() for s in self):
                raise ValueError("Sheet named '%s' already present in workbook" % name)
        if before is not None and not isinstance(before, Sheet):
            before = self(before)
        if after is not None and not isinstance(after, Sheet):
            after = self(after)
        impl = self.impl.add(before and before.impl, after and after.impl)
        if name is not None:
            impl.name = name
        return Sheet(impl=impl)


class ActiveAppBooks(Books):

    def __init__(self):
        pass

    # override class name which appears in repr
    _name = 'Books'

    @property
    def impl(self):
        return apps.active.books.impl


class ActiveBookSheets(Sheets):

    def __init__(self):
        pass

    # override class name which appears in repr
    _name = 'Sheets'

    @property
    def impl(self):
        return books.active.sheets.impl


books = ActiveAppBooks()

sheets = ActiveBookSheets()
