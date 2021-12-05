.. _syntax_overview:

Syntax Overview
===============

The xlwings object model is very similar to the one used by VBA.

All code samples below depend on the following import:

    >>> import xlwings as xw

Active Objects
--------------
::

    # Active app (i.e. Excel instance)
    >>> app = xw.apps.active

    # Active book
    >>> wb = xw.books.active  # in active app
    >>> wb = app.books.active  # in specific app

    # Active sheet
    >>> sheet = xw.sheets.active  # in active book
    >>> sheet = wb.sheets.active  # in specific book

    # Range on active sheet
    >>> xw.Range('A1')  # on active sheet of active book of active app

A Range can be instantiated with A1 notation, a tuple of Excel's 1-based indices, a named range or two Range objects:

::

    xw.Range('A1')
    xw.Range('A1:C3')
    xw.Range((1,1))
    xw.Range((1,1), (3,3))
    xw.Range('NamedRange')
    xw.Range(xw.Range('A1'), xw.Range('B2'))

Full qualification
------------------

Round brackets follow Excel's behavior (i.e. 1-based indexing), while square brackets use Python's 0-based indexing/slicing.
As an example, the following expressions all reference the same range::

    xw.apps[763].books[0].sheets[0].range('A1')
    xw.apps(10559).books(1).sheets(1).range('A1')
    xw.apps[763].books['Book1'].sheets['Sheet1'].range('A1')
    xw.apps(10559).books('Book1').sheets('Sheet1').range('A1')

Note that the apps keys are different for you as they are the process IDs (PID). You can get the list of your PIDs via
``xw.apps.keys()``.

App context manager
-------------------

If you want to open a new Excel instance via ``App()``, you usually should use ``App`` as a context manager as this will make sure that the Excel instance is closed and cleaned up again properly::

    with xw.App() as app:
        book = app.books['Book1']

Range indexing/slicing
----------------------

Range objects support indexing and slicing, a few examples:

>>> rng = xw.Book().sheets[0].range('A1:D5')
>>> rng[0, 0]
 <Range [Workbook1]Sheet1!$A$1>
>>> rng[1]
 <Range [Workbook1]Sheet1!$B$1>
>>> rng[:, 3:]
<Range [Workbook1]Sheet1!$D$1:$D$5>
>>> rng[1:3, 1:3]
<Range [Workbook1]Sheet1!$B$2:$C$3>

Range Shortcuts
---------------

Sheet objects offer a shortcut for range objects by using index/slice notation on the sheet object. This evaluates to either
``sheet.range`` or ``sheet.cells`` depending on whether you pass a string or indices/slices:

    >>> sheet = xw.Book().sheets['Sheet1']
    >>> sheet['A1']
    <Range [Book1]Sheet1!$A$1>
    >>> sheet['A1:B5']
    <Range [Book1]Sheet1!$A$1:$B$5>
    >>> sheet[0, 1]
    <Range [Book1]Sheet1!$B$1>
    >>> sheet[:10, :10]
    <Range [Book1]Sheet1!$A$1:$J$10>

Object Hierarchy
----------------

The following shows an example of the object hierarchy, i.e. how to get from an app to a range object
and all the way back:

>>> rng = xw.apps[10559].books[0].sheets[0].range('A1')
>>> rng.sheet.book.app
<Excel App 10559>
