.. _syntax_overview:

Syntax Overview
===============

The xlwings object model is very similar to the one used by VBA.

All code samples below depend on the following import:

    >>> import xlwings as xw

Active Objects
--------------

* Active app (i.e. Excel instance)

  >>> app = xw.apps.active

* Active book

  >>> wb = xw.books.active  # in active app
  >>> wb = app.books.active  # in specific app

* Active sheet

  >>> sht = xw.sheets.active  # in active book
  >>> sht = wb.sheets.active  # in specific book

* Range on active sheet

  >>> xw.Range('A1')  # on active sheet of active book of active app

Round vs. Square Brackets
-------------------------

Round brackets follow Excel's behavior (i.e. 1-based indexing), while square brackets use Python's 0-based indexing/slicing.

As an example, the following all reference the same range::

    xw.apps[0].books[0].sheets[0].range('A1')
    xw.apps(1).books(1).sheets(1).range('A1')
    xw.apps[0].books['Book1'].sheets['Sheet1'].range('A1')
    xw.apps(1).books('Book1').sheets('Sheet1').range('A1')


Range
-----

A Range object can be instantiated using A1 notation, a tuple of Excel-1-based indexes, a named range or by
by using two Range objects:

::

    xw.Range('A1')
    xw.Range('A1:C3')
    xw.Range((1,1))
    xw.Range((1,1), (3,3))
    xw.Range('NamedRange')
    xw.Range(xw.Range('A1'), xw.Range('B2'))

Range indexing/slicing
----------------------

Range objects support indexing and slicing, a few examples:

>>> rng = xw.Book().sheets[0].range('A1:D5')
>>> rng[0, 0]
 <Range [Workbook1]Sheet1!$A$1>
>>> rng[1]
 <Range [Workbook6]Sheet1!$B$1>
>>> rng[:, 3:]
<Range [Workbook6]Sheet1!$D$1:$D$5>
>>> rng[1:3, 1:3]
<Range [Workbook6]Sheet1!$B$2:$C$3>

Range Shortcuts
---------------

Sheet objects offer a shortcut to access range objects by using index/slice notation on the sheet object. This evaluates to either
``sheet.range`` or ``sheet.cells`` depending on whether you pass a string or indices/slices:

    >>> import xlwings as xw
    >>> sht = xw.Book().sheets['Sheet1']
    >>> sht['A1']
    <Range [Book1]Sheet1!$A$1>
    >>> sht['A1:B5']
    <Range [Book1]Sheet1!$A$1:$B$5>
    >>> sht[0, 1]
    <Range [Book1]Sheet1!$B$1>
    >>> sht[:10, :10]
    <Range [Book1]Sheet1!$A$1:$J$10>

