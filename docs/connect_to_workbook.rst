.. _connect_to_workbook:

Connect to Workbooks
====================

First things first: to be able to do anything with xlwings, you need to establish a connection with an
Excel Workbook first.

Python to Excel
---------------

There are various ways to connect to an Excel workbook from Python. Create a connection with...

* a new workbook: ``wb = Workbook()``
* the active workbook (works with all instances of Excel on Windows): ``wb = Workbook.active()``
* an unsaved workbook: ``wb = Workbook('Book1')``
* a saved (open) workbook by name (incl. xlsx etc.): ``wb = Workbook('MyWorkbook.xlsx')``
* a saved (open or closed) workbook by path: ``wb = Workbook(r'C:\path\to\file.xlsx')``

Note that when specifying file paths on Windows, you should either use raw strings by
putting an ``r`` in front of the string or by using double back-slashes like so: ``C:\\path\\to\\file.xlsx``.

Excel to Python
---------------

To make a connection from Excel, i.e. when calling a Python script with ``RunPython``, use ``Workbook.caller()``.
Check out the section about :ref:`debugging` to see how you can call a script from both sides, Python and Excel, without
the need to constantly change between ``Workbook.caller()`` and one of the methods explained above.