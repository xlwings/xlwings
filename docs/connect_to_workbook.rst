.. _connect_to_workbook:

Connect to Workbooks
====================

First things first: to be able to talk to Excel from Python or call ``RunPython`` in VBA, you need to establish a connection with
an Excel Workbook.

Python to Excel
---------------

There are various ways to connect to an Excel workbook from Python:

* ``wb = Workbook()`` connects to a a new workbook
* ``wb = Workbook.active()`` connects to the active workbook (supports multiple Excel instances)
* ``wb = Workbook('Book1')`` connects to an unsaved workbook
* ``wb = Workbook('MyWorkbook.xlsx')`` connects to a saved (open) workbook by name (incl. xlsx etc.)
* ``wb = Workbook(r'C:\path\to\file.xlsx')`` connects to a saved (open or closed) workbook by path

.. note::
  When specifying file paths on Windows, you should either use raw strings by putting
  an ``r`` in front of the string or use double back-slashes like so: ``C:\\path\\to\\file.xlsx``.

Excel to Python (RunPython)
---------------------------

To make a connection from Excel, i.e. when calling a Python script with ``RunPython``, use ``Workbook.caller()``, see
:ref:`run_python`.
Check out the section about :ref:`debugging` to see how you can call a script from both sides, Python and Excel, without
the need to constantly change between ``Workbook.caller()`` and one of the methods explained above.

User Defined Functions (UDFs)
-----------------------------

UDFs work differently and don't need the explicit instantiation of a ``Workbook``, see :ref:`udfs`.
However, ``xw.Workbook.caller()`` can be used in UDFs although just read-only.
