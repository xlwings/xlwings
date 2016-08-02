.. _connect_to_workbook:

Connect to a Book
=================

When reading/writing data to the active sheet, you don't need a book object:

  >>> import xlwings as xw
  >>> xw.Range('A1').value = 'something'

Python to Excel
---------------

The easiest way to connect to a book is offered by ``xw.Book``: it looks for the book in all app instances and
returns an error, should the same book be open in multiple instances.
To connect to a book in the active app instance, use ``xw.books`` and to refer to a specific app, use:

>>> app = xw.App()  # or something like xw.apps[0] for existing apps
>>> app.books['Book1']

+--------------------+--------------------------------------+--------------------------------------------+
|                    | xw.Book                              | xw.books                                   |
+====================+======================================+============================================+
| New book           | ``xw.Book()``                        | ``xw.books.add()``                         |
+--------------------+--------------------------------------+--------------------------------------------+
| Unsaved book       | ``xw.Book('Book1')``                 | ``xw.books['Book1']``                      |
+--------------------+--------------------------------------+--------------------------------------------+
| Book by (full)name | ``xw.Book(r'C:/path/to/file.xlsx')`` | ``xw.books.open(r'C:/path/to/file.xlsx')`` |
+--------------------+--------------------------------------+--------------------------------------------+

.. note::
  When specifying file paths on Windows, you should either use raw strings by putting
  an ``r`` in front of the string or use double back-slashes like so: ``C:\\path\\to\\file.xlsx``.

Excel to Python (RunPython)
---------------------------

To reference the calling book when using ``RunPython`` in VBA, use ``xw.Book.caller()``, see
:ref:`run_python`.
Check out the section about :ref:`debugging` to see how you can call a script from both sides, Python and Excel, without
the need to constantly change between ``xw.Book.caller()`` and one of the methods explained above.

User Defined Functions (UDFs)
-----------------------------

Unlike ``RunPython``, UDFs don't need a call to ``xw.Book.caller()``, see :ref:`udfs`.
However, it's available (restricted to read-only though), which sometimes proofs to be useful.
