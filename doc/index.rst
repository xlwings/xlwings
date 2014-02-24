.. xlwings documentation master file, created by
   sphinx-quickstart on Sat Feb 22 22:00:11 2014.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

xlwings - make Excel fly!
=========================

xlwings is a BSD-licensed Python library that allows you to call Python from Excel and vice versa:

* Interact with Excel from Python using a syntax that is close to VBA yet Pythonic.
* Replace your VBA macros with Python code and still pass around your workbooks easily.

Currently, xlwings only works on Windows.

Interact with Excel from Python
-------------------------------
.. code-block:: python

    >>> from xlwings import Workbook, Range, Chart
    >>> wb = Workbook()  # Creates a connection with a new workbook
    >>> Range('A1').value = ['Foo 1', 'Foo 2', 'Foo 3', 'Foo 4']
    >>> Range('A2').value = [10, 20, 30, 40]
    >>> Range('A1').table.value
    [[u'Foo 1', u'Foo 2', u'Foo 3', u'Foo 4'], [10.0, 20.0, 30.0, 40.0]]
    >>> chart = Chart().add()
    >>> chart.set_source_data(Range('A1').table)


Replace VBA macros with Python
------------------------------

Let's say you want to fill your spreadsheet with standard normally distributed random numbers. Your VBA code is just one
line:

.. code-block:: vb.net

    Sub RandomNumbers()
        RunPython ("import mymodule; mymodule.rand_numbers()")
    End Sub

This essentially hands over control to ``mymodule.py``:

.. code-block:: python

    import numpy as np
    from xlwings import Workbook, Range

    wb = Workbook()  # Creates a reference to the calling Excel file

    def rand_numbers():
        """ produces standard normally distributed random numbers with shape (n,n)"""
        n = Range('Sheet1', 'B1').value
        rand_num = np.random.randn(n, n)
        Range('Sheet1', 'C3').value = rand_num


To make this run, just import de VBA module ``xlwings.bas`` in the VBA editor. It can be found in the directory of
your xlwings installation.

Easy deployment
---------------

* Just zip-up your Spreadsheet with your Python code and the ``xlwings.py`` file and send it around. The receiver only
  needs to have an installation of Python with `pywin32 <http://sourceforge.net/projects/pywin32/>`_ (and obviously
  all the other packages you're using).
* There is no need to install any Excel add-in.
* If this still sounds too complicated, just freeze your Python code into an exectuable and use
  ``RunFrozenPython`` instead ``RunPython``. This gives youa a standalone version of your Spreadsheet tool without any
  dependencies.







