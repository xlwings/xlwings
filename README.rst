xlwings - Make Excel fly with Python!
=====================================

xlwings is a `BSD-licensed <http://opensource.org/licenses/BSD-3-Clause>`_ Python library that makes it easy to call
Python from Excel and vice versa:

* Interact with Excel from Python using a syntax that is close to VBA yet Pythonic.
* Replace your VBA macros with Python code and still pass around your workbooks as easy as before.

xlwings fully supports NumPy arrays and Pandas DataFrames. It works with Microsoft Excel on Windows and Mac.

.. note:: xlwings is currently in an early stage.
   The API might change in backward incompatible ways.

Interact with Excel from Python
-------------------------------

Writing/reading values to/from Excel and adding a chart is as easy as:

.. code-block:: python

    >>> from xlwings import Workbook, Sheet, Range, Chart
    >>> wb = Workbook()  # Creates a connection with a new workbook
    >>> Range('A1').value = 'Foo 1'
    >>> Range('A1').value
    'Foo 1'
    >>> Range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> Range('A1').table.value  # or: Range('A1:C2').value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> Sheet(1).name
    'Sheet1'
    >>> chart = Chart.add(source_data=Range('A1').table)

The Range and Chart objects as used above will refer to the active sheet of the current Workbook ``wb``. Include the
Sheet name like this:

.. code-block:: python

    Range('Sheet1', 'A1:C3').value
    Range(1, (1,1), (3,3)).value  # index notation
    Chart.add('Sheet1', source_data=Range('Sheet1', 'A1').table)

Qualify the Workbook additionally like this:

.. code-block:: python

    Range('Sheet1', 'A1', wkb=wb).value
    Chart.add('Sheet1', wkb=wb, source_data=Range('Sheet1', 'A1', wkb=wb).table)
    Sheet(1, wkb=wb).name

or simply set the current workbook first:

.. code-block:: python

    wb.set_current()
    Range('Sheet1', 'A1').value
    Chart.add('Sheet1', source_data=Range('Sheet1', 'A1').table)
    Sheet(1).name

These commands also work seamlessly with **NumPy arrays** and **Pandas DataFrames**.

Call Python from Excel
----------------------

If, for example, you want to fill your spreadsheet
with standard normally distributed random numbers, your VBA code is just one line:

.. code-block:: vb.net

    Sub RandomNumbers()
        RunPython ("import mymodule; mymodule.rand_numbers()")
    End Sub

This essentially hands over control to ``mymodule.py``:

.. code-block:: python

    import numpy as np
    from xlwings import Workbook, Range

    def rand_numbers():
        """ produces standard normally distributed random numbers with shape (n,n)"""
        wb = Workbook.caller()  # Creates a reference to the calling Excel file
        n = int(Range('Sheet1', 'B1').value)  # Write desired dimensions into Cell B1
        rand_num = np.random.randn(n, n)
        Range('Sheet1', 'C3').value = rand_num

To make this run, just import the VBA module ``xlwings.bas`` in the VBA editor (Open the VBA editor with ``Alt-F11``,
then go to ``File > Import File...`` and import the ``xlwings.bas`` file. ). It can be found in the directory of
your ``xlwings`` installation.

.. note:: Always instantiate the ``Workbook`` within the function that is called from Excel and not outside as global
    variable. Older versions of the docs/samples were showing the wrong approach.

Easy deployment
---------------

Deployment is really the part where xlwings shines:

* Just zip-up your Spreadsheet with your Python code and send it around. The receiver only needs to have an
  installation of Python with xlwings (and obviously all the other packages you're using).
* There is no need to install any Excel add-in.
* If this still sounds too complicated, just freeze your Python code into an executable and use
  ``RunFrozenPython`` instead of ``RunPython``. This gives you a standalone version of your Spreadsheet tool without any
  dependencies (``RunFrozenPython`` is currently only available on Windows).

Installation
------------

The easiest way to install xlwings is via pip::

    pip install xlwings


Alternatively it can be installed from source. From within the ``xlwings`` directory, execute::

    python setup.py install

Dependencies
------------

* **Windows**: ``pywin32``, ``comtypes``

  On Windows, it is recommended to use one of the scientific Python distributions like
  `Anaconda <https://store.continuum.io/cshop/anaconda/>`_,
  `WinPython <https://winpython.github.io/>`_ or
  `Canopy <https://www.enthought.com/products/canopy/>`_ as they already include pywin32. Otherwise it needs to be
  installed from `here <http://sourceforge.net/projects/pywin32/files/pywin32/>`_.

* **Mac**: ``psutil``, ``appscript``

  On Mac, the dependencies are automatically being handled if xlwings is installed with ``pip``. However,
  the Xcode command line tools need to be available. Mac OS X 10.4 (*Tiger*) or later is required.
  The recommended Python distribution for Mac is `Anaconda <https://store.continuum.io/cshop/anaconda/>`_.

Optional Dependencies
---------------------

* NumPy
* Pandas

These packages are not required but highly recommended as NumPy arrays and Pandas DataFrames/Series play very nicely
with xlwings.

Python version support
----------------------

xlwings runs on Python 2.6-2.7 and 3.1-3.4

Links
-----

* Homepage: http://xlwings.org
* Documentation: http://docs.xlwings.org
* Source Code: http://github.com/zoomeranalytics/xlwings


