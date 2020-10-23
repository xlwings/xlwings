.. _udfs:

User Defined Functions (UDFs)
=============================

This tutorial gets you quickly started on how to write User Defined Functions.

.. note::
    * UDFs are currently only available on Windows.
    * For details of how to control the behaviour of the arguments and return values, have a look at :ref:`converters`.
    * For a comprehensive overview of the available decorators and their options, check out the corresponding API docs: :ref:`udf_api`.

One-time Excel preparations
---------------------------

1) Enable ``Trust access to the VBA project object model`` under
``File > Options > Trust Center > Trust Center Settings > Macro Settings``

2) Install the add-in via command prompt: ``xlwings addin install`` (see :ref:`xlwings_addin`).


Workbook preparation
--------------------

The easiest way to start a new project is to run ``xlwings quickstart myproject`` on a command prompt (see :ref:`command_line`).
This automatically adds the xlwings reference to the generated workbook.

A simple UDF
------------

The default addin settings expect a Python source file in the way it is created by ``quickstart``:

* in the same directory as the Excel file
* with the same name as the Excel file, but with a ``.py`` ending instead of ``.xlsm``.

Alternatively, you can point to a specific module via ``UDF Modules`` in the xlwings ribbon.

Let's assume you have a Workbook ``myproject.xlsm``, then you would write the following code in ``myproject.py``::

    import xlwings as xw

    @xw.func
    def double_sum(x, y):
        """Returns twice the sum of the two arguments"""
        return 2 * (x + y)


* Now click on ``Import Python UDFs`` in the xlwings tab to pick up the changes made to ``myproject.py``.
* Enter the formula ``=double_sum(1, 2)`` into a cell and you will see the correct result:

  .. figure:: images/double_sum.png
    :scale: 80%

* The docstring (in triple-quotes) will be shown as function description in Excel.

.. note::
  * You only need to re-import your functions if you change the function arguments or the function name.
  * Code changes in the actual functions are picked up automatically (i.e. at the next calculation of the formula,
    e.g. triggered by ``Ctrl-Alt-F9``), but changes in imported modules are not. This is the very behaviour of how Python
    imports work. If you want to make sure everything is in a fresh state, click ``Restart UDF Server``.
  * The ``@xw.func`` decorator is only used by xlwings when the function is being imported into Excel. It tells xlwings
    for which functions it should create a VBA wrapper function, otherwise it has no effect on how the functions behave
    in Python.


Array formulas: Get efficient
-----------------------------

Calling one big array formula in Excel is much more efficient than calling many single-cell formulas, so it's generally
a good idea to use them, especially if you hit performance problems.

You can pass an Excel Range as a function argument, as opposed to a single cell and it will show up in Python as
list of lists.

For example, you can write the following function to add 1 to every cell in a Range::

    @xw.func
    def add_one(data):
        return [[cell + 1 for cell in row] for row in data]

To use this formula in Excel,

* Click on ``Import Python UDFs`` again
* Fill in the values in the range ``A1:B2``
* Select the range ``D1:E2``
* Type in the formula ``=add_one(A1:B2)``
* Press ``Ctrl+Shift+Enter`` to create an array formula. If you did everything correctly, you'll see the formula
  surrounded by curly braces as in this screenshot:

.. figure:: images/array_formula.png
    :scale: 80%

Number of array dimensions: ndim
********************************

The above formula has the issue that it expects a "two dimensional" input, e.g. a nested list of the form
``[[1, 2], [3, 4]]``.
Therefore, if you would apply the formula to a single cell, you would get the following error:
``TypeError: 'float' object is not iterable``.

To force Excel to always give you a two-dimensional array, no matter whether the argument is a single cell, a
column/row or a two-dimensional Range, you can extend the above formula like this::

    @xw.func
    @xw.arg('data', ndim=2)
    def add_one(data):
        return [[cell + 1 for cell in row] for row in data]

Array formulas with NumPy and Pandas
------------------------------------

Often, you'll want to use NumPy arrays or Pandas DataFrames in your UDF, as this unlocks the full power of Python's
ecosystem for scientific computing.

To define a formula for matrix multiplication using numpy arrays, you would define the following function::

    import xlwings as xw
    import numpy as np

    @xw.func
    @xw.arg('x', np.array, ndim=2)
    @xw.arg('y', np.array, ndim=2)
    def matrix_mult(x, y):
        return x @ y

.. note:: If you are not on Python >= 3.5 with NumPy >= 1.10, use ``x.dot(y)`` instead of ``x @ y``.

A great example of how you can put Pandas at work is the creation of an array-based ``CORREL`` formula. Excel's
version of ``CORREL`` only works on 2 datasets and is cumbersome to use if you want to quickly get the correlation
matrix of a few time-series, for example. Pandas makes the creation of an array-based ``CORREL2`` formula basically
a one-liner::

    import xlwings as xw
    import pandas as pd

    @xw.func
    @xw.arg('x', pd.DataFrame, index=False, header=False)
    @xw.ret(index=False, header=False)
    def CORREL2(x):
        """Like CORREL, but as array formula for more than 2 data sets"""
        return x.corr()


@xw.arg and @xw.ret decorators
------------------------------

These decorators are to UDFs what the ``options`` method is to ``Range`` objects: they allow you to apply converters and their
options to function arguments (``@xw.arg``) and to the return value (``@xw.ret``). For example, to convert the argument ``x`` into
a pandas DataFrame and suppress the index when returning it, you would do the following::

    @xw.func
    @xw.arg('x', pd.DataFrame)
    @xw.ret(index=False)
    def myfunction(x):
       # x is a DataFrame, do something with it
       return x

For further details see the :ref:`converters` documentation.

Dynamic Array Formulas
----------------------

.. note::
    If your version of Excel supports the new native dynamic arrays, then you don't have to do anything special, 
    and you shouldn't use the ``expand`` decorator! To check if your version of Excel supports it, see if you
    have the ``=UNIQUE()`` formula available. Native dynamic arrays were introduced in Office 365 Insider Fast
    at the end of September 2018.

As seen above, to use Excel's array formulas, you need to specify their dimensions up front by selecting the
result array first, then entering the formula and finally hitting ``Ctrl-Shift-Enter``. In practice, it often turns
out to be a cumbersome process, especially when working with dynamic arrays such as time series data.
Since v0.10, xlwings offers dynamic UDF expansion:

This is a simple example that demonstrates the syntax and effect of UDF expansion:

.. code-block:: python

    import numpy as np

    @xw.func
    @xw.ret(expand='table')
    def dynamic_array(r, c):
        return np.random.randn(int(r), int(c))

.. figure:: images/dynamic_array1.png
  :scale: 40%

.. figure:: images/dynamic_array2.png
  :scale: 40%

.. note::
    * Expanding array formulas will overwrite cells without prompting
    * Pre v0.15.0 doesn't allow to have volatile functions as arguments, e.g. you cannot use functions like ``=TODAY()`` as arguments.
      Starting with v0.15.0, you can use volatile functions as input, but the UDF will be called more than 1x.
    * Dynamic Arrays have been refactored with v0.15.0 to be proper legacy arrays: To edit a dynamic array
      with xlwings >= v0.15.0, you need to hit ``Ctrl-Shift-Enter`` while in the top left cell. Note that you don't
      have to do that when you enter the formula for the first time.

Docstrings
----------

The following sample shows how to include docstrings both for the function and for the arguments x and y that then
show up in the function wizard in Excel:

.. code-block:: python

    import xlwings as xw

    @xw.func
    @xw.arg('x', doc='This is x.')
    @xw.arg('y', doc='This is y.')
    def double_sum(x, y):
        """Returns twice the sum of the two arguments"""
        return 2 * (x + y)


The "caller" argument
---------------------

You often need to know which cell called the UDF. For this, xlwings offers the reserved argument ``caller`` which returns the calling cell as xlwings range object::

    @xw.func
    def get_caller_address(caller):
        # caller will not be exposed in Excel, so use it like so:
        # =get_caller_address()
        return caller.address

Note that ``caller`` will not be exposed in Excel but will be provided by xlwings behind the scenes.

The "vba" keyword
-----------------

By using the ``vba`` keyword, you can get access to any Excel VBA object in the form of a pywin32 object. For example, if you wanted to pass the sheet object in the form of its ``CodeName``, you can do it as follows::

    @xw.func
    @xw.arg('sheet1', vba='Sheet1')
    def get_name(sheet1):
        # call this function in Excel with:
        # =get_name()
        return sheet1.Name

Note that ``vba`` arguments are not exposed in the UDF but automatically provided by xlwings.

.. _decorator_macros:

Macros
------

On Windows, as an alternative to calling macros via :ref:`RunPython <run_python>`, you can also use the ``@xw.sub``
decorator::

    import xlwings as xw

    @xw.sub
    def my_macro():
        """Writes the name of the Workbook into Range("A1") of Sheet 1"""
        wb = xw.Book.caller()
        wb.sheets[0].range('A1').value = wb.name

After clicking on ``Import Python UDFs``, you can then use this macro by executing it via ``Alt + F8`` or by
binding it e.g. to a button. To do the latter, make sure you have the ``Developer`` tab selected under ``File >
Options > Customize Ribbon``. Then, under the ``Developer`` tab, you can insert a button via ``Insert > Form Controls``.
After drawing the button, you will be prompted to assign a macro to it and you can select ``my_macro``.

.. _call_udfs_from_vba:

Call UDFs from VBA
------------------

Imported functions can also be used from VBA. For example, for a function returning a 2d array:

.. code-block:: vb.net

    Sub MySub()
    
    Dim arr() As Variant
    Dim i As Long, j As Long
    
        arr = my_imported_function(...)
        
        For j = LBound(arr, 2) To UBound(arr, 2)
            For i = LBound(arr, 1) To UBound(arr, 1)
                Debug.Print "(" & i & "," & j & ")", arr(i, j)
            Next i
        Next j
    
    End Sub


.. _async_functions:

Asynchronous UDFs
-----------------

.. note::
    This is an experimental feature

.. versionadded:: v0.14.0

xlwings offers an easy way to write asynchronous functions in Excel. Asynchronous functions return immediately with
``#N/A waiting...``. While the function is waiting for its return value, you can use Excel to do other stuff and whenever
the return value is available, the cell value will be updated.

The only available mode is currently ``async_mode='threading'``, meaning that it's useful for I/O-bound tasks, for example when
you fetch data from an API over the web.

You make a function asynchronous simply by giving it the respective argument in the function decorator. In this example,
the time consuming I/O-bound task is simulated by using ``time.sleep``::

    import xlwings as xw
    import time

    @xw.func(async_mode='threading')
    def myfunction(a):
        time.sleep(5)  # long running tasks
        return a



You can use this function like any other xlwings function, simply by putting ``=myfunction("abcd")`` into a cell
(after you have imported the function, off course).

Note that xlwings doesn't use the native asynchronous functions that were introduced with Excel 2010, so xlwings
asynchronous functions are supported with any version of Excel.