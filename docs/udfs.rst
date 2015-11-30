.. _udfs:

User Defined Functions (UDFs)
=============================

.. note:: This functionality is currently available on Windows only.

Installation of Excel Add-in (Recommended)
------------------------------------------

It is recommended (although not required) to work with the xlwings developer add-in to import the functions:

1) Install the add-in on a command prompt: ``xlwings addin install`` (see :ref:`command_line`)
2) Enable ``Trust access to the VBA project object model`` under ``File > Options > Trust Center > Trust Center Settings > Macro Settings``

Workbook preparation
--------------------

1) Create a new Excel file from the template: ``xlwings template open`` (see :ref:`command_line`) or just import
   the xlwings VBA module (``xlwings.bas``) manually, see :ref:`vba`.
2) Save the Workbook as ``Excel Macro-Enabled Workbook (*.xlsm)``.


Simple User-Defined Functions
-----------------------------

The default settings (see :ref:`VBA settings <vba_settings>`) expect a Python source file:

* in the same directory as the Excel file
* with the same name as the Excel file, but with a ``.py`` ending instead of ``.xlsm``.

Alternatively, you can point to a specific source file by setting the ``UDF_PATH`` in the VBA settings.

Let's assume you've got a Workbook ``Book1.xlsm``, then you would create a file ``Book1.py`` in the same directory with
the following sample function::

    from xlwings import xlfunc, xlarg

    @xlfunc
    def double_sum(x, y):
        """Returns twice the sum of the two arguments"""
        return 2 * (x + y)


* Now click on ``Import Python UDFs`` in the xlwings tab to pick up the changes made to ``Book1.py``. If you don't
  want to install/use the add-in, you could also run the ``ImportPythonUDFs`` macro directly (one possibility to do that
  is to hit ``Alt + F8`` and select the macro from the pop-up menu).
* Enter the formula ``=double_sum(1, 2)`` into a cell and you will see the correct result:

  .. figure:: images/double_sum.png
    :scale: 80%

Note that the formula can be used in VBA, too.

Array Formulas I: without NumPy
-------------------------------

You can pass an Excel Range as a function argument, as opposed to a single cell and it will show up in Python as tuple of tuples.

For example, you can write the following function to add 1 to every cell in a Range::

    @xlfunc
    def add_one(data):
        return [[cell + 1 for cell in row] for row in data]

To use this formula in Excel,

* Click on ``Import Python UDFs`` again
* Fill in the values in ``Range("A1:B2")``
* Select ``Range("D1:E2")``
* Type in the formula ``=add_one("A1:B2")``
* Press ``Ctrl+Shift+Enter`` to create an array formula. If you did everything correctly, you'll see the formula
  surrounded by curly braces as in this screenshot:

.. figure:: images/array_formula.png
    :scale: 80%

Number of array dimensions: ndim
********************************

The above formula has the issue that it expects a "two dimensional" input, e.g. a nested list of the form
``[[1, 2], [3, 4]]``.
Therefore, if you would apply the formula to a single cell or a row/column, you would get the following error:
``TypeError: 'float' object is not iterable``.

To force Excel to always give you a two-dimensional array, you can extend the above formula like this::

    @xlfunc
    @xlarg('data', ndim=2)
    def add_one(data):
        return [[cell + 1 for cell in row] for row in data]

Now, you can use the formula with single cells, rows/columns and two-dimensional ranges.
Accordingly, you can use ``ndim=1`` to force a single cell to arrive as tuple.

Array Formulas II: with NumPy
-----------------------------

Most of the time, you'll want to use NumPy arrays as this unlocks the full power of Python's ecosystem for scientific computing.

To define a formula for matrix multiplication, you would define the following function::

    @xlfunc
    @xlarg('x', 'nparray', ndim=2)
    @xlarg('y', 'nparray', ndim=2)
    def matrix_mult(x, y):
        return x @ y

.. note:: If you are not on Python >= 3.5 with NumPy >= 1.10, use ``x.dot(y)`` instead of ``x @ y``.

Macros
------

On Windows, as alternative to calling macros via :ref:`RunPython <run_python>`, you can also use a decorator based
approach that works the same as with user-defined functions::

    from xlwings import Workbook, xlsub

    @xlsub
    def my_macro():
        """Writes the name of the Workbook into Range("A1") of Sheet 1"""
        wb = Workbook.caller()
        Range(1, 'A1').value = wb.name

After clicking on ``Import Python UDFs``, you can then use this macro by executing it via ``Alt + F8`` or by
binding it e.g. to a button. To to the latter, make sure you have the ``Developer`` tab selected under ``File >
Options > Customize Ribbon``. Then, under the ``Developer`` tab, you can insert a button via ``Insert > Form Controls``.
After drawing the button, you will be prompted to assign a macro to it and you can select ``my_macro``.