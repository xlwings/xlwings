Quickstart
==========

This guide assumes you have xlwings already installed. If that's not the case, head over to :ref:`installation`.

Interact with Excel from Python
-------------------------------

Writing/reading values to/from Excel and adding a chart is as easy as:

.. code-block:: python

    >>> import xlwings as xw
    >>> wb = xw.Workbook()  # Creates a connection with a new workbook
    >>> xw.Range('A1').value = 'Foo 1'
    >>> xw.Range('A1').value
    'Foo 1'
    >>> xw.Range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> xw.Range('A1').table.value  # or: Range('A1:C2').value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> xw.Sheet(1).name
    'Sheet1'
    >>> chart = xw.Chart.add(source_data=xw.Range('A1').table)

The Range and Chart objects as used above will refer to the active sheet of the current Workbook ``wb``. Include the
Sheet name like this:

.. code-block:: python

    xw.Range('Sheet1', 'A1:C3').value
    xw.Range(1, (1,1), (3,3)).value  # index notation
    xw.Chart.add('Sheet1', source_data=xw.Range('Sheet1', 'A1').table)

Qualify the Workbook additionally like this:

.. code-block:: python

    xw.Range('Sheet1', 'A1', wkb=wb).value
    xw.Chart.add('Sheet1', wkb=wb, source_data=xw.Range('Sheet1', 'A1', wkb=wb).table)
    xw.Sheet(1, wkb=wb).name

or simply set the current workbook first:

.. code-block:: python

    wb.set_current()
    xw.Range('Sheet1', 'A1').value
    xw.Chart.add('Sheet1', source_data=xw.Range('Sheet1', 'A1').table)
    xw.Sheet(1).name

These commands also work seamlessly with **NumPy arrays** and **Pandas DataFrames**, see :ref:`datastructures` for details.

**Matplotlib** figures can be shown as pictures in Excel:

.. code-block:: python

    import matplotlib.pyplot as plt
    fig = plt.figure()
    plt.plot([1, 2, 3, 4, 5])

    plot = xw.Plot(fig)
    plot.show('Plot1')

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
    variable.

For further details, see :ref:`vba`.

User Defined Functions (UDFs) - Currently Windows only
------------------------------------------------------

Writing a UDF in Python is as easy as:

.. code-block:: python

    import xlwings as xw

    @xw.func
    def double_sum(x, y):
        """Returns twice the sum of the two arguments"""
        return 2 * (x + y)

This then needs to be imported into Excel: For further details, see :ref:`udfs`.

Easy deployment
---------------

Deployment is really the part where xlwings shines:

* Just zip-up your Spreadsheet with your Python code and send it around. The receiver only needs to have an
  installation of Python with xlwings (and obviously all the other packages you're using).
* There is no need to install any Excel add-in.



