Quickstart
==========

This guide assumes you have xlwings already installed. If that's not the case, head over to :ref:`installation`.

1. Scripting: Automate/interact with Excel from Python
------------------------------------------------------

Establish a connection to a workbook:

    >>> import xlwings as xw
    >>> wb = xw.Book()  # this will create a new workbook
    >>> wb = xw.Book('FileName.xlsx')  # connect to an existing file in the current working directory
    >>> wb = xw.Book(r'C:\path\to\file.xlsx')  # on Windows: use raw strings to escape backslashes

If you have the same file open in two instances of Excel, you need to fully qualify it and include the app instance:

    >>> xw.apps[0].books['FileName.xlsx']

Instantiate a sheet object:

    >>> sht = wb.sheets['Sheet1']

Reading/writing values to/from ranges is as easy as:

    >>> sht.range('A1').value = 'Foo 1'
    >>> sht.range('A1').value
    'Foo 1'

There are many **convenience features** available, e.g. Range expanding:

    >>> sht.range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> sht.range('A1').expand().value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]

**Powerful converters** handle most data types of interest, including Numpy arrays and Pandas DataFrames in both directions:

    >>> import pandas as pd
    >>> df = pd.DataFrame([[1,2], [3,4]], columns=['a', 'b'])
    >>> sht.range('A1').value = df
    >>> sht.range('A1').options(pd.DataFrame, expand='table').value
           a    b
    0.0  1.0  2.0
    1.0  3.0  4.0

**Matplotlib** figures can be shown as pictures in Excel:

    >>> import matplotlib.pyplot as plt
    >>> fig = plt.figure()
    >>> plt.plot([1, 2, 3, 4, 5])
    [<matplotlib.lines.Line2D at 0x1071706a0>]
    >>> sht.pictures.add(fig, name='MyPlot', update=True)
    <Picture 'MyPlot' in <Sheet [Workbook4]Sheet1>>

**Shortcut** for the active sheet: ``xw.Range``

If you want to quickly talk to the active sheet in the active workbook, you don't need instantiate a workbook
and sheet object, but can simply do:

    >>> import xlwings xw
    >>> xw.Range('A1').value = 'Foo'
    >>> xw.Range('A1').value
    'Foo'

**Note:** You should only use ``xw.Range`` when interacting with Excel. In scripts, you should always
go via book and sheet objects as shown above.

2. Macros: Call Python from Excel
---------------------------------

You can call Python functions from VBA using the ``RunPython`` function:

.. code-block:: vb.net

    Sub HelloWorld()
        RunPython ("import hello; hello.world()")
    End Sub

Per default, ``RunPython`` expects ``hello.py`` in the same directory as the Excel file. Refer to the calling Excel
book by using ``xw.Book.caller``:

.. code-block:: python

    # hello.py
    import numpy as np
    import xlwings as xw

    def world():
        wb = xw.Book.caller()
        wb.sheets[0].range('A1').value = 'Hello World!'


To make this run, you'll need to have the xlwings VBA module in your Excel book. The easiest way to get everything set
up is to use the xlwings command line client from either a command prompt on Windows or a terminal on Mac: ``xlwings quickstart myproject``.

To import the xlwings VBA module differently, and for more details, see :ref:`vba`.

3. UDFs: User Defined Functions (Windows only)
----------------------------------------------

Writing a UDF in Python is as easy as:

.. code-block:: python

    import xlwings as xw

    @xw.func
    def hello(name):
        return 'Hello {0}'.format(name)

Converters can be used with UDFs, too. Again a Pandas DataFrame example:


.. code-block:: python

    import xlwings as xw
    import pandas as pd

    @xw.func
    @xw.arg('x', pd.DataFrame)
    def correl2(x):
        # x arrives as DataFrame
        return x.corr()

Import this function into Excel by clicking the import button of the xlwings add-in: For further details, see :ref:`udfs`.
