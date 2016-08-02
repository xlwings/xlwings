Quickstart
==========

This guide assumes you have xlwings already installed. If that's not the case, head over to :ref:`installation`.

1. Scripting: Automate/interact with Excel from Python
------------------------------------------------------

Reading/writing values to/from the **active sheet** is as easy as:

    >>> import xlwings as xw
    >>> xw.Range('A1').value = 'Foo 1'
    >>> xw.Range('A1').value
    'Foo 1'

There are many **convenience features** available, e.g. Range expanding:

    >>> xw.Range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> xw.Range('A1').expand().value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]

**Powerful converters** handle most data types of interest, including Numpy arrays and Pandas DataFrames in both directions:

    >>> import pandas as pd
    >>> df = pd.DataFrame([[1,2], [3,4]], columns=['a', 'b'])
    >>> xw.Range('A1').value = df
    >>> xw.Range('A1').options(pd.DataFrame, expand='table').value
           a    b
    0.0  1.0  2.0
    1.0  3.0  4.0

**Full qualification**: Instantiate a new book, add a new sheet and write a value to a specific sheet:

    >>> wb = xw.Book()
    >>> wb.sheets.add()
    <Sheet [Workbook1]Sheet2>
    >>> wb.sheets['Sheet1'].range('A1').value = 'Foo1'

Usually, you can just use ``xw.Book`` and it finds your workbook over all instances of Excel:

    >>> xw.Book('FileName.xlsx')

If you need more control (e.g. you have the same file open in two Excel instance), you'll need to fully qualify it like so:

    >>> xw.apps[0].books['FileName.xlsx']


**Matplotlib** figures can be shown as pictures in Excel:

    >>> import matplotlib.pyplot as plt
    >>> fig = plt.figure()
    >>> plt.plot([1, 2, 3, 4, 5])
    [<matplotlib.lines.Line2D at 0x1071706a0>]
    >>> wb = xw.Book()
    >>> wb.sheets[0].pictures.add(fig, name='MyPlot', update=True)
    <Picture 'MyPlot' in <Sheet [Workbook4]Sheet1>>

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
