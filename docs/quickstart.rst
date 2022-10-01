.. _quickstart:

Quickstart
==========

This guide assumes you have xlwings already installed. If that's not the case, head over to :ref:`installation`.

1. Interacting with Excel from a Jupyter notebook
-------------------------------------------------

If you're just interested in getting a pandas DataFrame in and out of your Jupyter notebook, you can use the ``view`` and ``load`` functions, see  :ref:`jupyternotebooks`.

2. Scripting: Automate/interact with Excel from Python
------------------------------------------------------

Establish a connection to a workbook:

    >>> import xlwings as xw
    >>> wb = xw.Book()  # this will open a new workbook
    >>> wb = xw.Book('FileName.xlsx')  # connect to a file that is open or in the current working directory
    >>> wb = xw.Book(r'C:\path\to\file.xlsx')  # on Windows: use raw strings to escape backslashes

If you have the same file open in two instances of Excel, you need to fully qualify it and include the app instance.
You will find your app instance key (the PID) via ``xw.apps.keys()``:

    >>> xw.apps[10559].books['FileName.xlsx']

Instantiate a sheet object:

    >>> sheet = wb.sheets['Sheet1']

Reading/writing values to/from ranges is as easy as:

    >>> sheet['A1'].value = 'Foo 1'
    >>> sheet['A1'].value
    'Foo 1'

There are many **convenience features** available, e.g. Range expanding:

    >>> sheet['A1'].value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> sheet['A1'].expand().value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]

**Powerful converters** handle most data types of interest, including Numpy arrays and Pandas DataFrames in both directions:

    >>> import pandas as pd
    >>> df = pd.DataFrame([[1,2], [3,4]], columns=['a', 'b'])
    >>> sheet['A1'].value = df
    >>> sheet['A1'].options(pd.DataFrame, expand='table').value
           a    b
    0.0  1.0  2.0
    1.0  3.0  4.0

**Matplotlib** figures can be shown as pictures in Excel:

    >>> import matplotlib.pyplot as plt
    >>> fig = plt.figure()
    >>> plt.plot([1, 2, 3, 4, 5])
    [<matplotlib.lines.Line2D at 0x1071706a0>]
    >>> sheet.pictures.add(fig, name='MyPlot', update=True)
    <Picture 'MyPlot' in <Sheet [Workbook4]Sheet1>>

3. Macros: Call Python from Excel
---------------------------------

You can call Python functions either by clicking the ``Run`` button (new in v0.16) in  the add-in or from VBA using the ``RunPython`` function:

The ``Run`` button expects a function called ``main`` in a Python module with the same name as your workbook. The 
great thing about that approach is that you don't need your workbooks to be macro-enabled, you can save it as ``xlsx``.

If you want to call any Python function no matter in what module it lives or what name it has, use ``RunPython``:

.. code-block:: vb.net

    Sub HelloWorld()
        RunPython "import hello; hello.world()"
    End Sub


.. note::
    Per default, ``RunPython`` expects ``hello.py`` in the same directory as the Excel file with the same name, **but you can change both of these things**: if your Python file is an a different folder, add that folder to the ``PYTHONPATH`` in the config. If the file has a different name, change the ``RunPython`` command accordingly.

Refer to the calling Excel book by using ``xw.Book.caller()``:

.. code-block:: python

    # hello.py
    import numpy as np
    import xlwings as xw

    def world():
        wb = xw.Book.caller()
        wb.sheets[0]['A1'].value = 'Hello World!'


To make this run, you'll need to have the xlwings add-in installed or have the workbooks setup in the standalone mode. The easiest way to get everything set up is to use the xlwings command line client from either a command prompt on Windows or a terminal on Mac: ``xlwings quickstart myproject``.

For details about the addin, see :ref:`xlwings_addin`.

4. UDFs: User Defined Functions (Windows only)
----------------------------------------------

Writing a UDF in Python is as easy as:

.. code-block:: python

    import xlwings as xw

    @xw.func
    def hello(name):
        return f'Hello {name}'

Converters can be used with UDFs, too. Again a Pandas DataFrame example:


.. code-block:: python

    import xlwings as xw
    import pandas as pd

    @xw.func
    @xw.arg('x', pd.DataFrame)
    def correl2(x):
        # x arrives as DataFrame
        return x.corr()

Import this function into Excel by clicking the import button of the xlwings add-in: for a step-by-step tutorial, see :ref:`udfs`.
