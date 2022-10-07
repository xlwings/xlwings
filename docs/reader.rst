.. _file_reader:

Excel File Reader :bdg-secondary:`PRO`
=====================================

This feature requires at least v0.28.0.

xlwings PRO comes with an ultra fast Excel file reader. Compared with ``pandas.read_excel()``, you should be able to see speedups anywhere between 5 to 25 times when reading a single sheet. The exact speed will depend on your content, file format, and Python version. The following Excel file formats are supported:

* ``xlsx`` / ``xlsm`` / ``xlam``
* ``xlsb``
* ``xls``

Other advantages include:

* Support for named ranges
* Support for dynamic ranges via ``myrange.expand()`` or ``myrange.options(expand="table")``, respectively.
* Support for converters so you can read in ranges not just as pandas DataFrames, but also as NumPy arrays, lists, scalar values, dictionaries, etc.
* You can read out cell errors like ``#DIV/0!`` or ``#N/A`` as strings instead of converting them all into ``NaN``

Unlike the classic ("interactive") use of xlwings that requires Excel to be installed, reading a file doesn't depend on an installation of Excel and therefore works everywhere where Python runs. However, reading directly from a file requires the workbook to be saved before xlwings is able to pick up any changes.

Reading a specific range
------------------------
To open a file in read mode, provide the ``mode="r"`` argument: ``xw.Book("myfile.xlsx", mode="r")``. You usually want to use ``Book`` as a context manager so that the file is automatically closed and resources cleaned up once the code leaves the body of the ``with`` statement:

.. code-block:: python

    import xlwings as xw

    with xw.Book("myfile.xlsx", mode="r") as book:
        sheet1 = book.sheets[0]
        data = sheet1["A1:B2"].value

If you don't use the ``with`` statement, make sure to close the book manually via ``book.close()``.

Reading an entire sheet
-----------------------

To read an entire sheet, use the ``cells`` property:

.. code-block:: python

    with xw.Book("myfile.xlsx", mode="r") as book:
        sheet1 = book.sheets[0]
        data = sheet1.cells.value

Converters: DataFrames etc.
---------------------------

You can use the usual converters, for example to read in a range as a DataFrame:

.. code-block:: python

    with xw.Book("myfile.xlsx", mode="r") as book:
        sheet1 = book.sheets[0]
        df = sheet1["A1:B2"].options("df").value
        # As usual, you can also provide more options
        df = sheet1["A1:B2"].options("df", index=False).value

For more details, see :ref:`converters`.

Named Ranges
------------

Named ranges can be accessed like so:

.. code-block:: python

    with xw.Book("myfile.xlsx", mode="r") as book:
        sheet1 = book.sheets[0]
        data = sheet1["myname"].value  # get values
        address = sheet1["myname"].address  # get address

Alternatively, you can also access them via the :meth:`Names <xlwings.main.Names>` collection:

.. code-block:: python

    with xw.Book("myfile.xlsx", mode="r") as book:
        for name in book.names:
            print(name.refers_to_range.value)

Dynamic Ranges
--------------

You can make use of the usual range expansion to read in a range of dynamic size:

.. code-block:: python

    with xw.Book("myfile.xlsx", mode="r") as book:
        sheet1 = book.sheets[0]
        data = sheet1["A1"].expand().value

Cell errors
-----------

While xlwings reads in cell errors such as ``#N/A`` as ``None`` by default, you may want to read them in as strings if you're specifically looking for these by using the ``err_to_str`` option:

.. code-block:: python

    with xw.Book("myfile.xlsx", mode="r") as book:
        sheet1 = book.sheets[0]
        data = sheet1["A1:B2"].option(err_to_str=True).value


Limitations
-----------
* The reader is currently only available via ``pip install xlwings``. Installation via ``conda`` is not yet supported, but you can still use pip to install xlwings into a Conda environment!
* Date cells: Excel cells with a Date/Time are currently only converted to a ``datetime`` object in Python for ``xlsx`` file formats. For ``xlsb`` format, pandas has the same restriction though (it uses ``pyxlsb`` under the hood).
* Dynamic ranges: ``myrange.expand()`` is currently inefficient, so will slow down the reading considerably if the dynamic range is big.
* Named ranges: Named ranges with sheet scope are currently not shown with their proper name: E.g. ``mybook.names[0].name`` will show the name ``mylocalname`` instead of including the sheet name like so ``Sheet1!mylocalname``. Along the same lines, the ``names`` property can only be accessed via ``book`` object, not via ``sheet`` object.
* Excel tables: Accessing data via table names isn't supported at the moment.
* Options: except for ``err_to_str``, non-default options are currently inefficient and will slow down the read operation. This includes ``dates``, ``empty``, and ``numbers``.
* Formulas: currently only the cell values are supported, but not the cell formulas.
* This is only a file reader, writing files is currently not supported.