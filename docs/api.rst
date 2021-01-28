.. _api:

Python API
==========

Top-level functions
-------------------

.. automodule:: xlwings
    :members: view, load

Object model
------------

.. _python_apps:

Apps
****

.. autoclass:: xlwings.main.Apps
    :members:

App
***

.. autoclass:: App
    :members:

.. _python_books:

Books
*****

.. autoclass:: xlwings.main.Books
    :members:

.. _python_book:

Book
****

.. autoclass:: Book
    :members:

Sheets
******

.. autoclass:: xlwings.main.Sheets
    :members:

Sheet
*****

.. autoclass:: Sheet
    :members:

Range
*****

.. autoclass:: Range
    :members:


RangeRows
*********

.. autoclass:: RangeRows
    :members:


RangeColumns
************

.. autoclass:: RangeColumns
    :members:


Shapes
******

.. autoclass:: xlwings.main.Shapes
    :members:
    :inherited-members:

Shape
*****

.. autoclass:: Shape
    :members:

Charts
******

.. autoclass:: xlwings.main.Charts
    :members:
    :inherited-members:

Chart
*****

.. autoclass:: Chart
    :members:

Pictures
********

.. autoclass:: xlwings.main.Pictures
    :members:
    :inherited-members:

Picture
*******

.. autoclass:: Picture
    :members:

Names
*****

.. autoclass:: xlwings.main.Names
    :members:

Name
****

.. autoclass:: Name
    :members:

Tables
******

.. autoclass:: xlwings.main.Tables
    :members:

Table
*****

.. autoclass:: xlwings.main.Table
    :members:

.. _udf_api:

UDF decorators
--------------


.. py:function:: xlwings.func(category="xlwings", volatile=False, call_in_wizard=True)

    Functions decorated with ``xlwings.func`` will be imported as ``Function`` to Excel when running
    "Import Python UDFs".

    Arguments
    ---------

    category : int or str, default "xlwings"
        1-14 represent built-in categories, for user-defined categories use strings

        .. versionadded:: 0.10.3

    volatile : bool, default False
        Marks a user-defined function as volatile. A volatile function must be recalculated
        whenever calculation occurs in any cells on the worksheet. A nonvolatile function is
        recalculated only when the input variables change. This method has no effect if it's
        not inside a user-defined function used to calculate a worksheet cell.

        .. versionadded:: 0.10.3

    call_in_wizard : bool, default True
        Set to False to suppress the function call in the function wizard.

        .. versionadded:: 0.10.3

.. py:function:: xlwings.sub()

    Functions decorated with ``xlwings.sub`` will be imported as ``Sub`` (i.e. macro) to Excel when running
    "Import Python UDFs".

.. py:function:: xlwings.arg(arg, convert=None, **options)

    Apply converters and options to arguments, see also :meth:`Range.options`.


    **Examples:**

    Convert ``x`` into a 2-dimensional numpy array:

    .. code-block:: python

        import xlwings as xw
        import numpy as np

        @xw.func
        @xw.arg('x', np.array, ndim=2)
        def add_one(x):
            return x + 1


.. py:function:: xlwings.ret(convert=None, **options)

    Apply converters and options to return values, see also :meth:`Range.options`.

    **Examples**

    1) Suppress the index and header of a returned DataFrame:

    .. code-block:: python

        import pandas as pd

        @xw.func
        @xw.ret(index=False, header=False)
        def get_dataframe(n, m):
            return pd.DataFrame(np.arange(n * m).reshape((n, m)))


    2) Dynamic array:

    .. note::
        If your version of Excel supports the new native dynamic arrays, then you don't have to do anything special,
        and you shouldn't use the ``expand`` decorator! To check if your version of Excel supports it, see if you
        have the ``=UNIQUE()`` formula available. Native dynamic arrays were introduced in Office 365 Insider Fast
        at the end of September 2018.

    ``expand='table'`` turns the UDF into a dynamic array. Currently you must not use volatile functions
    as arguments of a dynamic array, e.g. you cannot use ``=TODAY()`` as part of a dynamic array. Also
    note that a dynamic array needs an empty row and column at the bottom and to the right and will overwrite
    existing data without warning.

    Unlike standard Excel arrays, dynamic arrays are being used from a single cell like a standard function
    and auto-expand depending on the dimensions of the returned array:

    .. code-block:: python

        import xlwings as xw
        import numpy as np

        @xw.func
        @xw.ret(expand='table')
        def dynamic_array(n, m):
            return np.arange(n * m).reshape((n, m))


    .. versionadded:: 0.10.0

.. _reports_api:

Reports
-------

.. automodule:: xlwings.pro.reports
    :members: