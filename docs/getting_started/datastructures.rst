.. _datastructures:

Data Structures Tutorial
========================

This tutorial gives you a quick introduction to the most common use cases and default behaviour of xlwings when reading
and writing values. For an in-depth documentation of how to control the behavior using the ``options`` method, have a
look at :ref:`converters`.

All code samples below depend on the following import:

    >>> import xlwings as xw

Single Cells
------------
Single cells are by default returned either as ``float``, ``unicode``, ``None`` or ``datetime`` objects, depending on
whether the cell contains a number, a string, is empty or represents a date:

.. code-block:: python

    >>> import datetime as dt
    >>> sheet = xw.Book().sheets[0]
    >>> sheet['A1'].value = 1
    >>> sheet['A1'].value
    1.0
    >>> sheet['A2'].value = 'Hello'
    >>> sheet['A2'].value
    'Hello'
    >>> sheet['A3'].value is None
    True
    >>> sheet['A4'].value = dt.datetime(2000, 1, 1)
    >>> sheet['A4'].value
    datetime.datetime(2000, 1, 1, 0, 0)

Lists
-----
* 1d lists: Ranges that represent rows or columns in Excel are returned as simple lists, which means that once
  they are in Python, you've lost the information about the orientation. If that is an issue, the next point shows
  you how to preserve this info:

  .. code-block:: python

    >>> sheet = xw.Book().sheets[0]
    >>> sheet['A1'].value = [[1],[2],[3],[4],[5]]  # Column orientation (nested list)
    >>> sheet['A1:A5'].value
    [1.0, 2.0, 3.0, 4.0, 5.0]
    >>> sheet['A1'].value = [1, 2, 3, 4, 5]
    >>> sheet['A1:E1'].value
    [1.0, 2.0, 3.0, 4.0, 5.0]

  To force a single cell to arrive as list, use::

    >>> sheet['A1'].options(ndim=1).value
    [1.0]

  .. note::
    To write a list in column orientation to Excel, use ``transpose``: ``sheet.range('A1').options(transpose=True).value = [1,2,3,4]``

* 2d lists: If the row or column orientation has to be preserved, set ``ndim`` in the Range options. This will return the
  Ranges as nested lists ("2d lists"):

  .. code-block:: python

    >>> sheet['A1:A5'].options(ndim=2).value
    [[1.0], [2.0], [3.0], [4.0], [5.0]]
    >>> sheet['A1:E1'].options(ndim=2).value
    [[1.0, 2.0, 3.0, 4.0, 5.0]]


* 2 dimensional Ranges are automatically returned as nested lists. When assigning (nested) lists to a Range in Excel,
  it's enough to just specify the top left cell as target address. This sample also makes use of index notation to read the
  values back into Python:

  .. code-block:: python

    >>> sheet['A10'].value = [['Foo 1', 'Foo 2', 'Foo 3'], [10, 20, 30]]
    >>> sheet.range((10,1),(11,3)).value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]


.. note:: Try to minimize the number of interactions with Excel. It is always more efficient to do
    ``sheet.range('A1').value = [[1,2],[3,4]]`` than ``sheet.range('A1').value = [1, 2]`` and ``sheet.range('A2').value = [3, 4]``.

Range expanding
---------------

You can get the dimensions of Excel Ranges dynamically through either the method ``expand`` or through the ``expand``
keyword in the ``options`` method. While ``expand`` gives back an expanded Range object, options are only evaluated when
accessing the values of a Range. The difference is best explained with an example:

.. code-block:: python

    >>> sheet = xw.Book().sheets[0]
    >>> sheet['A1'].value = [[1,2], [3,4]]
    >>> range1 = sheet['A1'].expand('table')  # or just .expand()
    >>> range2 = sheet['A1'].options(expand='table')
    >>> range1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> range2.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> sheet['A3'].value = [5, 6]
    >>> range1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> range2.value
    [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]

``'table'`` expands to ``'down'`` and ``'right'``, the other available options which can be used for column or row only
expansion, respectively.

.. note:: Using ``expand()`` together with a named Range as top left cell gives you a flexible setup in
    Excel: You can move around the table and change its size without having to adjust your code, e.g. by using
    something like ``sheet.range('NamedRange').expand().value``.

NumPy arrays
------------

NumPy arrays work similar to nested lists. However, empty cells are represented by ``nan`` instead of
``None``. If you want to read in a Range as array, set ``convert=np.array`` in the ``options`` method:

.. code-block:: python

    >>> import numpy as np
    >>> sheet = xw.Book().sheets[0]
    >>> sheet['A1'].value = np.eye(3)
    >>> sheet['A1'].options(np.array, expand='table').value
    array([[ 1.,  0.,  0.],
           [ 0.,  1.,  0.],
           [ 0.,  0.,  1.]])

Pandas DataFrames
-----------------

.. code-block:: python

    >>> sheet = xw.Book().sheets[0]
    >>> df = pd.DataFrame([[1.1, 2.2], [3.3, None]], columns=['one', 'two'])
    >>> df
       one  two
    0  1.1  2.2
    1  3.3  NaN
    >>> sheet['A1'].value = df
    >>> sheet['A1:C3'].options(pd.DataFrame).value
       one  two
    0  1.1  2.2
    1  3.3  NaN
    # options: work for reading and writing
    >>> sheet['A5'].options(index=False).value = df
    >>> sheet['A9'].options(index=False, header=False).value = df

Pandas Series
-------------

.. code-block:: python

    >>> import pandas as pd
    >>> import numpy as np
    >>> sheet = xw.Book().sheets[0]
    >>> s = pd.Series([1.1, 3.3, 5., np.nan, 6., 8.], name='myseries')
    >>> s
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    Name: myseries, dtype: float64
    >>> sheet['A1'].value = s
    >>> sheet['A1:B7'].options(pd.Series).value
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    Name: myseries, dtype: float64

.. note:: You only need to specify the top left cell when writing a list, a NumPy array or a Pandas
    DataFrame to Excel, e.g.: ``sheet['A1'].value = np.eye(10)``

Chunking: Read/Write big DataFrames etc.
----------------------------------------

When you read and write from or to big ranges, you may have to chunk them or you will hit a timeout or a memory error. The ideal ``chunksize`` will depend on your system and size of the array, so you will have to try out a few different chunksizes to find one that works well:

.. code-block:: python

    import pandas as pd
    import numpy as np
    sheet = xw.Book().sheets[0]
    data = np.arange(75_000 * 20).reshape(75_000, 20)
    df = pd.DataFrame(data=data)
    sheet['A1'].options(chunksize=10_000).value = df
        
And the same for reading:

.. code-block:: python

    # As DataFrame
    df = sheet['A1'].expand().options(pd.DataFrame, chunksize=10_000).value
    # As list of list
    df = sheet['A1'].expand().options(chunksize=10_000).value
