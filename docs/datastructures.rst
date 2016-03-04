.. _datastructures:

Data Structures Tutorial
========================

This page gives you a quick introduction to the most common use cases and default behaviour of xlwings when reading and
writing values. For an in-depth documentation of how to control things using the ``options`` method, have a look at :ref:`converters`.

Single Cells
------------
Single cells are by default returned either as ``float``, ``unicode``, ``None`` or ``datetime`` objects, depending on
whether the cell contains a number, a string, is empty or represents a date:

.. code-block:: python

    >>> from xlwings import Workbook, Range
    >>> from datetime import datetime
    >>> wb = Workbook()
    >>> Range('A1').value = 1
    >>> Range('A1').value
    1.0
    >>> Range('A2').value = 'Hello'
    >>> Range('A2').value
    'Hello'
    >>> Range('A3').value is None
    True
    >>> Range('A4').value = datetime(2000, 1, 1)
    >>> Range('A4').value
    datetime.datetime(2000, 1, 1, 0, 0)

Lists
-----
* 1d lists: Ranges that represent rows or columns in Excel are returned as simple lists, which means that once
  they are in Python, you've lost the information about the orientation. If that is an issue, read under the next
  point how this information can be preserved:

  .. code-block:: python

    >>> wb = Workbook()
    >>> Range('A1').value = [[1],[2],[3],[4],[5]]  # Column orientation (nested list)
    >>> Range('A1:A5').value
    [1.0, 2.0, 3.0, 4.0, 5.0]
    >>> Range('A1').value = [1, 2, 3, 4, 5]
    >>> Range('A1:E1').value
    [1.0, 2.0, 3.0, 4.0, 5.0]

  To force a single cell to arrive as list, use::

    >>> Range('A1').options(ndim=1).value
    [1.0]

  .. note::
    To write a list in column orientation to Excel, use ``transpose``: ``Range('A1').options(transpose=True).value = [1,2,3,4]``

* 2d lists: If the row or column orientation has to be preserved, set ``ndim`` in the Range options. This will return the
  Ranges as nested lists ("2d lists"):

  .. code-block:: python

    >>> Range('A1:A5').options(ndim=2).value
    [[1.0], [2.0], [3.0], [4.0], [5.0]]
    >>> Range('A1:E1').options(ndim=2).value
    [[1.0, 2.0, 3.0, 4.0, 5.0]]


* 2 dimensional Ranges are automatically returned as nested lists. When assigning (nested) lists to a Range in Excel,
  it's enough to just specify the top left cell as target address. This sample also makes use of index notation to read the
  values back into Python:

  .. code-block:: python

    >>> Range('A10').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10, 20, 30]]
    >>> Range((10,1),(11,3)).value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]


.. note:: Try to minimize the number of interactions with Excel. It is always more efficient to do
    ``Range('A1').value = [[1,2],[3,4]]`` than ``Range('A1').value = [1, 2]`` and ``Range('A2').value = [3, 4]``.

Range expanding: "table", "vertical" and "horizontal"
-----------------------------------------------------

You can get the dimensions of Excel Ranges dynamically through either the Range
properties ``table``, ``vertical`` and ``horizontal`` or through ``options(expand='table')``
(same for ``'vertical'`` and ``'horizontal'``). While properties give back a changed Range object,
options are only evaluated when accessing the values of a Range. The difference is best explained with an example:

.. code-block:: python

    >>> wb = Workbook()
    >>> Range('A1').value = [[1,2], [3,4]]
    >>> rng1 = Range('A1').table
    >>> rng2 = Range('A1').options(expand='table')
    >>> rng1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> rng2.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> Range('A3').value = [5, 6]
    >>> rng1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> rng2.value
    [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]

.. note:: Using ``table`` together with a named Range as top left cell gives you a flexible setup in
    Excel: You can move around the table and change it's size without having to adjust your code, e.g. by using
    something like ``Range('NamedRange').table.value``.

NumPy arrays
------------

NumPy arrays work similar to nested lists. However, empty cells are represented by ``nan`` instead of
``None``. If you want to read in a Range as array, set ``convert=np.array``:

.. code-block:: python

    >>> import numpy as np
    >>> wb = Workbook()
    >>> Range('A1').value = np.eye(3)
    >>> Range('A1').options(np.array, expand='table').value
    array([[ 1.,  0.,  0.],
           [ 0.,  1.,  0.],
           [ 0.,  0.,  1.]])

Pandas DataFrames
-----------------

.. code-block:: python

    >>> wb = Workbook()
    >>> df = pd.DataFrame([[1.1, 2.2], [3.3, None]], columns=['one', 'two'])
    >>> df
       one  two
    0  1.1  2.2
    1  3.3  NaN
    >>> Range('A1').value = df
    >>> Range('A1:C3').options(pd.DataFrame).value
       one  two
    0  1.1  2.2
    1  3.3  NaN
    # options: work for reading and writing
    >>> Range('A5').options(index=False).value = df
    >>> Range('A9').options(index=False, header=False).value = df

Pandas Series
-------------

.. code-block:: python

    >>> import pandas as pd
    >>> import numpy as np
    >>> wb = Workbook()
    >>> s = pd.Series([1.1, 3.3, 5., np.nan, 6., 8.], name='myseries')
    >>> s
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    Name: myseries, dtype: float64
    >>> Range('A1').value = s
    >>> Range('A1:B7').options(pd.Series).value
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    Name: myseries, dtype: float64

.. note:: You only need to specify the top left cell when writing a list, an NumPy array or a Pandas
    DataFrame to Excel, e.g.: ``Range('A1').value = np.eye(10)``