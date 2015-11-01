.. _datastructures:

Working with Data Structures
============================

Single Cells
------------
Single cells are returned either as ``float``, ``unicode``, ``None`` or ``datetime`` objects, depending on whether the
cell contains a number, a string, is empty or represents a date:

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
* 1d lists: Ranges that represent rows or columns in Excel are returned as simple lists:

  .. code-block:: python

    >>> wb = Workbook()
    >>> Range('A1').value = [[1],[2],[3],[4],[5]]  # Column orientation (nested list)
    >>> Range('A1:A5').value
    [1.0, 2.0, 3.0, 4.0, 5.0]
    >>> Range('A1').value = [1, 2, 3, 4, 5]
    >>> Range('A1:E1').value
    [1.0, 2.0, 3.0, 4.0, 5.0]

* 2d lists: If the row or column orientation has to be preserved, use the ``atleast_2d`` keyword. This will return the
  Ranges as nested lists ("2d lists"):

  .. code-block:: python

    >>> Range('A1:A5', atleast_2d=True).value
    [[1.0], [2.0], [3.0], [4.0], [5.0]]
    >>> Range('A1:E1', atleast_2d=True).value
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

The "table", "vertical" and "horizontal" properties
---------------------------------------------------

Continuing the sample from above, you can get the dimensions of Excel Ranges dynamically through the properties
``table``, ``vertical`` and ``horizontal``. All that's needed is the top left cell together with one of these
properties.

.. code-block:: python

    >>> Range('A10').table.value
    [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
    >>> Range('A10').horizontal.value
    ['Foo 1', 'Foo 2', 'Foo 3']
    >>> Range('A10').vertical.value
    ['Foo 1', 10.0]

.. note:: Using these properties together with a named Range as top left cell gives you an extremely flexible setup in
    Excel: You can move around the table and change it's size without having to adjust your code, e.g. by using
    something like ``Range('NamedRange').table.value``.

NumPy Arrays
------------

NumPy arrays work similar to nested lists. However, empty cells are represented by ``nan`` instead of
``None``. If you want to read or write in a Range as array, use the ``array`` method of the ``Range`` object:

.. code-block:: python

    >>> import numpy as np
    >>> wb = Workbook()
    >>> Range('A1').array().value = np.eye(5) # Range('A1').value = np.eye(5) also works and calls array() internally
    >>> Range('A1').table.array().value
    array([[ 1.,  0.,  0.,  0.,  0.],
           [ 0.,  1.,  0.,  0.,  0.],
           [ 0.,  0.,  1.,  0.,  0.],
           [ 0.,  0.,  0.,  1.,  0.],
           [ 0.,  0.,  0.,  0.,  1.]])

Pandas DataFrames and Series
----------------------------

Pandas DataFrames and Series are also easy to work with:

* Series:

  .. code-block:: python

    >>> import pandas as pd
    >>> import numpy as np
    >>> wb = Workbook()
    >>> s = pd.Series([1.1, 3.3, 5., np.nan, 6., 8.])
    >>> s
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    dtype: float64
    >>> Range('A1').value = s
    >>> data = Range('A1', asarray=True).table.value
    >>> pd.Series(data[:,1], index=data[:,0])
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    dtype: float64

* DataFrame:

  .. code-block:: python

    >>> wb = Workbook()
    >>> Range('A1').value = [['one', 'two'], [1.1, 2.2], [3.3, None]]
    >>> data = Range('A1').table.dataframe(index=False).value
    >>> df
       one  two
    0  1.1  2.2
    1  3.3  NaN
    >>> Range('A5').dataframe().value = df # Export per default both index and header
    >>> Range('A9').dataframe(index=False).value = df  # Control index and header
    >>> Range('A13').dataframe(index=False, header=False).value = df # Control index and header

.. note:: You only need to specify the top left cell when writing a list, a NumPy array or a Pandas
    DataFrame to Excel, e.g.: ``Range('A1').array().value = np.eye(10)``

.. note:: You do need to specify a full range when reading a NumPy array or a Pandas
    DataFrame to Excel, e.g.: ``arr = Range('A1:D10').array().value``.

* Formats:

  It is also possible to register "formats" of arrays and dataframes to simplify the calls to
  dataframe() and array(). This may help in avoiding the repetition of the different arguments
  in multiple part of your code if you often access (read or write) a dataframe with the same
  format.

  .. code-block:: python

    >>> wb = Workbook()
    >>> # register a format to access a DataFrame with some arguments defined
    >>> register_format("my format",
                        DataFrameAccessor(header=True, index=False, autotable=True))
    >>> df1 = Range('A1').format("my format").value
    >>> # is fully equivalent to
    >>> df2 = Range('A1').dataframe(header=True, index=False, autotable=True).value
    >>> df1 == df2

  It also allows to define functions that transform the resulting array/dataframe before writing to or after reading from Excel.

  .. code-block:: python

    >>> wb = Workbook()
    >>> # register a format to access a DataFrame and localise the index (ie add a timezone information)
    >>> register_format("my format",
                        DataFrameAccessor(header=True, index=False, autotable=True,
                                          on_result_read=lambda df:df.localize("Europe/Brussels"),
                                          ))
    >>> df1 = Range('A1').format("my format").value
    >>> # is fully equivalent to
    >>> df2 = Range('A1').dataframe(header=True, index=False, autotable=True).value
    >>> df2 = df2.localize("Europe/Brussels")
    >>> df1 == df2

  It is also possible to transform the data before passing it to the array/dataframe parser.

  .. code-block:: python

    >>> wb = Workbook()
    >>> # register a format to access a DataFrame but, before letting pandas transform the data,
    >>> # replace the value of some cell (for instance, if two columns have the same name, this may
    >>> # cause incorrect interpretation by pandas)
    >>> def preprocess_data(data):
    >>>   data[0][1] = "my new column name"
    >>>   return data
    >>> register_format("my format",
                        DataFrameAccessor(header=True, index=True, autotable=True,
                                          on_data_read=preprocess_data,
                                          ))
    >>> df1 = Range('A1').format("my format").value
    >>> df1.columns[0]=="my new column name"

  Finally, the DataFrameAccessor can be generated from a Range

  .. code-block:: python

    >>> wb = Workbook()
    >>> # register a format to access a DataFrame directly from a Range.dataframe
    >>> register_format("my format", Range("A1").dataframe(index=True))
