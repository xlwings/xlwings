.. _converters:

Converters
==========

Introduced with v0.7.0, converters define how Excel Ranges and their values are being converted both during
**reading** and **writing** operations. They also provide a consistent experience across **xlwings.Range** objects and
**User Defined Functions** (UDFs).

Converters are explicitely set with the ``as_`` argument in the ``options`` method when manipulating ``xlwings.Range`` objects
or in the ``@xw.arg`` and ``@xw.ret`` decorators when using UDFs. If no converter is specified, the base converter
is applied when reading. When writing, xlwings will automatically apply the correct converter (if available) according to the
object's type that is being written to Excel. If no converter is found for that type, it falls back to the base converter.

**Syntax:**

==============================  ===========================================================  ===========
****                            **Range**                                                    **UDF**
==============================  ===========================================================  ===========
**reading**                     ``Range.options(as_=None, **kwargs).value``                  ``@arg('x', as_=None, **kwargs)``
**writing**                     ``Range.options(as_=None, **kwargs).value = myvalue``        ``@ret(as_=None, **kwargs)``
==============================  ===========================================================  ===========

.. note:: Keyword arguments (``kwargs``) may refer to the specific converter or the base converter.
  For example, to set the ``numbers`` option in the base converter and the ``index`` option in the DataFrame converter,
  you would write::

      Range('A1:C3').options(pd.DataFrame, index=False, numbers=int).value

Base Converter
--------------

If no options are set, the following conversions are performed:

* single cells are read in as ``floats`` in case the Excel cell holds a number, as ``unicode`` in case it holds text,
  as ``datetime`` if it contains a date and as ``None`` in case it is empty.
* columns/rows are read in as lists, e.g. ``[None, 1.0, 'a string']``
* 2d cell ranges are read in as list of lists, e.g. ``[[None, 1.0, 'a string'], [None, 2.0, 'another string']]``

The following options can be set:

* **ndim**

  Number of dimensions: This may be set to 1 or 2:

  >>> import xlwings as xw
  >>> wb = Workbook()
  >>> xw.Range('A1').value = [[1, 2], [3, 4]]
  >>> xw.Range('A1').value
  1.0
  >>> xw.Range('A1').options(ndim=1).value
  [1.0]
  >>> xw.Range('A1').options(ndim=2).value
  [[1.0]]
  >>> xw.Range('A1:A2').value
  [1.0 3.0]
  >>> xw.Range('A1:A2').options(ndim=2).value
  [[1.0], [3.0]]

* **numbers**

  The base converter reads in numbers as ``float``, you can change that like so::

    >>> xw.Range('A1').value = 1
    >>> xw.Range('A1').value
    1.0
    >>> xw.Range('A1').options(numbers=int).value
    1

  Using this on UDFs looks like this::

    @xw.func
    @xw.arg('x', numbers=int)
    def myfunction(x):
        # all numbers in x arrive as int
        return x

  Note that this option can only be used for reading, as Excel always stores numbers internally as floats.

* **dates**

  Cells with dates are converted per default into ``datetime.datetime``, you can change it to ``datetime.date``:

  - Range::

    >>> import datetime as dt
    >>> xw.Range('A1').options(dates=dt.date).value

  - UDFs: ``@xw.arg('x', dates=dt.date)``

* **empty**

  Empty cells are converted per default into ``None``, you can change this as follows:

  - Range: ``>>> xw.Range('A1').options(empty='NA').value``

  - UDFs:  ``@xw.arg('x', empty='NA')``

* **transpose**

  This works for reading and writing and allows us to e.g. write a list in column orientation to Excel:

  - Range: ``Range('A1').options(transpose=True).value = [1, 2, 3]``

  - UDFs:

    .. code-block:: python

        @xw.arg('x', transpose=True)
        @xw.ret(transpose=True)
        def myfunction(x):
            # x will be returned unchanged as transposed both when reading and writing
            return x

* **expand**

  This works the same as the Range properties ``table``, ``vertical`` and ``horizontal`` but is
  only evaluated when getting the values of a Range::

    >>> import xlwings as xw
    >>> wb = xw.Workbook()
    >>> xw.Range('A1').value = [[1,2], [3,4]]
    >>> rng1 = xw.Range('A1').table
    >>> rng2 = xw.Range('A1').options(expand='table')
    >>> rng1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> rng2.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> xw.Range('A3').value = [5, 6]
    >>> rng1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> rng2.value
    [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]

  .. note:: The ``expand`` option is only available on ``Range`` objects as UDFs only allow to manipulate the calling cells.

Built-in converters
-------------------

xlwings offers several built-in converters that perform additional conversions on top of the base converters for
**dictionaries**, **NumPy arrays**, **Pandas Series** and **DataFrames**. New, customized converters can also be
added (docs will follow).
Again, the samples below may be used with both ``xlwings.Range`` objects and UDFs, but the samples may only show one
version.

Dictionary converter
********************

The dictionary converter turns two Excel columns into a dictionary. If the data is in row orientation, use ``transpose``:

.. figure:: images/dict_converter.png
    :scale: 80%

::

    >>> Range('A1:B2').options(dict).value
    {'a': 1.0, 'b': 2.0}
    >>> Range('A4:B5').options(dict, transpose=True).value
    {'a': 1.0, 'b': 2.0}

Numpy array converter
*********************

**options:** ``dtype=None, copy=True, order=None, ndim=None``

The first 3 options behave the same as when using ``np.array()`` directly. Also, ``ndim`` works the same as shown above
for lists (under base converter) and hence returns either numpy scalars, 1d arrays or 2d arrays.

**Example**::

    >>> import numpy as np
    >>> Range('A1').options(transpose=True).value = np.array([1, 2, 3])
    >>> xw.Range('A1:A3').options(np.array, ndim=2).value
    array([[ 1.],
           [ 2.],
           [ 3.]])

Pandas series converter
***********************

**options:** ``dtype=None, copy=False, index=1, header=True``

The first 2 options behave the same as when using ``pd.Series()`` directly. ``ndim`` doesn't have an effect on
Pandas series as they are always expected and returned in column orientation.

``index``: int or Boolean
    | When reading, it expects the number of index columns shown in Excel.
    | When writing, include or exclude the index by setting it to ``True`` or ``False``.

``header``: Boolean
    | When reading, set it to ``False`` if Excel doesn't show either index or series names.
    | When writing, include or exclude the index and series names by setting it to ``True`` or ``False``.

For ``index`` and ``header``, ``1`` and ``True`` may be used interchangeably.

**Example:**

.. figure:: images/series_conv.png
    :scale: 80%

::

    >>> s = xw.Range('A1').options(pd.Series, expand='table').value
    >>> s
    date
    2001-01-01    1
    2001-01-02    2
    2001-01-03    3
    2001-01-04    4
    2001-01-05    5
    2001-01-06    6
    Name: series name, dtype: float64
    >>> xw.Range('D1', header=False).value = s

Pandas DataFrame converter
**************************

**options:** ``dtype=None, copy=False, index=1, header=1``

The first 2 options behave the same as when using ``pd.DataFrame()`` directly. ``ndim`` doesn't have an effect on
Pandas DataFrames as they are automatically read in with ``ndim=2``.

``index``: int or Boolean
    | When reading, it expects the number of index columns shown in Excel.
    | When writing, include or exclude the index by setting it to ``True`` or ``False``.

``header``: int or Boolean
    | When reading, it expects the number of column headers shown in Excel.
    | When writing, include or exclude the index and series names by setting it to ``True`` or ``False``.

For ``index`` and ``header``, ``1`` and ``True`` may be used interchangeably.

**Example:**

.. figure:: images/df_converter.png
  :scale: 55%

::

    >>> df = xw.Range('A1:D5').options(pd.DataFrame, header=2).value
    >>> df
        a     b
        c  d  e
    ix
    10  1  2  3
    20  4  5  6
    30  7  8  9

    # Writing back using the defaults:
    >>> Range('A1').value = df

    # Writing back and changing some of the options, e.g. getting rid of the index:
    >>> Range('B7').options(index=False).value = df

The same sample for **UDF** (starting in ``Range('A13')`` on screenshot) looks like this::

    @xw.func
    @xw.arg('x', pd.DataFrame, header=2)
    @xw.ret(index=False)
    def myfunction(x):
       # x is a DataFrame, do something with it
       return x