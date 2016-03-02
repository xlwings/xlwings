Converters
==========

Introduced with v0.7.0, converters define how Excel Ranges and their values are being converted both when being
**read** and **written**. They also provide a consistent experience across **xlwings.Range** objects and
**User Defined Functions** (UDFs).

Converters are explicitely set with the ``as_`` argument in the ``options`` method when manipulating ``xlwings.Range`` objects
or in the ``@xw.arg`` and ``@xw.ret`` decorators when using UDFs. If no converter is specified, only the base converter
is applied.

**Syntax:**

==============================  ===========================================================  ===========
****                            **Range**                                                    **UDF**
==============================  ===========================================================  ===========
**reading**                     ``Range.options(as_=None, **kwargs).value``                  ``@arg('x', as_=None, **kwargs)``
**writing**                     ``Range.options(as_=None, **kwargs).value = myvalue``        ``@ret(as_=None, **kwargs)``
==============================  ===========================================================  ===========

.. note:: Keyword arguments (``kwargs``) may refer to the specific converter or the base converter.
  For example, to set the ``numbers`` option in the base converter and the ``index`` option in the DataFrame converter,
  you would do::

      Range('A1:C3').options(pd.DataFrame, index=False, numbers=int).value

Base Converter
--------------
When no options are specified, the following rules are applied:

* single cells are read in as ``floats`` in case the Excel cell holds a number, as ``unicode`` in case it holds text,
  as ``datetime`` if it contains a date and as ``None`` in case it is empty.
* columns/rows are read in as lists, e.g. ``[None, 1.0, 'a string']``
* 2d cell ranges are read in as list of lists, e.g. ``[[None, 1.0, 'a string'], [None, 2.0, 'another string']]``

The following options can be set:

* **numbers**

  The base converter reads in numbers as ``float``, you can change that like so::

    >>> import xlwings as xw
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

There are built-in converters for **dictionaries**, **NumPy arrays**, **Pandas Series** and **DataFrames**. New,
customized converters can also be added (docs will follow).

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