Converters
==========

Introduced with v0.7.0, converters define how Excel Ranges and their values are being converted both when being
**read** and **written**. They also provide a consistent experience across **xlwings.Range** objects and
**User Defined Functions** (UDFs).

Converters are set with the ``as_`` argument in the ``options`` method when manipulating ``xlwings.Range`` objects
or in the ``@xw.arg`` and ``@xw.ret`` decorators when using UDFs.

**Syntax:**

==============================  ===========================================================  ===========
****                            **Range**                                                    **UDF**
==============================  ===========================================================  ===========
**reading**                     ``Range.options(as_=None, **kwargs).value``                  ``@arg('x', as_=None, **kwargs)``
**writing**                     ``Range.options(as_=None, **kwargs).value = myvalue``        ``@ret(as_=None, **kwargs)``
==============================  ===========================================================  ===========

**Note:** Converter-specific keyword arguments (``**kwargs``) may be mixed with standard converter keyword arguments.
The following example mixes options from the Pandas DataFrame converter (``index``) with options from the standard
converter (``dates``): ``Range('A1:C3').options(pd.DataFrame, index=False, dates=dt.date).value``.

Standard Converter
------------------
When no options are specified, the following default conversions are applied:

* single cells are read in as ``floats`` in case the Excel cell is a number, as ``unicode`` in case it is a cell with text,
  as ``datetime`` if it contains a date and as ``None`` in case it is empty.
* columns/rows are read in as lists, e.g. ``[None, 1.0, 'a string']``
* multiple cells are read in as list of lists, e.g. ``[[None, 1.0, 'a string'], [None, 2.0, 'another string']]``

``numbers``, ``dates`` and ``empty`` can be changed as follows:

* **numbers**

  The standard converter for Excel cells with numbers is ``float``, you can change that with the ``numbers`` keyword::

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

transpose and expand options
----------------------------

* ``transpose`` option: This works for reading and writing and allows us to e.g. write a list in column orientation to Excel:

  - Range: ``Range('A1').options(transpose=True).value = [1, 2, 3]``

  - UDFs:

    .. code-block:: python

        @xw.arg('x', transpose=True)
        @xw.ret(transpose=True)
        def myfunction(x):
            # x will be returned unchanged as transposed both when reading and writing
            return x

* ``expand`` option: This works the same as the Range properties ``table``, ``vertical`` and ``horizontal`` but is
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

There are built-in converters for **dictionaries**, **numpy arrays**, **pandas series** and **dataframes**. New,
specialized converters can also be written, see [TODO].
Converters are chosen via the first argument (``as_``) in ``Range.options`` or ``@xw.arg`` and ``@xw.ret``, respecitvely.

Dictionary converter
--------------------

The dictionary converter turns two Excel columns into dictionaries. If you want data laid out in rows to be read in
as dictionary, use ``transpose``:

  .. figure:: images/dict_converter.png
    :scale: 80%

  ::

    >>> Range('A1:B2').options(dict).value
    {'a': 1.0, 'b': 2.0}
    >>> Range('A4:B5').options(dict, transpose=True).value
    {'a': 1.0, 'b': 2.0}