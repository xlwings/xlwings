.. _udfs:

User Defined Functions (UDFs)
=============================

.. note:: This feature is currently experimental and only available on Windows. It is the first step of an integration
    with `ExcelPython <http://ericremoreynolds.github.io/excelpython//>`_. Currently, there is no support for
    ExcelPython's decorator functionality and add-in. Also, ExcelPython's config file is obsolete as settings are being
    handled in the xlwings VBA module.


In it's current implementation, the UDF functionality is most useful to get access to existing Python code like for example
NumPy's powerful set of functions. This is best explained in a simple example: To expose NumPy's
`pseudo-inverse <http://docs.scipy.org/doc/numpy/reference/generated/numpy.linalg.pinv.html>`_, you would write the
following VBA code:

.. code-block:: vb.net

    Public Function pinv(x As Range)
    On Error GoTo Fail:
        Set numpy_array = Py.GetAttr(Py.Module("numpy"), "array")
        Set pseudo_inv = Py.GetAttr(Py.GetAttr(Py.Module("numpy"), "linalg"), "pinv")
        Set x_array = Py.Call(numpy_array, Py.Tuple(x.Value))
        Set result_array = Py.Call(pseudo_inv, Py.Tuple(x_array))
        Set result_list = Py.Call(result_array, "tolist")
        pinv = Py.Var(result_list)
        Exit Function
    Fail:
        pinv = Err.Description
    End Function

This then enables you to use ``pinv()`` as array function from Excel cells:

.. figure:: images/udf_example.png

1. Fill a spreadsheet with the following numbers as shown on the image:

   ``A1``: 1, ``A2``: 2, ``B1``: 2, ``B2``: 4

2. Select cells ``D1:E2``.

3. Type ``pinv(A1:B2)`` into D1 while D1:E2 are still selected.

4. Hit ``Ctrl-Shift-Enter`` to enter the array formula. If done right, the formula bar will automatically
   wrap the formula with curly braces ``{}``.

.. note:: Note that UDFs use a COM server and don't automatically reload when the Python code is changed (same behavior
    as if you set ``OPTIMIZED_CONNECTION = True`` for macros). To reload, you currently need to kill the ``pythonw.exe`` process
    manually from the Windows Task Manager. Recalculating the UDFs then causes the COM server to restart.

Further documentation
---------------------

For more in depth documentation at this point in time, please refer directly to the
`ExcelPython <http://ericremoreynolds.github.io/excelpython//>`_ project, mainly the following docs:

* `A very simple usage example <https://github.com/ericremoreynolds/excelpython/blob/master/docs/tutorials/Usage01.md>`_
* `A more practical use of ExcelPython <https://github.com/ericremoreynolds/excelpython/blob/master/docs/tutorials/Usage02.md>`_
* `Putting it all together <https://github.com/ericremoreynolds/excelpython/blob/master/docs/tutorials/Usage03.md>`_
* `Ranges, lists and SAFEARRAYs <https://github.com/ericremoreynolds/excelpython/blob/master/docs/tutorials/Usage04.md>`_
