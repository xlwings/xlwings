.. _jupyternotebooks:

Excel and Jupyter Notebooks
===========================

When you work with Jupyter notebooks, you may use Excel as an interactive data viewer or scratchpad. If you want to quickly read in a pandas DataFrame from Excel or want to view a DataFrame that lives in a Jupyter notebook in Excel, you can use the two convenience functions :meth:`view <xlwings.view>` and :meth:`read <xlwings.read>`.

.. note::
    The :meth:`view <xlwings.view>` and :meth:`read <xlwings.read>` functions should exclusively be used for interactive work. If you write scripts, use the full xlwings API as introduced under :ref:`syntax_overview`.

The view function
-----------------

The view function accepts pretty much any object of interest, whether that's a number, a string, a nested list or a NumPy array or a pandas DataFrame. By default, it writes the data into an Excel table in a new workbook. If you wanted to reuse the same workbook, provide a ``sheet`` object, e.g. ``view(df, sheet=xw.sheets.active)``, for further options see :meth:`view <xlwings.view>`.

.. figure:: images/xw_view.png

.. versionchanged:: 0.21.5 Earlier versions were not formatting the output as Excel table

The read function
-----------------

To read in a range in an Excel sheet as pandas DataFrame, use the ``read`` function. If you only select one cell, it will auto-expand to cover the whole range. If, however, you select a specific range that is bigger than one cell, it will read in only the selected cells. If the data in Excel does not have an index or header, set them to ``False`` like this: ``xw.read(index=False)``, see also :meth:`read <xlwings.read>`.

.. figure:: images/xw_read.png

.. versionadded:: 0.21.5
