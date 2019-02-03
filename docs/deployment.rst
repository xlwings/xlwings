.. _deployment:

Deployment
==========

Zip files
---------

.. versionadded:: 0.15.2

To make it easier to distribute, you can zip up your Python code into a zip file. If you use UDFs, this will disable the
automatic code reload, so this is a feature meant for distribution, not development. In practice, this means that when
your code is inside a zip file, you'll have to click on re-import to get any changes.

If you name your zip file like your Excel file (but with ``.zip`` extension) and place it in the same folder as your
Excel workbook, xlwings will automatically find it (similar to how it works with a single python file).

If you want to use a different directory, make sure to add it to the ``PYTHONPATH`` in your config (Ribbon or config file):

.. code-block:: bash

    PYTHONPATH, "C:\path\to\myproject.zip"

RunFrozenPython
---------------

.. versionchanged:: 0.15.2

You can use a freezer like PyInstaller, cx_Freeze, py2exe etc. to freeze your Python module into an executable so that
the recipient doesn't have to install a full Python distribution.

.. note::
    * This does not work with UDFs.
    * Currently only available on Windows, but support for Mac should be easy to add.
    * You need at least 0.15.2 to support arguments

Use it as follows:

.. code-block:: basic

    Sub MySample()
        RunFrozenPython "C:\path\to\dist\myproject\myproject.exe arg1 arg2"
    End Sub


