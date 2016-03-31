.. _vba:

VBA: Calling Python from Excel
==============================

xlwings VBA module
------------------

To get access to the ``RunPython`` function and/or to be able to run User Defined Functions (UDFs), you need to have the
xlwings VBA module available in your Excel workbook.

For new projects, by far the easiest way to get started is by using the command line client with the quickstart option,
see :ref:`command_line` for details::

    $ xlwings quickstart myproject


This will create a new folder in your current directory with a fully prepared Excel file and an empty Python file.

Alternatively, you can also open a new spreadsheet from a template (``$ xlwings template open``) or manually insert
the module in an existing workbook like so:

* Open the VBA editor with ``Alt-F11``
* Then go to ``File > Import File...`` and import the ``xlwings.bas`` file. It can be found in the directory of
  your ``xlwings`` installation.

If you don't know the location of your xlwings installation, you can find it as follows::

    $ python
    >>> import xlwings
    >>> xlwings.__path__

.. _vba_settings:

Settings
--------

While the defaults will often work out-of-the box, you can change the settings at the top of the xlwings VBA module
under ``Function Settings``::

    PYTHON_WIN = ""
    PYTHON_MAC = ""
    PYTHON_FROZEN = ThisWorkbook.Path & "\build\exe.win32-2.7"
    PYTHONPATH = ThisWorkbook.Path
    UDF_MODULES = ""
    UDF_DEBUG_SERVER = False
    LOG_FILE = ThisWorkbook.Path & "\xlwings_log.txt"
    SHOW_LOG = True
    OPTIMIZED_CONNECTION = False

* ``PYTHON_WIN``: This is the directory of the Python interpreter on Windows. ``""`` resolves to your default Python
  installation on the PATH, i.e. the one you can start by just typing ``python`` at a command prompt.
* ``PYTHON_MAC``: This is the directory of the Python interpreter on Mac OSX. ``""`` resolves to your default
  installation as per PATH on .bash_profile. To get special folders
  on Mac, type ``GetMacDir("Name")`` where ``Name`` is one of the following: ``Home``, ``Desktop``, ``Applications``,
  ``Documents``.
* ``PYTHON_FROZEN`` [Optional]: Currently only on Windows, indicates the directory of the exe file that has been frozen
  by either using ``cx_Freeze`` or ``py2exe``. Can be set to ``""`` if unused.
* ``PYTHONPATH`` [Optional]: If the source file of your code is not found, add the path here. Otherwise set it to ``""``.
* ``UDF_MODULES`` [Optional, Windows only]: Names of Python modules (without .py extension) from which the UDFs are being imported.
  Separate multiple modules by ";".
  Example: ``UDF_PATH = "common_udfs;myproject"``
  Default: ``UDF_PATH = ""`` defaults to a file in the same directory of the Excel spreadsheet with the same name but ending in ``.py``.
* ``UDF_DEBUG_SERVER``: Set this to True if you want to run the xlwings COM server manually for debugging, see :ref:`debugging`.
* ``LOG_FILE`` [Optional]: Leave empty for default location (see below) or provide directory including file name.
* ``SHOW_LOG``: If False, no pop-up with the Log messages (usually errors) will be shown. Use with care.
* ``OPTIMIZED_CONNECTION``: Currently only on Windows, use a COM Server for an efficient connection (experimental!)

.. _log:

LOG_FILE default locations
**************************

* Windows: ``%APPDATA%\xlwings_log.txt``
* Mac with Excel 2011: ``/tmp/xlwings_log.txt``
* Mac with Excel 2016: ``~/Library/Containers/com.microsoft.Excel/Data/xlwings_log.txt``

.. note:: If the settings (especially ``PYTHONPATH`` and ``LOG_FILE``) need to work on Windows on Mac, use backslashes
    in relative file path, i.e. ``ThisWorkbook.Path & "\mydirectory"``.

.. _run_python:

Call Python with "RunPython"
----------------------------

After your workbook contains the xlwings VBA module with potentially adjusted Settings, go to ``Insert > Module`` (still
in the VBA-Editor). This will create a new Excel module where you can write your Python call as follows (note that the ``quickstart``
or ``template`` commands already add an empty Module1, so you don't need to insert a new module manually):

.. code-block:: vb.net

    Sub MyMacro()
        RunPython ("import mymodule; mymodule.rand_numbers()")
    End Sub

This essentially hands over control to ``mymodule.py``:

.. code-block:: python

    import numpy as np
    from xlwings import Workbook, Range

    def rand_numbers():
        """ produces std. normally distributed random numbers with shape (n,n)"""
        wb = Workbook.caller()  # Creates a reference to the calling Excel file
        n = int(Range('Sheet1', 'B1').value)  # Write desired dimensions into Cell B1
        rand_num = np.random.randn(n, n)
        Range('Sheet1', 'C3').value = rand_num

You can then attach ``MyMacro`` to a button or run it directly in the VBA Editor by hitting ``F5``.

.. note:: Always place ``Workbook.caller()`` within the function that is being called from Excel and not outside as
    global variable. Otherwise it prevents Excel from shutting down properly upon exiting and
    leaves you with a zombie process when you use ``OPTIMIZED_CONNECTION = True``.

Function Arguments and Return Values
------------------------------------

While it's technically possible to include arguments in the function call within ``RunPython``, it's not very convenient.
To do that easily and to also be able to return values from Python, use UDFs, see :ref:`udfs` - however, this is currently limited
to Windows only.