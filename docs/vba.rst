.. _vba:

VBA: Calling Python from Excel
==============================

Import the xlwings VBA module into Excel
----------------------------------------

To get access to the ``RunPython`` function in VBA, you need to import the VBA module ``xlwings.bas`` into the VBA
editor:

* Open the VBA editor with ``Alt-F11``
* Then go to ``File > Import File...`` and import the ``xlwings.bas`` file. It can be found in the directory of
  your ``xlwings`` installation.

If you don't know the location of your xlwings installation, you can find it as follows::

    $ python
    >>> import xlwings
    >>> xlwings.__path__

An even easier way is to start from a template that already includes the xlwings VBA module and
boilerplate code. Use the command line client like this (for details see: :ref:`command_line`)::

    $ xlwings template open

.. _vba_settings:

Settings
--------

While the defaults will often work out-of-the box, you can change the settings at the top of the xlwings VBA module
under ``Function Settings``::

    PYTHON_WIN = ""
    PYTHON_MAC = ""
    PYTHON_FROZEN = ThisWorkbook.Path & "\build\exe.win32-2.7"
    PYTHONPATH = ThisWorkbook.Path
    UDF_PATH = ""
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
* ``UDF_PATH`` [Optional, Windows only]: Full path to a Python file from which the User Defined Functions are being imported.
  Example: ``UDF_PATH = ThisWorkbook.Path & "\functions.py"``
  Default: ``UDF_PATH = ""`` defaults to a file in the same directory of the Excel spreadsheet with the same name but ending in ``.py``.
* ``LOG_FILE`` [Optional]: Leave empty for default location (see below) or provide directory including file name.
* ``SHOW_LOG``: If False, no pop-up with the Log messages (usually errors) will be shown. Use with care.
* ``OPTIMIZED_CONNECTION``: Currently only on Windows, use a COM Server for an efficient connection (experimental!)

.. _log:

LOG_FILE default locations
**************************

* Windows: ``%APPDATA%\xlwings_log.txt``
* Mac 2011: ``/tmp/xlwings_log.txt``
* Mac 2016: ``/Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings_log.txt``

.. note:: If the settings (especially ``PYTHONPATH`` and ``LOG_FILE``) need to work on Windows on Mac, use backslashes
    in relative file path, i.e. ``ThisWorkbook.Path & "\mydirectory"``.

.. note:: ``OPTIMIZED_CONNECTION = True`` works currently on **Windows only** and is still experimental! This will
  use a COM server that will keep the connection to Python alive between different calls and is therefore much more
  efficient. However, changes in the Python code are not being picked up until the ``pythonw.exe`` process is restarted
  by killing it manually in the Windows Task Manager. The suggested workflow is hence to set
  ``OPTIMIZED_CONNECTION = False`` for development and to only set it to ``True`` for production.


Subtle difference between the Windows and Mac Version
-----------------------------------------------------

* **Windows**: After calling the Macro (e.g. by pressing a button), Excel waits until Python is done.

* **Mac**: After calling the Macro, the call returns instantly but Excel's Status Bar turns into ``Running...`` during the
  duration of the Python call.

.. _run_python:

Call Python with "RunPython"
----------------------------

After you have imported the xlwings VBA module and potentially adjusted the Settings, go to ``Insert > Module`` (still
in the VBA-Editor). This will create a new Excel module where you can write your Python call as follows:

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