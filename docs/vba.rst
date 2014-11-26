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
    >>> xlwings.__file__



Settings
--------

While the defaults will often work out-of-the box, you can change the settings at the top of the xlwings VBA module
under ``Function Settings``::

    PYTHON_WIN = ""
    PYTHON_MAC = ""
    PYTHON_FROZEN = ThisWorkbook.Path & "\build\exe.win32-2.7"
    PYTHONPATH = ThisWorkbook.Path
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
* ``LOG_FILE``: Directory **including** the file name. This file is necessary for error handling.
* ``SHOW_LOG``: If False, no pop-up with the Log messages (usually errors) will be shown. Use with care.
* ``OPTIMIZED_CONNECTION``: Currently only on Windows, use a COM Server for an efficient connection (experimental!)

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
        n = Range('Sheet1', 'B1').value  # Write desired dimensions into Cell B1
        rand_num = np.random.randn(n, n)
        Range('Sheet1', 'C3').value = rand_num

You can then attach ``MyMacro`` to a button or run it directly in the VBA Editor by hitting ``F5``.

.. note:: Always place ``Workbook.caller()`` within the function that is called from Excel and not outside as
    module-wide global variable. Otherwise it doesn't get garbage collected with ``OPTIMIZED_CONNECTION = True``
    which prevents Excel from shutting down properly upon exiting and and leaves you with a zombie process.