.. _xlwings_addin:

Add-in
======

.. figure:: images/ribbon.png
    :scale: 80%

The xlwings add-in is the preferred way to be able to use ``RunPython`` or ``UDFs``. Note that you don't need an add-in
if you just want to manipulate Excel from Python via xlwings.

.. note:: The ribbon of the add-in is compatible with Excel >= 2007 on Windows and >= 2016 on Mac. You could, however,
  use the add-in with earlier versions but you would need to change the settings directly in the config file, see below.
  On Mac, all UDF related functionality doesn't do anything, although the respecitve button and fields are currently visible.

Installation
------------

To install the add-in, it's easiest to use the command line client: ``xlwings addin install``.

Then, to use ``RunPython`` or ``UDFs`` in a workbook, you need to set a reference to ``xlwings`` in the VBA editor, see
screenshot (Windows: ``Tools > References...``, Mac: it's on the lower left corner of the VBA editor). Note that when
you create a workbook via ``xlwings quickstart``, the reference is already set.

.. figure:: images/vba_reference.png
    :scale: 40%

Global Settings
---------------

While the defaults will often work out-of-the box, you can change the global settings directly in the add-in:

* ``Interpreter``: This is the path to the Python interpreter (works also with virtual or conda envs),
  e.g. ``"C:\Python35\pythonw.exe"`` or ``"/usr/local/bin/python3.5"``. An empty field defaults to ``pythonw`` that
  expects the interpreter to be set in the ``PATH`` on Windows or ``.bash_profile`` on Mac.
* ``PYTHONPATH``: If the source file of your code is not found, add the path here.
* ``UDF_MODULES``: Names of Python modules (without .py extension) from which the UDFs are being imported.
  Separate multiple modules by ";".
  Example: ``UDF_MODULES = "common_udfs;myproject"``
  The default imports a file in the same directory as the Excel spreadsheet with the same name but ending in ``.py``.
* ``Debug UDFs``: Check this box if you want to run the xlwings COM server manually for debugging, see :ref:`debugging`.
* ``Log File``: Leave empty for default location (see below) or provide the full path, e.g. .
* ``RunPython: Use UDF Server``:  Uses the same COM Server for RunPython as for UDFs. This will be faster, as the
  interpreter doesn't shut down after each call.
* ``Restart UDF Server``: This shuts down the UDF Server/Python interpreter. It'll be restarted upon the next function call.

Config File
-----------

The settings in the xlwings Ribbon are stored in a config file that can also be manipulated externally. The location is

* Windows: ``.xlwings\xlwings.conf`` in your user folder  
* Mac: ``~/Library/Containers/com.microsoft.Excel/Data/xlwings.conf``

The format is as follows (keys are uppercase):

.. code-block:: bash

    "INTERPRETER","pythonw"
    "PYTHONPATH",""

Workbook Settings
-----------------

The global settings of the Ribbon/Config file can be overridden for each workbook by adding a sheet with the 
name ``xlwings.conf``. When you create a new project with ``xlwings quickstart``, it'll already have such a sheet
but you need to rename it to ``xlwings.conf`` to make it active.


.. figure:: images/workbook_config.png
    :scale: 40%


Alternative: Standalone VBA module
----------------------------------

Sometimes it might be useful to run xlwings code without having to install an add-in first. To do so, you
need to use the ``standalone`` option when creating a new project: ``xlwings quickstart myproject --standalone``.

This will add the content of the add-in as a single VBA module so you don't need to set a reference to the add-in anymore.
It will still read in the settings from your ``xlwings.conf`` if you don't override them by using a sheet with the name ``xlwings.conf``.


.. _log:

Log File default locations
--------------------------

These log files are used for the error pop-up windows:

* Windows: ``%APPDATA%\xlwings.log``
* Mac with Excel 2011: ``/tmp/xlwings.log``
* Mac with Excel 2016: ``~/Library/Containers/com.microsoft.Excel/Data/xlwings.log``