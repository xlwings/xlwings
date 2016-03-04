.. _command_line:

Command Line Client
===================

xlwings comes with a command line client that makes it easy to set up workbooks and install the developer add-in.
On Windows, type the commands into a ``Command Prompt``, on Mac, type them into a ``Terminal``.

Quickstart
----------

* ``xlwings quickstart myproject``

This command is by far the fastest way to get off the ground: It creates a new folder ``myproject`` with the necessary
Excel workbook (including the xlwings VBA module) and a Python file, ready to be used right away:

.. code::

  myproject
    |--myproject.xlsm
    |--myproject.py

.. versionadded:: 0.6.4

Add-in (Currently Windows-only)
-------------------------------

The add-in is currently in an early stage and only provides one button to import User Defined Functions (UDFs). As
such, it is only a developer add-in and not necessary to run Workbooks with xlwings.

.. note:: Excel needs to be closed before installing/updating the add-in. If you're still getting an error,
  start the Task Manager and make sure there are no ``EXCEL.EXE`` processes left.

* ``xlwings addin install``: Copies the xlwings add-in to the XLSTART folder

* ``xlwings addin update``: Replaces the current add-in with the latest one

* ``xlwings addin remove``: Removes the add-in from the XLSTART folder

* ``xlwings addin status``: Shows if the add-in is installed together with the installation path

After installing the add-in, it will be available as xlwings tab on the Excel Ribbon.

.. versionadded:: 0.6.0

Template
--------

* ``xlwings template open``: Opens a new Workbook with the xlwings VBA module

* ``xlwings template install``: Copies the xlwings template file to the correct Excel folder, see below

* ``xlwings template update``: Replaces the current xlwings template with the latest one

* ``xlwings template remove``: Removes the template from Excel's template folder

* ``xlwings template status``: Shows if the template is installed together with the installation path

After installing, the templates are accessible via Excel's Menu:

* Win (Excel 2007, 2010): ``File > New > My templates``
* Win (Excel 2013, 2016): There's an additional step needed as explained `here <https://support.office.com/en-us/article/Where-are-my-custom-templates-88ed77ca-df34-49e9-9087-3f01ae296e6e/>`_
* Mac (Excel 2011, 2016): ``File > New from template``

.. versionadded:: 0.6.0

RunPython
---------

Only required if you are on Mac, are using Excel 2016 and have xlwings installed via conda or as part of Anaconda.
To enable the ``RunPython`` calls in VBA, run this one time:

``xlwings runpython install``

Alternatively, install xlwings with ``pip``.

.. versionadded:: 0.7.0