.. _command_line:

Command Line Client
===================

xlwings comes with a command line client that makes it easy to set up workbooks and install the add-in.
On Windows, type the commands into a ``Command Prompt``, on Mac, type them into a ``Terminal``.

Quickstart
----------

* ``xlwings quickstart myproject``

This command is by far the fastest way to get off the ground: It creates a new folder ``myproject`` with an
Excel workbook that already has the reference to the xlwings addin and a Python file, ready to be used right away:

.. code::

  myproject
    |--myproject.xlsm
    |--myproject.py

If you want to use xlwings via VBA module instead of addin, use the ``--standalone`` or ``-s`` flag:

``xlwings quickstart myproject --standalone``

Add-in
------

The `addin` command makes it easy on Windows to install/remove the addin. On Mac, you need to install it manually, but
``xlwings addin install`` will show you how to do it.

.. note:: Excel needs to be closed before installing/updating the add-in via command line. If you're still getting an error,
  start the Task Manager and make sure there are no ``EXCEL.EXE`` processes left.

* ``xlwings addin install``: Copies the xlwings add-in to the XLSTART folder

* ``xlwings addin update``: Replaces the current add-in with the latest one

* ``xlwings addin remove``: Removes the add-in from the XLSTART folder

* ``xlwings addin status``: Shows if the add-in is installed together with the installation path

After installing the add-in, it will be available as xlwings tab on the Excel Ribbon.

.. versionadded:: 0.6.0


RunPython
---------

Only required if you are on Mac, are using Excel 2016 and have xlwings installed via conda or as part of Anaconda.
To enable the ``RunPython`` calls in VBA, run this one time:

``xlwings runpython install``

Alternatively, install xlwings with ``pip``.

.. versionadded:: 0.7.0