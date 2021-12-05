.. _vba:

RunPython
=========

xlwings add-in
--------------

To get access to ``Run main`` (new in v0.16) button or the ``RunPython`` VBA function, you'll need the xlwings addin (or VBA module), see :ref:`xlwings_addin`.

For new projects, the easiest way to get started is by using the command line client with the quickstart command,
see :ref:`command_line` for details::

    $ xlwings quickstart myproject

.. _run_python:

Call Python with "RunPython"
----------------------------

In the VBA Editor (``Alt-F11``), write the code below into a VBA module. ``xlwings quickstart`` automatically
adds a new module with a sample call. If you rather want to start from scratch, you can add a new module via ``Insert > Module``.

.. code-block:: vb.net

    Sub HelloWorld()
        RunPython "import hello; hello.world()"
    End Sub

This calls the following code in ``hello.py``:

.. code-block:: python

    # hello.py
    import numpy as np
    import xlwings as xw

    def world():
        wb = xw.Book.caller()
        wb.sheets[0].range('A1').value = 'Hello World!'

You can then attach ``HelloWorld`` to a button or run it directly in the VBA Editor by hitting ``F5``.

.. note:: Place ``xw.Book.caller()`` within the function that is being called from Excel and not outside as
    global variable. Otherwise it prevents Excel from shutting down properly upon exiting and
    leaves you with a zombie process when you use ``Use UDF Server = True``.

Function Arguments and Return Values
------------------------------------

While it's technically possible to include arguments in the function call within ``RunPython``, it's not very convenient.
Also, ``RunPython`` does not allow you to return values. To overcome these issues, use UDFs, see :ref:`udfs` - however,
this is currently limited to Windows only.