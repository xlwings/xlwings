.. _customaddin:

Custom Add-ins
==============

.. versionadded:: 0.21.5

Custom add-ins work on Windows and macOS and are essentially white-labeled xlwings add-ins that include all your ``RunPython`` functions and UDFs (UDFs are Windows only). You can build add-ins with and without an Excel ribbon. This tutorial assumes you're familiar with how xlwings and its configuration works.

Quickstart
----------

Start by running the following command on a command line (to create an add-in without a ribbon, you would leave away the ``--ribbon`` flag:

.. code-block::

   $ xlwings quickstart myproject --addin --ribbon

This will create the familiar quickstart folder with a Python file and an Excel file, but the Excel file is now in the ``xlam`` format.

* Double-click the Excel add-in to open it in Excel
* Add a new empty workbook (``Ctrl+N`` on Windows or ``Command+N`` on macOS)

You should see a new ribbon tab called ``MyAddin`` like this:

.. figure:: images/custom_ribbon_addin.png
    :scale: 40%

The add-in and VBA project is currently always called ``myaddin``, no matter what name you chose in the quickstart command. We'll see at the end of this tutorial how we can change that, but for now we'll stick with it.

Configuration
-------------

Compared to the xlwings add-in, the custom add-in has an additional level: the configuration sheet of the add-in itself which is the easiest way to configure simple add-ins with a static configuration. Let's open the VBA editor by clicking on ``Alt+F11`` (Windows) or ``Option+F11`` (macOS). In our project, select ``ThisWorkbook``, then change the Property ``IsAddin`` from ``True`` to ``False``:

.. figure:: images/custom_addin_vba_properties.png
    :scale: 40%

This will make the sheet ``_myaddin.conf`` visible. Activate it by renaming it to ``myaddin.conf``, then set your ``Interpreter`` or ``Conda`` settings. Once done, switch back to the VBA editor, select ``ThisWorkbook`` again, and change ```IsAddin`` back to ``True``. Now click the ``Run`` button under the ``My Addin`` ribbon tab and if you've configured the Python interpreter correctly, it will print ``Hello xlwings!`` into cell ``A1``.
Configure the ``xlwings.conf`` so it points to the correct Python interpreter. Most likely, you want to use an environment variable here so that you can eas

Changing the Ribbon menu
------------------------

To change the buttons and items in the ribbon menu, download and install the `Office RibbonX Editor <https://github.com/fernandreu/office-ribbonx-editor/releases>`_. Then open your add-in with it so you can change the XML code that defines your buttons etc. You will find more information about this part






