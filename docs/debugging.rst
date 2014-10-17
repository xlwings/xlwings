.. _debugging:

Debugging
=========

Since xlwings runs in every Python environment, you can use your preferred way of debugging. When running xlwings
through Excel, there are a few tricks that make it easier to switch back and forth between Excel (for testing) and
Python (for development and debugging).

To begin with, Excel will show any Python errors (but not warnings) in a Message Box:

.. figure:: images/debugging_error.png
    :scale: 65%

.. note:: On Mac, if the ``import`` of a module/package fails before xlwings is imported, the popup will not be shown and the StatusBar
    will not be reset. However, the error will still be logged in the log file.

Consider the following code structure of your Python source code:

.. code-block:: python

    import os
    from xlwings import Workbook, Range

    def my_macro(workbook_path=None):
        wb = Workbook(workbook_path)
        Range('A1').value = 1

    if __name__ == '__main__':
        # To run from Python, not needed when called from Excel.
        # Expects the Excel file next to this source file, adjust accordingly.
        path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'myfile.xlsm'))
        my_macro(path)


``my_macro()`` can now easily be run from Python for debugging and from Excel for testing without having to change the
source code:

.. code-block:: vb.net

    Sub my_macro()
        RunPython ("import my_module; my_module.my_macro()")
    End Sub