Debugging
=========

Since xlwings runs in every Python environment, you can use your preferred ways of debugging. When running xlwings
through Excel, there are a few tricks that make it easier to switch back and forth between Excel for testing and Python
for development and debugging.

To begin with, Excel will show any Python errors (but not Warnings) in a Message Box:

.. figure:: images/debugging_error.png

Consider the following code structure of your Python source code:

.. code-block:: python

    import os
    from xlwings import Workbook, Range

    def get_workbook():
        if __name__ == '__main__':
            # This expects the Excel file to sit next to this source file. Adjust accordingly.
            xl_file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'myfile.xlsm'))
            return Workbook(xl_file_path)
        else:
            return Workbook()

    def my_macro():
        wb = get_workbook()
        Range('A1').value = 1

    if __name__ == '__main__':
        my_macro()


``my_macro()`` can now easily be run from Python for debugging and from Excel for testing without having to change the
source code:

.. code-block:: vb.net

    Sub my_macro()
        RunPython ("import my_module; my_module.my_macro()")
    End Sub