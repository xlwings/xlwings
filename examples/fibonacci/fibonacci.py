"""
Copyright (C) 2014-2015, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import os
from xlwings import Workbook, Range


def fibonacci(n):
    """
    Generates the first n Fibonacci numbers.
    Adopted from: https://docs.python.org/3/tutorial/modules.html
    """
    result = []
    a, b = 0, 1
    while len(result) < n:
        result.append(b)
        a, b = b, a + b
    return result


def xl_fibonacci():
    """
    This is a wrapper around fibonacci() to handle all the Excel stuff
    """
    # Create a reference to the calling Excel Workbook
    wb = Workbook.caller()

    # Get the input from Excel and turn into integer
    n = int(Range('B1').value)

    # Call the main function
    seq = fibonacci(n)

    # Clear output
    Range('C1').vertical.clear_contents()

    # Return the output to Excel
    # zip() is used to push a list over in column orientation (list() needed on PY3)
    Range('C1').value = list(zip(seq))

if __name__ == "__main__":
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'fibonacci.xlsm'))
    Workbook.set_mock_caller(path)
    xl_fibonacci()