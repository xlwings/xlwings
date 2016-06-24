"""
Copyright (C) 2014-2016, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
from xlwings import Book, Range


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
    wb = Book.caller()

    import time
    time.sleep(2)

    # Get the input from Excel and turn into integer
    n = wb.sheets[0].range('B1').options(numbers=int).value

    # Call the main function
    seq = fibonacci(n)

    # Clear output
    wb.sheets[0].range('C1').vertical.clear_contents()

    # Return the output to Excel in column orientation
    wb.sheets[0].range('C1').options(transpose=True).value = seq

if __name__ == "__main__":
    # Used for frozen executable
    xl_fibonacci()
