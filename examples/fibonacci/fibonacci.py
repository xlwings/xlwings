"""
Copyright (C) 2014, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
from xlwings import Workbook, Range


def fibonacci(n):
    """
    Generates the first n Fibonacci numbers.
    TODO: Pythonic implementation
    """
    seq = [1, 1]
    for i in range(1, n-1):
        seq.append(seq[i-1] + seq[i])
    return seq[:n]


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

    # Return the output to Excel - turn into list of list for column orientation
    Range('C1').value = [list(i) for i in zip(seq)]

if __name__ == "__main__":
    xl_fibonacci()