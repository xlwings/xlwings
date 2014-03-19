"""
Make Excel fly!
xlwings is the easiest way to deploy your Python powered Excel tools on Windows.
Homepage and documentation: http://xlwings.org
See also: http://zoomeranalytics.com

Copyright (C) 2014, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""

from xlwings import Workbook, Range

wb = Workbook()  # Create a reference to the calling Excel Workbook


def fibonacci(n):
    """
    Generates the first n Fibonacci numbers.
    """
    seq = [1, 1]
    for i in range(1, n-1):
        seq.append(seq[i-1] + seq[i])
    return seq[:n]


def xl_fibonacci():
    """
    This is a wrapper around fibonacci() to handle all the Excel stuff
    """
    # Get the input from Excel and turn into integer
    n = int(Range('B1').value)

    # Call the main function
    seq = fibonacci(n)

    # Clear output
    Range('C1').vertical.clear_contents()

    # Return the output to Excel - turn into list of list for column orientation
    Range('C1').value = [list(i) for i in zip(seq)]