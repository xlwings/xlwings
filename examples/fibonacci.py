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

    # Return the output to Excel - use zip to get column format
    Range('C1').value = zip(seq)