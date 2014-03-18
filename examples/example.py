from xlwings import Workbook, Range

wb = Workbook()  # Create a reference to the calling Excel Workbook

def fibonacci():
    """
    Generates the first n Fibonacci numbers.
    """
    # Get the input from Excel and turn into integer
    n = int(Range('B1').value)

    # Calculate the Fibonacci Sequence
    seq = [1, 1]
    for i in range(1, n):
        seq.append(seq[i-1] + seq[i])

    # Clear output
    Range('C1').vertical.clear_contents()

    # Return the output to Excel - use zip to get column format
    Range('C1').value = zip(seq[:n])