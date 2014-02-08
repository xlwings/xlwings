## The easiest way to deploy your Python powered Excel tools

xlwings makes it easy to deploy your Python powered Excel tools on Windows. Just zip up your Excel and Python files and send them to your users with the following instructions:

> 1. Install a scientific Python distribution like [Anaconda](https://store.continuum.io/cshop/anaconda/), [Canopy](https://enthought.com/downloads/) or [WinPython](https://code.google.com/p/winpython/) *.
> 2. Extract the zip file and work with the Excel workbook as usual.

*Note that xlwings works with every Python installation but then you need to manually install [pywin32](http://sourceforge.net/projects/pywin32/) which can sometimes be a bit troublesome.

## Give it a try!
Download the zip file (button at the top) and follow the two steps above. Open the `example.xlsm` and try the examples. Note that the only files needed to run the example are: `example.xlsm`, `example.py` and `xlwings.py`.

## Developing the Excel tool

On the developer side, xlwings is equally easy:

1. Open the VBA editor (`Alt-F11`), then go to `File > Import File...` and import the `xlwings.bas` file. Alternatively, just work off the `example.xlsm` and continue to the next step:
2. Call Python from VBA like so:
 
    ```VB.net
    Sub Example()
        RunPython ("import example; example.rand_numbers()")
    End Sub
    ```

3. This essentially hands over control to `example.py` :

    ```python
    import numpy as np
    from xlwings import Workbook

    wb = Workbook()  # Creates a reference to the calling Excel file

    def rand_numbers():
        """ produces standard normally distributed random numbers with dim (n,n)"""
        n = wb.range('Sheet1', 'B1').value
        rand_num = np.random.randn(n, n)
        wb.range('Sheet1', 'C3').value = rand_num
    ```

## Interactive use and debugging
xlwings let's you comfortably develop, debug and interact with Excel by calling `Workbook()` either with no arguments to work off a new file or with the full path to your Excel file (from your favorite Python environment). The `Range` object can be used as shortcut for `wb.range()`. It always refers to latest created workbook. Also, when omitting the sheet name, it refers to the currently active sheet.

```python
>>> from xlwings import Workbook, Range
>>> wb = Workbook(r'C:\full\path\to\file.xlsx')  # Use Workbook() for a new file
>>> Range('A1').value = 'Hello xlwings!'
>>> Range('A1').value
u'Hello xlwings!'
```

## License

BSD (3-clause) license