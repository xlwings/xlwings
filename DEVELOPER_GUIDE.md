# Developer Guide

## Python package

The source for the Python package is in the `xlwings` directory.

1. Fork xlwings repository on GitHub into your own account
2. Clone your forked reposistory: `git clone <your forked git url>`
3. `cd xlwings`
4. With the desired development environment activated: `python setup.py develop`. This will install xlwings like a standard package
   but runs from your cloned source code, i.e. you can edit/debug the xlwings code.
5. Install the optional dependencies according to your platform with: `pip install -r requirements/devwin.txt` or `.../devmac.txt`

## Addin

Install the addin in Excel by going to `Developer` > `Excel Add-in` > `Browse` and pointing to the addin in the source code,
i.e. under `xlwings/addin/xlwings.xlam`.

To change the buttons, you need to download the CustomUIEditor (only runs on Windows) from 
http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2006/05/26/customuieditor.aspx or
https://www.rondebruin.nl/win/winfiles/OfficeCustomUIEditorSetup.zip

The code is pure VBA code.

## dlls

The source for the dlls (used for the UDF related stuff on Windows) is under `src`. It's written in C++.

The dlls are almost never touched, so you should not bother to setup the development environment. What you should do is
download the latest version of xlwings from https://pypi.org, unpack it and place the two dlls next to the Python interpreter,
then rename them into `xlwings32-dev.dll` and `xlwings64-dev.dll`. Note that the bitness refers to the Excel
installation and not to the Python installation!

If you ever need to change the C++ source, then download Visual Studio Community 2015 to open and compile the 
project there.

## Tests

Run `nosetests` from the `xlwings` dir, see also `runtests.py`, but this is also outdated and might be replaced
by something like `tox` again now that numpy/pandas are available as wheels via pypi.
The tests are currently very slow. They were OK with older versions of Excel but they have to be rewritten
to run reasonably fast (i.e. not always open/close the whole workbook).
Also, the tests are standard unittests, so `nose` is not really required to run them.
