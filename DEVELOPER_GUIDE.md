# Developer Guide

## Python package

The source for the Python package is in the `xlwings` directory.

1. Fork xlwings repository on GitHub into your own account
2. Clone your forked repository: `git clone <your forked git url>`
3. `cd xlwings`
4. With the desired development environment activated: `pip install -e .`. This will install xlwings like a standard package
   but runs from your cloned source code, i.e. you can edit/debug the xlwings code.
5. Install the optional dependencies according to your platform with: `pip install -r requirements/devwin.txt` or `.../devmac.txt`

## Mac implementation

All mac specific code is in `xlwings/_xlmac.py`. To find out the syntax of a new feature, it sometimes works by just looking at the existing
code and comparing it with the dictionaries exported by `ASDictionary` (see under `resources/mac`).
If that doesn't work, you'll need to find out the corresponding syntax in `AppleScript`, e.g. by looking at `Excel2004AppleScriptRef.pdf`
under `resouces/mac` or by searching the internet. Then use `ASTranslate` to translate it into `appscript` syntax. Unfortunately,
ASTranslate fails on the latest versions of macOS. A workaround is to run it on an old macOS (e.g. OS 10.11 - El Capitan) in a VM.

Links:

[appscript homepage](http://appscript.sourceforge.net/)  
[ASTranslate & ASDictionary download](https://sourceforge.net/projects/appscript/files/)  
[appscript source code](https://sourceforge.net/p/appscript/code/HEAD/tree/)

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
project there. When installing, make sure to select `Custom Installation` so you can activate the checkbox for `Visual C++` under
programming languages.

If you need to debug the dll, in Visual Studio do the following:

* Double-click `xlwings.sln` to open the project in Visual Studio 2015.
* Build the project in `Debug` mode (select `Win32` if your Excel is 32 bit and `x64` if your Excel is 64 bit). Note that the Python bitness does not matter!
* Make sure that Excel calls that dll that was built in debug mode. In the xlwings addin (password: `xlwings`) you could e.g. override the path in `XLPyLoadDLL`.
* In Visual Studio, go to `Debug` > `Attach to process...` and select Excel

Now you can set breakpoints in the C++ code in VS where code execution will stop when called from Excel via running a UDF.

## Tests

Run `nosetests` from the `xlwings` dir, see also `runtests.py`, but this is also outdated and might be replaced
by something like `tox` again now that numpy/pandas are available as wheels via pypi.
The tests are currently very slow. They were OK with older versions of Excel but they have to be rewritten
to run reasonably fast (i.e. not always open/close the whole workbook).
Also, the tests are standard unittests, so `nose` is not really required to run them.


## Docs

### Build locally

```
cd docs
make html
```

### Build doc translations locally

See https://docs.readthedocs.io/en/stable/guides/manage-translations.html#manage-translations

### Add a translation to the published docs on readthedocs.org

* `.po` files must live under `docs/locales/<language>/LC_MESSAGES`
* Create a new project (`xlwings-<language>`) via generic Git integration as you can only import a project 1x via the GitHub integration.
* Set the language of that project to the language of the translation
* Add this project as `Translation` in the settings of the main `xlwings` project