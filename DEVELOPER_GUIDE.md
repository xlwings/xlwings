# Developer Guide

## Python package

The source for the Python package is in the `xlwings` directory.

1. Fork xlwings repository on GitHub into your own account
2. Clone your forked repository: `git clone <your forked git url>`
3. `cd xlwings`
4. With the desired development environment activated: `pip install -e ".[all]"`. This will install xlwings like a standard package
   but runs from your cloned source code, i.e. you can edit/debug the xlwings code. If you don't want the dependencies to be taken care of, you could also use `python setup.py develop`.

## Mac implementation

All mac specific code is in `xlwings/_xlmac.py`. To find out the syntax of a new feature, it sometimes works by just looking at the existing
code and comparing it with the dictionaries exported by `ASDictionary` (see under `resources/mac`).
If that doesn't work, you'll need to find out the corresponding syntax in `AppleScript`, e.g. by looking at `Excel2004AppleScriptRef.pdf`
under `resouces/mac` or by searching the internet. Then use `ASTranslate` to translate it into `appscript` syntax. Unfortunately,
ASTranslate fails on the latest versions of macOS. A workaround is to run it on an old macOS (e.g. OS 10.11 - El Capitan) in a VM.

Links:

[appscript homepage](http://appscript.sourceforge.net/)  
[ASTranslate & ASDictionary download](https://sourceforge.net/projects/appscript/files/)  
[appscript source code](https://github.com/hhas/appscript)

## Addin

Install the addin in Excel by going to `Developer` > `Excel Add-in` > `Browse` and pointing to the addin in the source code,
i.e. under `xlwings/addin/xlwings.xlam`.

To change the buttons, you need to download the Office RibbonX Editor (only runs on Windows) from 
https://github.com/fernandreu/office-ribbonx-editor/releases

The code is pure VBA code.

## dlls

The source for the dlls (used for the UDF related stuff on Windows) is under `src`. It's written in C++.

The dlls are almost never touched, so you should not bother to setup the development environment. What you should do is
download the latest version of xlwings from https://pypi.org, unpack it and place the two dlls next to the Python interpreter,
then rename them into `xlwings32-dev.dll` and `xlwings64-dev.dll`. Note that the bitness refers to the Excel
installation and not to the Python installation!

If you ever need to change the C++ source, then download Visual Studio Community 2019 to open and compile the 
project there. When installing, make sure to select `Custom Installation` so you can activate the checkbox for `Visual C++` under
programming languages.

If you need to debug the dll, in Visual Studio do the following:

* Double-click `xlwings.sln` to open the project in Visual Studio 2019.
* Build the project in `Debug` mode (select `Win32` if your Excel is 32 bit and `x64` if your Excel is 64 bit). Note that the Python bitness does not matter!
* Make sure that Excel calls that dll that was built in debug mode. In the xlwings addin (password: `xlwings`) you could e.g. override the path in `XLPyLoadDLL`.
* In Visual Studio, go to `Debug` > `Attach to process...` and select Excel

Now you can set breakpoints in the C++ code in VS where code execution will stop when called from Excel via running a UDF.

## Tests

Currently, we're migrating to `pytest`, so you'll find a mix between `unittest` and `pytest` format.
Running the whole tests suite is currently broken, so it's recommended to run single modules instead.
See e.g., `test_font.py` for the new style of tests that are also fast.

## Docs

### Build locally

```
pip install sphinx-autobuild
```

```
sphinx-autobuild docs docs/_build/html
```

without autobuild:

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
