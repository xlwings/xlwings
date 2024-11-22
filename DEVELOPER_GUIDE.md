# Developer Guide

## Python package

The source for the Python package is in the `xlwings` directory.

1. Fork xlwings repository on GitHub into your own account
2. Clone your forked repository: `git clone <your forked git url>`
3. `cd xlwings`
4. With the desired development environment activated: `pip install -e ".[all]"`. This will install xlwings like a standard package
   but runs from your cloned source code, i.e. you can edit/debug the xlwings code. If you don't want the dependencies to be taken care of or don't want to install Rust (used by xlwings Reader), you can use `python setup.py develop` (this requires you to temporarily remove `pyproject.toml` though).

## macOS

All macOS specific code is in `xlwings/_xlmac.py`. To find out the syntax of a new feature, it sometimes works by just looking at the existing
code and comparing it with the dictionaries exported by `ASDictionary` (see under `resources/mac`).
If that doesn't work, you'll need to find out the corresponding syntax in `AppleScript`, e.g. by looking at `Excel2004AppleScriptRef.pdf`
under `resouces/mac` or by searching the internet. Then use `ASTranslate` to translate it into `appscript` syntax.

FWIW, Apple has added support for JavaScript in the Script Editor in addition to AppleScript and the syntax is very close to `appscript`:

```js
Application('Microsoft Excel').worksheets['Sheet1'].rows[0].columns[0].value.get()
Application('Microsoft Excel').worksheets['Sheet1'].rows[0].columns[0].properties.get()
```

Links:

[appscript homepage](http://appscript.sourceforge.net/)  
[ASTranslate & ASDictionary download](https://sourceforge.net/projects/appscript/files/)  
[appscript source code](https://github.com/hhas/appscript)

## Excel add-in

* Install the addin in Excel by going to `Developer` > `Excel Add-in` > `Browse` and pointing to the addin in the source code,
i.e. under `xlwings/addin/xlwings.xlam`.

* Configure the addin manually (by setting `Interpreter` or `Conda Path`/`Conda Env` settings) or run `xlwings config create --force` to create a new `xlwings.conf` file (this is where the ribbon will read/write the configuration to). 

To change the ribbon UI, you need to download the Office RibbonX Editor (only runs on Windows) from 
https://github.com/fernandreu/office-ribbonx-editor/releases

The code is pure VBA code. The suggested way to edit the VBA code is:

1. In Excel, make sure that `Trust access` to the VBA project object model is enabled under `File` > `Options` > `Trust Center` > `Trust Center Settings` > `Macro Settings`.
1. Open the VBA editor via `Alt+F11`, then click on the source code of xlwings and unlock it with the password `xlwings`.
2. Run the following on a command prompt: 

   ```
   cd xlwings/addin
   xlwings vba edit -f xlwings.xlam
   ```
   
   Confirm the prompt.


3. Make changes to the source code under `xlwings/addin` in an external editor: the changes are synced automatically to the VBA editor.



## Windows dlls

These are used for the UDFs on Windows. The source for the dlls is under `xlwingsdll`. It's written in C++.

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

## Rust extension

This is used for the file reader. The source is under `src` together with various Cargo files/dirs in the root directory.

* Requires Rust: https://rustup.rs
* `pip install maturin`
* Run the following for local development:
  `maturin develop` or `maturin develop --release` for optimized builds
  This will build and install the extension in the current environment.

The 3rd party Open Source licenses document is built with `cargo about generate about.hbs > docs/_static/opensource_licenses2.html` this requires `cargo install --locked cargo-about`.

## Code formatting/linting

This repo uses the following packages for code formatting/linting, see `pyproject.toml`:

* black
* ruff

You can use the pre-commit hook under `.pre-commit-config.yaml`, see instructions at top of the file.

## Tests

Currently, we're migrating to `pytest`, so you'll find a mix between `unittest` and `pytest` format.
Running the whole tests suite is currently broken, so it's recommended to run single modules instead.
See e.g., `test_font.py` for the new style of tests that are also fast.

For running the xlwings pro related tests, you'll need to use the `noncommercial` license key, see: [Activate a developer key](https://docs.xlwings.org/en/latest/pro/license_key.html#activate-a-developer-key).

To run the UDF tests, open `udf_tests.xlsm` in the `tests/udfs` directory and follow the instructions at the top of the file. The other files work accordingly.

## Docs

### Build locally

```
pip install -r docs/requirements.txt
pip install sphinx-autobuild
```

```
sphinx-autobuild docs docs/_build/html --watch ./xlwings
```

without autobuild:

```
cd docs
make html
```

To double-check the Sphinx warnings, it's best to run it as follows:

```
cd docs
clear && make clean html
```

### Build doc translations locally

See https://docs.readthedocs.io/en/stable/guides/manage-translations.html#manage-translations

### Add a translation to the published docs on readthedocs.org

* `.po` files must live under `docs/locales/<language>/LC_MESSAGES`
* Create a new project (`xlwings-<language>`) via generic Git integration as you can only import a project 1x via the GitHub integration.
* Set the language of that project to the language of the translation
* Add this project as `Translation` in the settings of the main `xlwings` project
