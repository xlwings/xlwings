# Installation

## Prerequisites

* xlwings (Open Source) requires an **installation of Excel** and therefore only works on **Windows** and **macOS**. Note that macOS currently does not support UDFs.
* xlwings PRO offers additional features:
    * [File Reader](pro/reader.md#xlwings-reader) (new in v0.28.0): Runs additionally on Linux and doesn't require an installation of Excel.
* xlwings requires at least Python 3.9.

Here are previous versions of xlwings that support older versions of Python:

* Python 3.8: 0.31.10
* Python 3.7: 0.30.9
* Python 3.6: 0.25.3
* Python 3.5: 0.19.5
* Python 2.7: 0.16.6

## xlwings Python package

xlwings comes pre-installed with [Anaconda](https://www.anaconda.com/products/individual) (Windows and macOS). Otherwise, you can also install it with pip:

```text
pip install xlwings
```

or conda:

```text
conda install xlwings
```

Note that the official conda package might be a few releases behind. You can, however,
use the `conda-forge` channel (replace `install` with `upgrade` if xlwings is already installed):

```text
conda install -c conda-forge xlwings
```

## xlwings Excel Add-in

To install the add-in, run the following command:

```text
xlwings addin install
```

To automate Excel from Python, you don't need an add-in. Also, you can use a single file VBA module (*standalone workbook*) instead of the add-in. For more details, see [Add-in & Settings](addin.md#add-in--settings).

```{note}
   The add-in needs to be the same version as the Python package. Make sure to run `xlwings add install` again after upgrading the xlwings package.
```

```{note}
  When you are on macOS and are using the VBA standalone module instead of the add-in, you need to run `$ xlwings runpython install` once.
```

## Dependencies

For automating Excel, you'll need the following dependencies:

* **Windows**: `pywin32`

* **Mac**: `psutil`, `appscript`

The dependencies are automatically installed via `conda` or `pip`.
If you would like to install xlwings without dependencies, you can run `pip install xlwings --no-deps`.

## How to activate xlwings PRO

See [xlwings PRO](pro/license_key.md#license-key).

## Optional Dependencies

* NumPy
* pandas
* Matplotlib
* Pillow
* Jinja2 (for xlwings.reports)

These packages are not required but highly recommended as they play very nicely with xlwings. They are all pre-installed with Anaconda. With pip, you can install xlwings with all optional dependencies as follows:

```text
pip install "xlwings[all]"
```

## Update

To update to the latest xlwings version, run the following in a command prompt:

```text
pip install --upgrade xlwings
```

or:

```text
conda update -c conda-forge xlwings
```

Make sure to keep your version of the Excel add-in in sync with your Python package by running the following (make sure to close Excel first):

```text
xlwings addin install
```

```{note}
If you get an `Object required` error with UDFs after an update, re-import the functions and recalculate the workbook via `Ctrl+Alt+F9`.
```

## Uninstall

To uninstall xlwings completely, first uninstall the add-in, then uninstall the xlwings package using the same method (pip or conda) that you used for installing it:

```text
xlwings addin remove
```

Then:

```text
pip uninstall xlwings
```

or:

```text
conda remove xlwings
```

Finally, manually remove the `.xlwings` directory in your home folder if it exists.
