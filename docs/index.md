# xlwings - Make Excel Fly!

xlwings (Open Source) is a [BSD-licensed](http://opensource.org/licenses/BSD-3-Clause) Python library that makes it easy to call Python from Excel and vice versa:

* **Scripting**: Automate/interact with Excel from Python using a syntax close to VBA.
* **Macros**: Replace VBA macros with clean and powerful Python code.
* **UDFs**: Write User Defined Functions (UDFs) in Python (Windows only).

**Numpy arrays** and **Pandas Series/DataFrames** are fully supported. xlwings-powered workbooks are easy to distribute and work
on **Windows** and **Mac**.

::::::{grid} 1 2 2 2
:gutter: 3
:margin: 5 0 0 0

:::::{grid-item-card} {octicon}`rocket;2em;sd-text-success` Getting Started
:link: quickstart
:link-type: doc

Start here if you are new to xlwings. Learn about the syntax, the `RunPython` call, the add-in and UDFs.
:::::

:::::{grid-item-card} {octicon}`light-bulb;2em;sd-text-success` Advanced Features
:link: converters
:link-type: doc

More in-depths explanations about converters, debugging or how to write your own add-in.
:::::

:::::{grid-item-card} {octicon}`star;2em;sd-text-success` xlwings Lite & xlwings Server

To run xlwings without having to install Python, you have two options:

* [xlwings Lite](https://lite.xlwings.org): Install the free add-in from Excel's add-in store and you're done. It comes with a VS Code-like editor and Wingman, a powerful AI assistant.
* [xlwings Server](https://server.xlwings.org): Create your own modern Excel add-in with Python instead of JavaScript. Python runs on your own server and comes with enterprise features such as SSO and RBAC.
:::::

:::::{grid-item-card} {octicon}`code-square;2em;sd-text-success` API Reference
:link: api/index
:link-type: doc

This is a description of all the classes, methods, properties and functions that xlwings offers to work with the Excel object model.
:::::
::::::
```{toctree}
:maxdepth: 2
:hidden:

quickstart
getting_started/index
```

```{toctree}
:maxdepth: 2
:hidden:

advanced_features/index
```

```{toctree}
:maxdepth: 2
:caption: xlwings PRO
:hidden:

pro/license_key
pro/reports/index
pro/reader
pro/release
```

```{toctree}
:maxdepth: 2
:caption: About
:hidden:

whatsnew
license
api/index
```
