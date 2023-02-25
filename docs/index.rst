xlwings - Make Excel Fly!
=========================

xlwings (Open Source) is a `BSD-licensed <http://opensource.org/licenses/BSD-3-Clause>`_ Python library that makes it easy to call Python from Excel and vice versa:

* **Scripting**: Automate/interact with Excel from Python using a syntax close to VBA.
* **Macros**: Replace VBA macros with clean and powerful Python code.
* **UDFs**: Write User Defined Functions (UDFs) in Python (Windows only).

**Numpy arrays** and **Pandas Series/DataFrames** are fully supported. xlwings-powered workbooks are easy to distribute and work
on **Windows** and **Mac**.

.. grid:: 2
    :margin: 5 0 0 0

    .. grid-item-card::  :octicon:`rocket;2em;sd-text-success` Getting Started
        :link: quickstart
        :link-type: doc

        Start here if you are new to xlwings. Learn about the syntax, the ``RunPython`` call, the add-in and UDFs.

    .. grid-item-card::  :octicon:`light-bulb;2em;sd-text-success` Advanced Features
        :link: advanced_features/converters
        :link-type: doc

        More in-depths explanations about converters, debugging or how to write your own add-in.

    .. grid-item-card::  :octicon:`star;2em;sd-text-success` xlwings PRO
        :link: pro/license_key
        :link-type: doc

        Use advanced features such as:

        * 1-click installer: bundle Python and all your packages
        * Embedded code: easy deployment
        * Ultra fast file reader: no Excel required
        * xlwings Reports: work with templates
        * xlwings Server: no local Python required
        * No more VBA: Call Python from Office Scripts and Office.js
        * Excel on the web & Google Sheets

        Free for non-commercial use only.

    .. grid-item-card::  :octicon:`code-square;2em;sd-text-success` API Reference
        :link: api/index
        :link-type: doc

        This is a description of all the classes, methods, properties and functions that xlwings offers to work with the Excel object model.

.. toctree::
    :maxdepth: 2
    :hidden:

    quickstart
    getting_started/index

.. toctree::
    :maxdepth: 2
    :hidden:

    advanced_features/index

.. toctree::
    :maxdepth: 2
    :caption: xlwings PRO
    :hidden:

    pro/license_key
    pro/server/index
    pro/reports/index
    pro/reader
    pro/release

.. toctree::
    :maxdepth: 2
    :caption: About
    :hidden:

    whatsnew
    license
    api/index
