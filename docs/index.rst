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
        :link: converters
        :link-type: doc

        More in-depths explanations about converters, debugging or how to write your own add-in.

    .. grid-item-card::  :octicon:`star;2em;sd-text-success` xlwings :bdg-secondary:`PRO`
        :link: pro
        :link-type: doc

        Use advanced features such as:

        * Ultra fast file reader: no Excel required
        * xlwings Reports: work with templates
        * xlwings Server: no local Python required
        * Google Sheets & Excel on the web
        * Embedded code: easy deployment
        * etc.

        Free for non-commercial use only.

    .. grid-item-card::  :octicon:`code-square;2em;sd-text-success` API Reference
        :link: api
        :link-type: doc

        This is a description of all the classes, methods, properties and functions that xlwings offers to work with the Excel object model.

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: Getting Started

    course
    installation
    quickstart
    connect_to_workbook
    syntax_overview
    datastructures
    addin
    vba
    udfs
    matplotlib
    jupyternotebooks
    command_line
    deployment
    onedrive_sharepoint
    troubleshooting

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: Advanced Features

    converters
    debugging
    extensions
    customaddin
    threading_and_multiprocessing
    missing_features
    other_office_apps

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: xlwings PRO

    pro
    release
    permissioning

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: File Reader

    reader

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: xlwings Reports

    reports
    markdown

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: xlwings Server

    remote_interpreter
    officejs_addins
    server_authentication


.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: About

    whatsnew
    license
    opensource_licenses

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: API Reference

    api
    rest_api
