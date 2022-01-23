xlwings - Make Excel Fly!
=========================

xlwings is a `BSD-licensed <http://opensource.org/licenses/BSD-3-Clause>`_ Python library that makes it easy to call Python from Excel and vice versa:

* **Scripting**: Automate/interact with Excel from Python using a syntax close to VBA.
* **Macros**: Replace VBA macros with clean and powerful Python code.
* **UDFs**: Write User Defined Functions (UDFs) in Python (Windows only).
* **REST API**: Expose your Excel workbooks via REST API.

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

    .. grid-item-card::  :octicon:`lock;2em;sd-text-success` xlwings PRO
        :link: pro
        :link-type: doc

        xlwings PRO offers additional functionality including xlwings Reports, the template-based reporting system.

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
    remote_interpreter
    reports
    markdown
    release
    permissioning

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: About

    whatsnew
    license

.. toctree::
    :maxdepth: 2
    :hidden:
    :caption: API Reference

    api
    rest_api





