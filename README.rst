xlwings - Make Excel fly with Python!
=====================================

xlwings (Open Source)
---------------------

xlwings is a `BSD-licensed <http://opensource.org/licenses/BSD-3-Clause>`_ Python library that makes it easy to call Python from Excel and vice versa:

* **Scripting**: Automate/interact with Excel from Python using a syntax that is close to VBA.
* **Macros**: Replace your messy VBA macros with clean and powerful Python code.
* **UDFs**: Write User Defined Functions (UDFs) in Python (Windows only).

**Numpy arrays** and **Pandas Series/DataFrames** are fully supported. xlwings-powered workbooks are easy to distribute and work
on **Windows** and **macOS**.

xlwings includes all files in the xlwings package except the ``pro`` folder, i.e., the ``xlwings.pro`` subpackage.

xlwings PRO
-----------

xlwings PRO offers additional functionality on top of xlwings (Open Source), including:

* Support for Google Sheets and Excel on the web
* xlwings Reports, the flexible, template-based reporting system
* Easy deployment via embedded code
* See the `full list of PRO features <https://docs.xlwings.org/en/stable/pro.html>`_

xlwings PRO is `source available <https://en.wikipedia.org/wiki/Source-available_software>`_ and dual-licensed under one of the following licenses:

* `PolyForm Noncommercial License 1.0.0 <https://polyformproject.org/licenses/noncommercial/1.0.0>`_ (noncommercial use is free)
* `xlwings PRO License <https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt>`_ (commercial use requires a `paid plan <https://www.xlwings.org/pricing>`_)

**License Key**

To use xlwings PRO, you need to install a license key on a Terminal/Command Prompt like so (alternatively, set the env var ``XLWINGS_LICENSE_KEY``::

    xlwings license update -k YOUR_LICENSE_KEY

See `the docs <https://docs.xlwings.org/en/latest/pro/license_key.html>`_ for more details.

**License key for noncommercial purpose**:

* To use xlwings PRO for free in a noncommercial context, use the following license key: ``noncommercial``.

**License key for commercial purpose**:

* To try xlwings PRO for free in a commercial context, request a trial license key: https://www.xlwings.org/trial
* To use xlwings PRO in a commercial context beyond the trial, you need to enroll in a paid plan (they include additional services like support and the ability to create one-click installers): https://www.xlwings.org/pricing

xlwings PRO licenses are developer licenses, are verified offline (i.e., no telemetry/license server involved) and allow royalty-free deployments to unlimited internal and external end-users and servers for a hassle-free management. Deployments use deploy keys that don't expire but instead are bound to a specific version of xlwings.

Links
-----

* Homepage: https://www.xlwings.org
* Quickstart: https://docs.xlwings.org/en/stable/quickstart.html
* Documentation: https://docs.xlwings.org
* Book (O'Reilly, 2021): https://www.xlwings.org/book
* Video Course: https://training.xlwings.org/p/xlwings
* Source Code: https://github.com/xlwings/xlwings

xltrail
-------

The Excel files are also tracked with `xltrail <https://www.xltrail.com>`_. You can see the diffs
`here <https://app.xltrail.com/#/?path=github.com%2Fxlwings%2Fxlwings.git&branch=main&public=true>`_.
