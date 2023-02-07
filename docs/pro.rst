.. _pro:

Overview :bdg-secondary:`PRO`
=============================

xlwings PRO is `source-available <https://en.wikipedia.org/wiki/Source-available_software>`_ and dual-licensed under one of the following licenses:

* `xlwings PRO License <https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt>`_ (commercial use requires a `paid plan <https://www.xlwings.org/pricing>`_)
* `PolyForm Noncommercial License 1.0.0 <https://polyformproject.org/licenses/noncommercial/1.0.0>`_ (non-commercial use is free)

PRO Features
------------

* :ref:`Ultra Fast File Reader <file_reader>`: Similar to ``pandas.read_excel()`` but 5-25 times faster and you can leverage the convenient xlwings syntax. Works without an Excel installation and therefore on all platforms including Linux.
* :ref:`xlwings Server <remote_interpreter>`: With xlwings Server, you don't need to install Python locally anymore. Instead, run it as a web app on a server. Works with Desktop Excel on Windows and macOS and with Google Sheet and Excel on the web. Runs on all platforms, including Linux, WSL and Docker.
* :ref:`Office.js Add-ins <officejs_addins>`: Build Office.js add-ins.
* :ref:`Embedded code <release>`: Store your Python source code directly in Excel for easy deployment.
* :ref:`xlwings Reports <reports_quickstart>`: A template-based reporting framework, allowing business users to change the layout of the report without having to touch the Python code.
* :ref:`Markdown Formatting <markdown>`: Support for Markdown formatting of text in cells and shapes like e.g., text boxes.
* :ref:`Permissioning <permissioning>`: Control which users can run which Python modules via xlwings.

Paid plans come with additional services like:

* :ref:`1-click Installer <zero_config_installer>`: Easily build your own Python installer including all dependencies---your end users don't need to know anything about Python
* `On-demand video course <https://training.xlwings.org/p/xlwings>`_
* Direct Support

Check out the `paid plans <https://www.xlwings.org/pricing>`_ for more details!

License Key Activation
----------------------

To use xlwings PRO, you need to install a license key on a Terminal/Command Prompt like so::

    xlwings license update -k YOUR_LICENSE_KEY

Make sure to replace ``LICENSE_KEY`` with your personal key (see below). This will store the license key in your ``xlwings.conf`` file (see :ref:`user_config` for where this is on your system). Instead of running this command, you can also store the license key as an environment variable with the name ``XLWINGS_LICENSE_KEY``.

**License Key for Commercial Purpose**:

* To try xlwings PRO for free in a commercial context, request a trial license key: https://www.xlwings.org/trial
* To use xlwings PRO in a commercial context beyond the trial, you need to enroll in a paid plan (they include additional services like support and the ability to create one-click installers): https://www.xlwings.org/pricing

xlwings PRO licenses are developer licenses, are verified offline (i.e., no telemetry/license server involved) and allow royalty-free deployments to unlimited internal end-users and servers for a hassle-free management. External end-users are included with the business plan. Deployments use deploy keys that don't expire but instead are bound to a specific version of xlwings.

**License Key for non-commercial Purpose**:

* To use xlwings PRO for free in a non-commercial context, use the following license key: ``noncommercial`` (Note that you need at least xlwings 0.26.0).
