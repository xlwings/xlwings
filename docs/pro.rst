.. _pro:

xlwings PRO Overview
====================

xlwings PRO is `source-available <https://en.wikipedia.org/wiki/Source-available_software>`_ and dual-licensed under one of the following licenses:

* `PolyForm Noncommercial License 1.0.0 <https://polyformproject.org/licenses/noncommercial/1.0.0>`_ (noncommercial use is free)
* `xlwings PRO License <https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt>`_ (commercial use requires a `paid plan <https://www.xlwings.org/pricing>`_)

**License Key**

To use xlwings PRO, you need to install a license key on a Terminal/Command Prompt like so::

    xlwings license update -k YOUR_LICENSE_KEY

Make sure to replace ``LICENSE_KEY`` with your personal key (see below). This will store the license key in your ``xlwings.conf`` file (see :ref:`user_config` for where this is on your system). Instead of running this command, you can also store the license key as an environment variable with the name ``XLWINGS_LICENSE_KEY``.

**License key for noncommercial purpose**:

* To use xlwings PRO for free in a noncommercial context, use the following license key: ``noncommercial``.

**License key for commercial purpose**:

* To try xlwings PRO for free in a commercial context, request a trial license key: https://www.xlwings.org/trial
* To use xlwings PRO in a commercial context beyond the trial, you need to enroll in a paid plan (they include additional services like support and the ability to create one-click installers): https://www.xlwings.org/pricing

xlwings PRO licenses are developer licenses, are verified offline (i.e., no telemetry/license server involved) and allow royalty-free deployments to unlimited internal and external end-users and servers for a hassle-free management. Deployments use deploy keys that don't expire but instead are bound to a specific version of xlwings.

xlwings PRO functionality requires additionally the ``cryptography`` package, which comes preinstalled with Anaconda and WinPython. Otherwise, install it via pip or Conda. With pip, you can also run ``pip install "xlwings[pro]"``: this will take care of the extra dependencies for xlwings PRO.

PRO Features
------------

* :ref:`Remote Interpreter <remote_interpreter>`: Work with Google Sheets and Excel on the web and a remote Python interpreter.
* :ref:`Embedded code <release>`: Store your Python source code directly in Excel for easy deployment.
* :ref:`reports_quickstart`: A template-based reporting mechanism, allowing business users to change the layout of the report without having to touch the Python code.
* :ref:`markdown`: Support for Markdown formatting of text in cells and shapes like e.g., text boxes.
* :ref:`permissioning`: Control which users can run which Python modules via xlwings.

Paid plans come with additional services like:

* :ref:`One-click Installer <zero_config_installer>`: Easily build your own Python installer including all dependencies---your end users don't need to know anything about Python
* `On-demand video course <https://training.xlwings.org/p/xlwings>`_
* Direct Support

Check out the `paid plans <https://www.xlwings.org/pricing>`_ for more details!
