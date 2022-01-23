.. _installation:

Installation
============

Prerequisites
-------------

* xlwings requires an **installation of Excel** and therefore only works on **Windows** and **macOS**. Note that macOS currently does not support UDFs.
* xlwings requires at least Python 3.7.

Here are the last versions of xlwings to support:

* Python 3.6: 0.25.3
* Python 3.5: 0.19.5
* Python 2.7: 0.16.6

Installation
------------

xlwings comes pre-installed with

* `Anaconda <https://www.anaconda.com/products/individual>`_ (Windows and macOS)
* `WinPython <https://winpython.github.io>`_ (Windows only) Make sure **not** to take the ``dot`` version as this only contains Python.

If you are new to Python or have trouble installing xlwings, one of these distributions is highly recommended. Otherwise, you can also install it manually with pip::

    pip install xlwings

or conda::

    conda install xlwings

Note that the official ``conda`` package might be a few releases behind. You can, however,
use the ``conda-forge`` channel (replace ``install`` with ``upgrade`` if xlwings is already installed)::

  conda install -c conda-forge xlwings

.. note::
  When you are on macOS and are installing xlwings with ``conda`` (or use the version that comes with Anaconda),
  you'll need to run ``$ xlwings runpython install`` once to enable the ``RunPython`` calls from VBA. This is done automatically if you install the addin via ``$ xlwings addin install``.

Add-in
------

To install the add-in, run the following command::

    xlwings addin install

To call Excel from Python, you don't need an add-in. Also, you can use a single file VBA module (*standalone workbook*) instead of the add-in. For more details, see :ref:`xlwings_addin`.

.. note::
   The add-in needs to be the same version as the Python package. Make sure to re-install the add-in after upgrading the xlwings package.

Dependencies
------------

* **Windows**: ``pywin32``

* **Mac**: ``psutil``, ``appscript``

The dependencies are automatically installed via ``conda`` or ``pip``.
If you would like to install xlwings without dependencies, you can run ``pip install xlwings --no-deps`` or set the environment variable ``XLWINGS_NO_DEPS=1`` before running ``pip install xlwings``.

How to activate xlwings PRO
---------------------------

xlwings PRO offers access to :ref:`additional functionality <pro>`. All PRO features are marked with xlwings :guilabel:`PRO` in the docs.

.. note::
    To get access to the additional functionality of xlwings PRO, you need a license key and at least xlwings v0.19.0. Everything under the ``xlwings.pro`` subpackage is distributed under a :ref:`commercial license <commercial_license>`. See :ref:`pro` for more details.

To activate the license key, run the following command::

    xlwings license update -k LICENSE_KEY

Make sure to replace ``LICENSE_KEY`` with your personal key. This will store the license key under your ``xlwings.conf`` file (see :ref:`user_config` for where this is on your system). Alternatively, you can also store the license key as an environment variable with the name ``XLWINGS_LICENSE_KEY``.

xlwings PRO requires additionally the ``cryptography`` and ``Jinja2`` packages, which come preinstalled with Anaconda and WinPython. Otherwise, install them via pip or conda.

With pip, you can also run ``pip install "xlwings[pro]"``: this will take care of the extra dependencies for xlwings PRO.

Optional Dependencies
---------------------

* NumPy
* Pandas
* Matplotlib
* Pillow/PIL
* Flask (for REST API)
* cryptography (for xlwings.pro)
* Jinja2 (for xlwings.pro.reports)
* requests (for permissioning)

These packages are not required but highly recommended as they play very nicely with xlwings. They are all pre-installed with Anaconda. With pip, you can install xlwings with all optional dependencies as follows::

    pip install "xlwings[all]"

Update
------

To update to the latest xlwings version, run the following in a command prompt::

    pip install --upgrade xlwings

or::

    conda update -c conda-forge xlwings

Make sure to keep your version of the Excel add-in in sync with your Python package by running the following (make sure to close Excel first)::

    xlwings addin install

Uninstall
---------

To uninstall xlwings completely, first uninstall the add-in, then uninstall the xlwings package using the same method (pip or conda) that you used for installing it::

    xlwings addin remove

Then ::

    pip uninstall xlwings

or::

    conda remove xlwings

Finally, manually remove the ``.xlwings`` directory in your home folder if it exists.