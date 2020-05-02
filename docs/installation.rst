.. _installation:

Installation
============

Prerequisites
-------------

* xlwings requires an **installation of Excel** and therefore only works on **Windows** and **macOS**. Note that macOS currently does not support UDFs.
* xlwings requires at least Python 3.5.

Installation
------------

xlwings comes pre-installed with `Anaconda <https://www.anaconda.com/distribution>`_. If you are new to Python or have troubles installing xlwings, Anaconda is highly recommended. Otherwise you can also install it manually with pip::

    pip install xlwings

or conda::

    conda install xlwings

Note that the official ``conda`` package might be few releases behind. You can, however, 
use the ``conda-forge`` channel, see: https://anaconda.org/conda-forge/xlwings::

  conda install -c conda-forge xlwings

.. note::
  When you are on macOS and are installing xlwings with ``conda`` (or use the version that comes with Anaconda),
  you'll need to run ``$ xlwings runpython install`` once to enable the ``RunPython`` calls from VBA.

How to activate xlwings PRO
---------------------------

xlwings PRO offers access to additional functionality. All PRO features are marked with xlwings :guilabel:`PRO` in the docs.

.. note::
    To get access to the additional functionality of xlwings PRO, you need a license key and at least xlwings v0.19.0. Everything under the ``xlwings.pro`` subpackage is distributed under a :ref:`commercial license <commercial_license>`. See :ref:`pro` for more details.

To activate the license key, run the following command::

    xlwings license update -k LICENSE_KEY

This will store the license key under your ``xlwings.conf`` file in your home folder. Alternatively, you can also store the license key under an environment variable with the name ``XLWINGS_LICENSE_KEY``.

xlwings PRO requires additionally the ``cryptography`` and ``Jinja2`` packages which come pre-installed with Anaconda. Otherwise, install them via pip or conda.

With pip, you can also run ``pip install "xlwings[pro]"`` which will take care of the extra dependencies for xlwings PRO.

Dependencies
------------

* **Windows**: ``pywin32``, ``comtypes``

* **Mac**: ``psutil``, ``appscript``

The dependencies are automatically installed via ``conda`` or ``pip``.

Optional Dependencies
---------------------

* NumPy
* Pandas
* Matplotlib
* Pillow/PIL
* Flask (for REST API)
* cryptography (for xlwings PRO)
* Jinja2 (for xlwings PRO)

These packages are not required but highly recommended as they play very nicely with xlwings. They are all pre-installed with Anaconda. With pip, you can install xlwings with all optional dependencies as follows::

    pip install xlwings[all]

Add-in
------

Please see :ref:`xlwings_addin` on how to install the xlwings add-in.

.. note::
   The add-in needs to be the same version as the Python package. Make sure to re-install the add-in after upgrading the xlwings package.

Update
------

To update to the latest xlwings version, run the following in a command prompt::

    pip install --upgrade xlwings

or::

    conda update -c conda-forge xlwings

Make sure to keep your version of the Excel add-in in sync with your Python package by running the following (make sure to close Excel first)::

    xlwings addin install

On **macOS only**, additionaly run::

    xlwings runpython install
