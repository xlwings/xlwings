.. _installation:

Installation
============

The easiest way to install xlwings is via pip::

    pip install xlwings

or conda::

    conda install xlwings

Note that the official ``conda`` package might be few releases behind. You can, however, 
use the ``conda-forge`` channel (see: https://anaconda.org/conda-forge/xlwings) which should usually be up to date (but might still be a day or so behind the pip release)::

  conda install -c conda-forge xlwings

.. note::
  When you are using Mac Excel 2016 and are installing xlwings with ``conda`` (or use the version that comes with Anaconda),
  you'll need to run ``$ xlwings runpython install`` once to enable the ``RunPython`` calls from VBA. Alternatively, you can simply
  install xlwings with ``pip``.

Dependencies
------------

* **Windows**: ``pywin32``, ``comtypes``

  On Windows, the dependencies are automatically being handled if xlwings is installed with ``conda`` or ``pip``.

* **Mac**: ``psutil``, ``appscript``

  On Mac, the dependencies are automatically being handled if xlwings is installed with ``conda`` or ``pip``. However,
  with pip, the Xcode command line tools need to be available. Mac OS X 10.4 (*Tiger*) or later is required.
  The recommended Python distribution for Mac is `Anaconda <https://www.anaconda.com/distribution>`_. With ``conda``
  on the other hand, you'll need to manually run the command ``xlwings runpython install``.

Optional Dependencies
---------------------

* NumPy
* Pandas
* Matplotlib
* Pillow/PIL
* Flask (for REST API only)

These packages are not required but highly recommended as they play very nicely with xlwings.

Add-in
------

Please see :ref:`xlwings_addin` on how to install the xlwings add-in.

Python version support
----------------------

xlwings is tested on Python 2.7 and 3.3+