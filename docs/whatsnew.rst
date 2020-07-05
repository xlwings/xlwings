What's New
==========

v0.19.5 (Jul 5, 2020)
----------------------

* [Enhancement] When you install the add-in via ``xlwings addin install``, it autoconfigures the add-in if it can't find an existing user config file (:issue:`1322`).
* [Feature] New ``xlwings config create [--force]`` command that autogenerates the user config file with the Python settings from which you run the command. Can be used to reset the add-in settings with the ``--force`` option (:issue:`1322`).
* [Feature]: There is a new option to show/hide the console window. Note that with ``Conda Path`` and ``Conda Env`` set, the console always pops up when using the UDF server. Currently only available on Windows (:issue:`1182`).
* [Enhancement] The ``Interpreter`` setting has been deprecated in favor of platform-specific settings: ``Interpreter_Win`` and ``Interpreter_Mac``, respectively. This allows you use the sheet config unchanged on both platforms (:issue:`1345`).
* [Enhancement] On macOS, you can now use a few environment-like variables in your settings: ``$HOME``, ``$APPLICATIONS``, ``$DOCUMENTS``, ``$DESKTOP`` (:issue:`615`).
* [Bug Fix]: Async functions sometimes caused an error on older Excel versions without dynamic arrays (:issue:`1341`).

v0.19.4 (May 20, 2020)
----------------------

* [Feature] ``xlwings addin install`` is now available on macOS. On Windows, it has been fixed so it should now work reliably (:issue:`704`).
* [Bug Fix] Fixed a ``dll load failed`` issue with ``pywin32`` when installed via ``pip`` on Python 3.8 (:issue:`1315`).

v0.19.3 (May 19, 2020)
----------------------

* :guilabel:`PRO` [Feature]: Added possibility to create deployment keys, see :ref:`deployment_key`.

v0.19.2 (May 11, 2020)
----------------------

* [Feature] New methods :meth:`xlwings.Shape.scale_height` and :meth:`xlwings.Shape.scale_width` (:issue:`311`).
* [Bug Fix] Using ``Pictures.add`` is not distorting the proportions anymore (:issue:`311`).

* :guilabel:`PRO` [Feature]: Added support for :ref:`plotly` (:issue:`1309`).

.. figure:: images/plotly.png
    :scale: 40%

v0.19.1 (May 4, 2020)
---------------------

* [Bug Fix] Fixed an issue with the xlwings PRO license key when there was no ``xlwings.conf`` file (:issue:`1308`).

v0.19.0 (May 2, 2020)
---------------------

* [Bug Fix] Native dynamic array formulas can now be used with async formulas (:issue:`1277`)
* [Enhancement] Quickstart references the project's name when run from Python instead of the active book (:issue:`1307`)

**Breaking Change**:

* ``Conda Base`` has been renamed into ``Conda Path`` to reduce the confusion with the ``Conda Env`` called ``base``. Please adjust your settings accordingly! (:issue:`1194`)

v0.18.0 (Feb 15, 2020)
----------------------

* [Feature] Added support for merged cells: :attr:`xlwings.Range.merge_area`, :attr:`xlwings.Range.merge_cells`, :meth:`xlwings.Range.merge`
  :meth:`xlwings.Range.unmerge` (:issue:`21`).
* [Bug Fix] ``RunPython`` now works properly with files that have a URL as ``fullname``, i.e. OneDrive and SharePoint (:issue:`1253`).
* [Bug Fix] Fixed a bug with ``wb.names['...'].refers_to_range`` on macOS (:issue:`1256`).


v0.17.1 (Jan 31, 2020)
----------------------

* [Bug Fix] Handle ``np.float64('nan')`` correctly (:issue:`1116`).

v0.17.0 (Jan 6, 2020)
---------------------

This release drops support for Python 2.7 in xlwings CE. If you still rely on Python 2.7, you will need to stick to v0.16.6.

v0.16.6 (Jan 5, 2020)
---------------------

* [Enhancement] CLI changes with respect to ``xlwings license`` (:issue:`1227`). 

v0.16.5 (Dec 30, 2019)
----------------------

* [Enhancement] Improvements with regards to the ``Run main`` ribbon button (:issue:`1207` and :issue:`1222`).

v0.16.4 (Dec 17, 2019)
----------------------

* [Enhancement] Added support for :meth:`xlwings.Range.copy` (:issue:`1214`).
* [Enhancement] Added support for :meth:`xlwings.Range.paste` (:issue:`1215`). 
* [Enhancement] Added support for :meth:`xlwings.Range.insert` (:issue:`80`).
* [Enhancement] Added support for :meth:`xlwings.Range.delete` (:issue:`862`).

v0.16.3 (Dec 12, 2019)
----------------------

* [Bug Fix] Sometimes, xlwings would show an error of a previous run. Moreover, 0.16.2 introduced an issue that would
  not show errors at all on non-conda setups (:issue:`1158` and :issue:`1206`)
* [Enhancement] The xlwings CLI now prints the version number (:issue:`1200`)

**Breaking Change**:

* ``LOG FILE`` has been retired and removed from the configuration/add-in.

v0.16.2 (Dec 5, 2019)
---------------------

* [Bug Fix] ``RunPython`` can now be called in parallel from different Excel instances (:issue:`1196`).

v0.16.1 (Dec 1, 2019)
---------------------

* [Enhancement] :meth:`xlwings.Book()` and ``myapp.books.open()`` now accept parameters like 
  ``update_links``, ``password`` etc. (:issue:`1189`).
* [Bug Fix] ``Conda Env`` now works correctly with ``base`` for UDFs, too (:issue:`1110`).
* [Bug Fix] ``Conda Base`` now allows spaces in the path (:issue:`1176`).
* [Enhacement] The UDF server timeout has been increased to 2 minutes (:issue:`1168`).


v0.16.0 (Oct 13, 2019)
----------------------

This release adds a small but very powerful feature: There's a new ``Run main`` button in the add-in.
With that, you can run your Python scripts from standard ``xlsx`` files - no need to save your workbook
as macro-enabled anymore! 

The only condition to make that work is that your Python script has the same name as your workbook and that it contains
a function called ``main``, which will be called when you click the ``Run`` button. All settings from your config file or
config sheet are still respected, so this will work even if you have the source file in a different directory
than your workbook (as long as that directory is added to the ``PYTHONPATH`` in your config).

The ``xlwings quickstart myproject`` has been updated accordingly. It still produces an ``xlsm`` file at the moment
but you can save it as ``xlsx`` file if you intend to run it via the new ``Run`` button.

    .. figure:: images/ribbon.png
        :scale: 40%

v0.15.10 (Aug 31, 2019)
-----------------------

* [Bug Fix] Fixed a Python 2.7 incompatibility introduced with 0.15.9.

v0.15.9 (Aug 31, 2019)
----------------------

* [Enhancement] The ``sql`` extension now uses the native dynamic arrays if available (:issue:`1138`).
* [Enhancement] xlwings now support ``Path`` objects from ``pathlib`` for all file paths (:issue:`1126`).
* [Bug Fix] Various bug fixes: (:issue:`1118`), (:issue:`1131`), (:issue:`1102`).

v0.15.8 (May 5, 2019)
---------------------

* [Bug Fix] Fixed an issue introduced with the previous release that always showed the command prompt when running UDFs,
  not just when using conda envs (:issue:`1098`).

v0.15.7 (May 5, 2019)
---------------------

* [Bug Fix] ``Conda Base`` and ``Conda Env`` weren't stored correctly in the config file from the ribbon (:issue:`1090`).
* [Bug Fix] UDFs now work correctly with ``Conda Base`` and ``Conda Env``. Note, however, that currently there is no
  way to hide the command prompt in that configuration (:issue:`1090`).
* [Enhancement] ``Restart UDF Server`` now actually does what it says: it stops and restarts the server. Previously
  it was only stopping the server and only when the first call to Python was made, it was started again (:issue:`1096`).

v0.15.6 (Apr 29, 2019)
----------------------

* [Feature] New default converter for ``OrderedDict`` (:issue:`1068`).
* [Enhancement] ``Import Functions`` now restarts the UDF server to guarantee a clean state after importing. (:issue:`1092`)
* [Enhancement] The ribbon now shows tooltips on Windows (:issue:`1093`)
* [Bug Fix] RunPython now properly supports conda environments on Windows (they started to require proper activation
  with packages like numpy etc). Conda >=4.6. required. A fix for UDFs is still pending (:issue:`954`).

**Breaking Change:**

* [Bug Fix] ``RunFronzenPython`` now accepts spaces in the path of the executable, but in turn requires to be called
  with command line arguments as a separate VBA argument.
  Example: ``RunFrozenPython "C:\path\to\frozen_executable.exe", "arg1 arg2"`` (:issue:`1063`).

v0.15.5 (Mar 25, 2019)
----------------------

* [Enhancement] ``wb.macro()`` now accepts xlwings objects as arguments such as ``range``, ``sheet`` etc. when the VBA macro expects the corresponding Excel object (e.g. ``Range``, ``Worksheet`` etc.) (:issue:`784` and :issue:`1084`)

**Breaking Change:**

* Cells that contain a cell error such as ``#DIV/0!``, ``#N/A``, ``#NAME?``, ``#NULL!``, ``#NUM!``, ``#REF!``, ``#VALUE!`` return now 
  ``None`` as value in Python. Previously they were returned as constant on Windows (e.g. ``-2146826246``) or ``k.missing_value`` on Mac.


v0.15.4 (Mar 17, 2019)
----------------------

* [Win] BugFix: The ribbon was not showing up in Excel 2007. (:issue:`1039`)
* Enhancement: Allow to install xlwings on Linux even though it's not a supported platform: ``export INSTALL_ON_LINUX=1; pip install xlwings`` (:issue:`1052`)


v0.15.3 (Feb 23, 2019)
----------------------

Bug Fix release:

* [Mac] `RunPython` was broken by the previous release. If you install via ``conda``, make sure to run ``xlwings runpython install`` again! (:issue:`1035`)
* [Win] Sometimes, the ribbon was throwing errors (:issue:`1041`)

v0.15.2 (Feb 3, 2019)
---------------------

Better support and docs for deployment, see :ref:`deployment`:

* You can now package your python modules into a zip file for easier distribution (:issue:`1016`).
* ``RunFrozenPython`` now allows to includes arguments, e.g. ``RunFrozenPython "C:\path\to\my.exe arg1 arg2"`` (:issue:`588`).

**Breaking changes**:

* Accessing a not existing PID in the ``apps`` collection raises now a ``KeyError`` instead of an ``Exception`` (:issue:`1002`).

v0.15.1 (Nov 29, 2018)
----------------------

Bug Fix release:

* [Win] Calling Subs or UDFs from VBA was causing an error (:issue:`998`).

v0.15.0 (Nov 20, 2018)
----------------------

**Dynamic Array Refactor**

While we're all waiting for the new native dynamic arrays, it's still going to take another while until the
majority can use them (they are not yet part of Office 2019).

In the meantime, this refactor improves the current xlwings dynamic arrays in the following way:

* Use of native ("legacy") array formulas instead of having a normal formula in the top left cell and writing around it
* It's up to 2x faster
* There's no empty row/col required outside of the dynamic array anymore
* It continues to overwrite existing cells (no change there)
* There's a small breaking change in the unlikely case that you were assigning values with the expand option:
  ``myrange.options(expand='table').value = [['b'] * 3] * 3``. This was previously clearing contiguous cells to
  the right and bottom (or one of them depending on the option), now you have to do that explicitly.

**Bug Fixes**:

* Importing multiple UDF modules has been fixed (:issue:`991`).

v0.14.1 (Nov 9, 2018)
---------------------

This is a bug fix release:

* [Win] Fixed an issue when the new ``async_mode`` was used together with numpy arrays (:issue:`984`)
* [Mac] Fixed an issue with multiple arguments in ``RunPython`` (:issue:`905`)
* [Mac] Fixed an issue with the config file (:issue:`982`)

v0.14.0 (Nov 5, 2018)
---------------------

**Features**:

This release adds support for asynchronous functions (like all UDF related functionality, this is only available on Windows).
Making a function asynchronous is as easy as::

    import xlwings as xw
    import time

    @xw.func(async_mode='threading')
    def myfunction(a):
        time.sleep(5)  # long running tasks
        return a

See :ref:`async_functions` for the full docs.

**Bug Fixes**:

* See :issue:`970` and :issue:`973`.


v0.13.0 (Oct 22, 2018)
----------------------

**Features**:

This release adds a REST API server to xlwings, allowing you to easily expose your workbook over the internet,
see :ref:`rest_api` for all the details!

**Enhancements**:

* Dynamic arrays are now more robust. Before, they often didn't manage to write everything when there was a lot going on in the workbook (:issue:`880`)
* Jagged arrays (lists of lists where not all rows are of equal length) now raise an error (:issue:`942`)
* xlwings can now be used with threading, see the docs: :ref:`threading` (:issue:`759`).
* [Win] xlwings now enforces pywin32 224 when installing xlwings on Python 3.7 (:issue:`959`)
* New :any:`xlwings.Sheet.used_range` property (:issue:`112`)

**Bug Fixes**:

* The current directory is now inserted in front of everything else on the PYTHONPATH (:issue:`958`)
* The standalone files had an issue in the VBA module (:issue:`960`)

**Breaking changes**:

* Members of the ``xw.apps`` collection are now accessed by key (=PID) instead of index, e.g.:
  ``xw.apps[12345]`` instead of ``xw.apps[0]``. The apps collection also has a new ``xw.apps.keys()`` method. (:issue:`951`)

v0.12.1 (Oct 7, 2018)
---------------------

[Py27] Bug Fix for a Python 2.7 glitch. 

v0.12.0 (Oct 7, 2018)
---------------------

**Features**:

This release adds support to call Python functions from VBA in all Office apps (e.g. Access, Outlook etc.), not just Excel. As
this uses UDFs, it is only available on Windows.
See the docs: :ref:`other_office_apps`. 


**Breaking changes**:

Previously, Python functions were always returning 2d arrays when called from VBA, no matter whether it was actually a 2d array or not.
Now you get the proper dimensionality which makes it easier if the return value is e.g. a string or scalar as you don't have to
unpack it anymore.

Consider the following example using the VBA Editor's Immediate Window after importing UDFs from a project created
using by ``xlwings quickstart``:

**Old behaviour** ::

    ?TypeName(hello("xlwings"))
    Variant()
    ?hello("xlwings")(0,0)
    hello xlwings

**New behaviour** ::

    ?TypeName(hello("xlwings"))
    String
    ?hello("xlwings")
    hello xlwings

**Bug Fixes**:

* [Win] Support expansion of environment variables in config values (:issue:`615`)
* Other bug fixes: :issue:`889`, :issue:`939`, :issue:`940`, :issue:`943`.

v0.11.8 (May 13, 2018)
----------------------

* [Win] pywin32 is now automatically installed when using pip (:issue:`827`)
* `xlwings.bas` has been readded to the python package. This facilitates e.g. the use of xlwings within other addins (:issue:`857`)

v0.11.7 (Feb 5, 2018)
----------------------

* [Win] This release fixes a bug introduced with v0.11.6 that would't allow to open workbooks by name (:issue:`804`)

v0.11.6 (Jan 27, 2018)
----------------------

Bug Fixes:

* [Win] When constantly writing to a spreadsheet, xlwings now correctly resumes after clicking into cells, previously it was crashing. (:issue:`587`)
* Options are now correctly applied when writing to a sheet (:issue:`798`)


v0.11.5 (Jan 7, 2018)
---------------------

This is mostly a bug fix release:

* Config files can now additionally be saved in the directory of the workbooks, overriding the global Ribbon config, see :ref:`config_file` (:issue:`772`)
* Reading Pandas DataFrames with a simple index was creating a MultiIndex with Pandas > 0.20 (:issue:`786`)
* [Win] The xlwings dlls are now properly versioned, allowing to use pre 0.11 releases in parallel with >0.11 releases (:issue:`743`)
* [Mac] Sheet.names.add() was always adding the names on workbook level (:issue:`771`)
* [Mac] UDF decorators now don't cause errors on Mac anymore (:issue:`780`)

v0.11.4 (Jul 23, 2017)
----------------------

This release brings further improvements with regards to the add-in:

* The add-in now shows the version on the ribbon. This makes it easy to check if you are using the correct version (:issue:`724`):

    .. figure:: images/addin_version.png
        :scale: 80%

* [Mac] On Mac Excel 2016, the ribbon now only shows the available functionality (:issue:`723`):

    .. figure:: images/mac_ribbon.png
        :scale: 80%

* [Mac] Mac Excel 2011 is now supported again with the new add-in. However, since Excel 2011 doesn't support the ribbon, 
  the config file has be created/edited manually, see :ref:`config_file` (:issue:`714`).

Also, some new docs:

* [Win] How to use imported functions in VBA, see :ref:`call_udfs_from_vba`.
* For more up-to-date installations via conda, use the ``conda-forge`` channel, see :ref:`installation`.
* A troubleshooting section: :ref:`troubleshooting`.

v0.11.3 (Jul 14, 2017)
----------------------

* Bug Fix: When using the ``xlwings.conf`` sheet, there was a subscript out of range error (:issue:`708`)
* Enhancement: The add-in is now password protected (pw: ``xlwings``) to declutter the VBA editor (:issue:`710`)

You need to update your xlwings add-in to get the fixes!


v0.11.2 (Jul 6, 2017)
---------------------

* Bug Fix: The sql extension was sometimes not correctly assigning the table aliases (:issue:`699`)
* Bug Fix: Permission errors during pip installation should be resolved now (:issue:`693`)


v0.11.1 (Jul 5, 2017)
---------------------

* Bug Fix: The sql extension installs now correctly (:issue:`695`)
* Added migration guide for v0.11, see :ref:`migrate_to_0.11`

v0.11.0 (Jul 2, 2017)
---------------------

Big news! This release adds a full blown **add-in**! We also throw in a great **In-Excel SQL Extension** and a few **bug fixes**:

Add-in
******

.. figure:: images/ribbon.png
    :scale: 80%

A few highlights:

* Settings don't have to be manipulated in VBA code anymore, but can be either set globally via Ribbon/config file or
  for the workbook via a special worksheet
* UDF server can be restarted directly from the add-in
* You can still use a VBA module instead of the add-in, but the recommended way is the add-in
* Get all the details here: :ref:`xlwings_addin`

In-Excel SQL Extension
**********************

The add-in can be extended with own code. We throw in an ``sql`` function, that allows you to perform SQL queries
on data in your spreadsheets. It's pretty awesome, get the details here: :ref:`extensions`.

Bug Fixes
*********

* [Win]: Running ``Debug > Compile`` is not throwing errors anymore (:issue:`678`)
* Pandas deprecation warnings have been fixed (:issue:`675` and :issue:`664`)
* [Mac]: Errors are again shown correctly in a pop up (:issue:`660`)
* [Mac]: Like Windows, Mac now also only shows errors in a popup. Before it was including stdout, too (:issue:`666`) 

Breaking Changes
****************

* ``RunFrozenPython`` now requires the full path to the executable.
* The xlwings CLI ``xlwings template`` functionality has been removed. Use ``quickstart`` instead.


.. _migrate_to_0.11:

Migrate to v0.11 (Add-in)
-------------------------

This migration guide shows you how you can start using the new xlwings add-in as opposed to the old xlwings VBA module
(and the old add-in that consisted of just a single import button).

Upgrade the xlwings Python package
**********************************

1. Check where xlwings is currently installed

    >>> import xlwings
    >>> xlwings.__path__

2. If you installed xlwings with pip, for once, you should first uninstall xlwings: ``pip uninstall xlwings``
3. Check the directory that you got under 1): if there are any files left over, delete the ``xlwings`` folder and the
   remaining files manually
4. Install the latest xlwings version: ``pip install xlwings``
5. Verify that you have >= 0.11 by doing

    >>> import xlwings
    >>> xlwings.__version__

Install the add-in
******************

1. If you have the old xlwings addin installed, find the location and remove it or overwrite it with the new version (see next step).
   If you installed it via the xlwings command line client, you should be able to do: ``xlwings addin remove``.
2. Close Excel. Run ``xlwings addin install`` from a command prompt. Reopen Excel and check if the xlwings Ribbon
   appears. If not, copy ``xlwings.xlam`` (from your xlwings installation folder under ``addin\xlwings.xlam`` manually
   into the ``XLSTART`` folder.
   You can find the location of this folder under Options > Trust Center > Trust Center Settings... > Trusted Locations,
   under the description ``Excel default location: User StartUp``. Restart Excel and you should see the add-in.


Upgrade existing workbooks
**************************

1. Make a backup of your Excel file
2. Open the file and go to the VBA Editor (``Alt-F11``)
3. Remove the xlwings VBA module
4. Add a reference to the xlwings addin, see :ref:`addin_installation`
5. If you want to use workbook specific settings, add a sheet ``xlwings.conf``, see :ref:`addin_wb_settings`


**Note**: To import UDFs, you need to have the reference to the xlwings add-in set!


v0.10.4 (Feb 19, 2017)
----------------------

* [Win] Bug Fix: v0.10.3 introduced a bug that imported UDFs by default with `volatile=True`, this has now been fixed.
  You will need to reimport your functions after upgrading the xlwings package.

v0.10.3 (Jan 28, 2017)
----------------------

This release adds new features to User Defined Functions (UDFs):

* categories
* volatile option
* suppress calculation in function wizard

Syntax:

.. code-block:: python

    import xlwings as xw
    @xw.func(category="xlwings", volatile=False, call_in_wizard=True)
    def myfunction():
        return ...

For details, check out the (also new) and comprehensive API docs about the decorators: :ref:`udf_api`

v0.10.2 (Dec 31, 2016)
----------------------

* [Win] Python 3.6 is now supported (:issue:`592`)


v0.10.1 (Dec 5, 2016)
---------------------

* Writing a Pandas Series with a MultiIndex header was not writing out the header (:issue:`572`)
* [Win] Docstrings for UDF arguments are now working (:issue:`367`)
* [Mac] ``Range.clear_contents()`` has been fixed (it was doing ``clear()`` instead) (:issue:`576`)
* ``xw.Book(...)`` and ``xw.books.open(...)`` raise now the same error in case the file doesn't exist (:issue:`540`)

v0.10.0 (Sep 20, 2016)
----------------------

Dynamic Array Formulas
**********************

This release adds an often requested & powerful new feature to User Defined Functions (UDFs): Dynamic expansion for
array formulas. While Excel offers array formulas, you need to specify their dimensions up front by selecting the
result array first, then entering the formula and finally hitting ``Ctrl-Shift-Enter``. While this makes sense from
a data integrity point of view, in practice, it often turns out to be a cumbersome limitation, especially when working
with dynamic arrays such as time series data.

This is a simple example that demonstrates the syntax and effect of UDF expansion:

.. code-block:: python

    import numpy as np

    @xw.func
    @xw.ret(expand='table')
    def dynamic_array(r, c):
        return np.random.randn(int(r), int(c))

.. figure:: images/dynamic_array1.png
  :scale: 40%

.. figure:: images/dynamic_array2.png
  :scale: 40%

**Note**: Expanding array formulas will overwrite cells without prompting and leave an empty border around them, i.e.
they will clear the row to the bottom and the column to the right of the array.

Bug Fixes
*********

* The ``int`` converter works now always as you would expect (e.g.: ``xw.Range('A1').options(numbers=int).value``). Before,
  it could happen that the number was off by 1 due to floating point issues (:issue:`554`).

v0.9.3 (Aug 22, 2016)
---------------------

* [Win] ``App.visible`` wasn't behaving correctly (:issue:`551`).
* [Mac] Added support for the new 64bit version of Excel 2016 on Mac (:issue:`549`).
* Unicode book names are again supported (:issue:`546`).
* :meth:`xlwings.Book.save()` now supports relative paths. Also, when saving an existing book under a new name
  without specifying the full path, it'll be saved in Python's current working directory instead of in Excel's default
  directory (:issue:`185`).

v0.9.2 (Aug 8, 2016)
--------------------

Another round of bug fixes:

* [Mac]: Sometimes, a column was referenced instead of a named range (:issue:`545`)
* [Mac]: Python 2.7 was raising a ``LookupError: unknown encoding: mbcs`` (:issue:`544`)
* Fixed docs regarding set_mock_caller (:issue:`543`)

v0.9.1 (Aug 5, 2016)
--------------------

This is a bug fix release: As to be expected after a rewrite, there were some rough edges that have now been taken care of:

* [Win] Opening a file via ``xw.Book()`` was causing an additional ``Book1`` to be opened in case Excel was not running yet (:issue:`531`)
* [Win] Some users were getting an ImportError (:issue:`533`)
* [PY 2.7] ``RunPython`` was broken with Python 2.7 (:issue:`537`)
* Some corrections in the docs (:issue:`538` and :issue:`536`)


.. _v0.9_release_notes:

v0.9.0 (Aug 2, 2016)
--------------------

Exciting times! v0.9.0 is a complete rewrite of xlwings with loads of syntax changes (hence the version jump). But more
importantly, this release adds a ton of new features and bug fixes that would have otherwise been impossible. Some of the
highlights are listed below, but make sure to check out the full :ref:`migration guide <migrate_to_0.9>` for the syntax changes in details.
Note, however, that the syntax for user defined functions (UDFs) did not change.
At this point, the API is fairly stable and we're expecting only smaller changes on our way towards a stable v1.0 release.

* **Active** book instead of **current** book: ``xw.Range('A1')`` goes against the active sheet of the active book
  like you're used to from VBA. Instantiating an explicit connection to a Book is not necessary anymore:

    >>> import xlwings as xw
    >>> xw.Range('A1').value = 11
    >>> xw.Range('A1').value
    11.0

* Excel Instances: Full support of multiple Excel instances (even on Mac!)

    >>> app1 = xw.App()
    >>> app2 = xw.App()
    >>> xw.apps
    Apps([<Excel App 1668>, <Excel App 1644>])

* New powerful object model based on collections and close to Excel's original, allowing to fully qualify objects:
  ``xw.apps[0].books['MyBook.xlsx'].sheets[0].range('A1:B2').value``

  It supports both Python indexing (square brackets) and Excel indexing (round brackets):

  ``xw.books[0].sheets[0]`` is the same as ``xw.books(1).sheets(1)``

  It also supports indexing and slicing of range objects:

    >>> rng = xw.Range('A1:E10')
    >>> rng[1]
    <Range [Workbook1]Sheet1!$B$1>
    >>> rng[:2, :2]
    <Range [Workbook1]Sheet1!$A$1:$B$2>

  For more details, see :ref:`syntax_overview`.

* UDFs can now also be imported from packages, not just modules (:issue:`437`)

* Named Ranges: Introduction of full object model and proper support for sheet and workbook scope (:issue:`256`)

* Excel doesn't become the active window anymore so the focus stays on your Python environment (:issue:`414`)

* When writing to ranges while Excel is busy, xlwings is now retrying until Excel is idle again (:issue:`468`)

* :meth:`xlwings.view()` has been enhanced to accept an optional sheet object (:issue:`469`)

* Objects like books, sheets etc. can now be compared (e.g. ``wb1 == wb2``) and are properly hashable

* Note that support for Python 2.6 has been dropped

Some of the new methods/properties worth mentioning are:

* :any:`xlwings.App.display_alerts`
* :meth:`xlwings.App.macro` in addition to :meth:`xlwings.Book.macro`
* :meth:`xlwings.App.kill`
* :any:`xlwings.Sheet.cells`
* :any:`xlwings.Range.rows`
* :any:`xlwings.Range.columns`
* :meth:`xlwings.Range.end`
* :any:`xlwings.Range.raw_value`

Bug Fixes
*********

* See `here <https://github.com/xlwings/xlwings/issues?q=is%3Aclosed+is%3Aissue+milestone%3Av0.9.0+label%3Abug>`_
  for details about which bugs have been fixed.


.. _migrate_to_0.9:

Migrate to v0.9
---------------

The purpose of this document is to enable you a smooth experience when upgrading to xlwings v0.9.0 and above by laying out
the concept and syntax changes in detail. If you want to get an overview of the new features and bug fixes, have a look at the
:ref:`release notes <v0.9_release_notes>`. Note that the syntax for User Defined Functions (UDFs) didn't change.

Full qualification: Using collections
*************************************

The new object model allows to specify the Excel application instance if needed:

* **old**: ``xw.Range('Sheet1', 'A1', wkb=xw.Workbook('Book1'))``

* **new**: ``xw.apps[0].books['Book1'].sheets['Sheet1'].range('A1')``

See :ref:`syntax_overview` for the details of the new object model.

Connecting to Books
*******************

* **old**: ``xw.Workbook()``
* **new**: ``xw.Book()`` or via ``xw.books`` if you need to control the app instance.

See :ref:`connect_to_workbook` for the details.

Active Objects
**************

::

    # Active app (i.e. Excel instance)
    >>> app = xw.apps.active

    # Active book
    >>> wb = xw.books.active  # in active app
    >>> wb = app.books.active  # in specific app

    # Active sheet
    >>> sht = xw.sheets.active  # in active book
    >>> sht = wb.sheets.active  # in specific book

    # Range on active sheet
    >>> xw.Range('A1')  # on active sheet of active book of active app

Round vs. Square Brackets
*************************

Round brackets follow Excel's behavior (i.e. 1-based indexing), while square brackets use Python's 0-based indexing/slicing.

As an example, the following all reference the same range::

    xw.apps[0].books[0].sheets[0].range('A1')
    xw.apps(1).books(1).sheets(1).range('A1')
    xw.apps[0].books['Book1'].sheets['Sheet1'].range('A1')
    xw.apps(1).books('Book1').sheets('Sheet1').range('A1')

Access the underlying Library/Engine
************************************

* **old**: ``xw.Range('A1').xl_range`` and ``xl_sheet`` etc.

* **new**: ``xw.Range('A1').api``, same for all other objects

This returns a ``pywin32`` COM object on Windows and an ``appscript`` object on Mac.


Cheat sheet
***********

Note that ``sht`` stands for a sheet object, like e.g. (in 0.9.0 syntax): ``sht = xw.books['Book1'].sheets[0]``

+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            | v0.9.0                                           | v0.7.2                                                             |
+============================+==================================================+====================================================================+
| Active Excel instance      | ``xw.apps.active``                               | unsupported                                                        |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| New Excel instance         | ``app = xw.App()``                               | unsupported                                                        |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Get app from book          | ``app = wb.app``                                 | ``app = xw.Application(wb)``                                       |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Target installation (Mac)  | ``app = xw.App(spec=...)``                       | ``wb = xw.Workbook(app_target=...)``                               |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Hide Excel Instance        | ``app = xw.App(visible=False)``                  | ``wb = xw.Workbook(app_visible=False)``                            |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Selected Range             | ``app.selection``                                | ``wb.get_selection()``                                             |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Calculation mode           | ``app.calculation = 'manual'``                   | ``app.calculation = xw.constants.Calculation.xlCalculationManual`` |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| All books in app           | ``app.books``                                    | unsupported                                                        |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Fully qualified book       | ``app.books['Book1']``                           | unsupported                                                        |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Active book in active app  | ``xw.books.active``                              | ``xw.Workbook.active()``                                           |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| New book in active app     | ``wb = xw.Book()``                               | ``wb = xw.Workbook()``                                             |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| New book in specific app   | ``wb = app.books.add()``                         | unsupported                                                        |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| All sheets in book         | ``wb.sheets``                                    | ``xw.Sheet.all(wb)``                                               |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Call a macro in an addin   | ``app.macro('MacroName')``                       | unsupported                                                        |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| First sheet of book wb     | ``wb.sheets[0]``                                 | ``xw.Sheet(1, wkb=wb)``                                            |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Active sheet               | ``wb.sheets.active``                             | ``xw.Sheet.active(wkb=wb)`` or ``wb.active_sheet``                 |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Add sheet                  | ``wb.sheets.add()``                              | ``xw.Sheet.add(wkb=wb)``                                           |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Sheet count                | ``wb.sheets.count`` or ``len(wb.sheets)``        | ``xw.Sheet.count(wb)``                                             |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Add chart to sheet         | ``chart = wb.sheets[0].charts.add()``            | ``chart = xw.Chart.add(sheet=1, wkb=wb)``                          |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Existing chart             | ``wb.sheets['Sheet 1'].charts[0]``               | ``xw.Chart('Sheet 1', 1)``                                         |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Chart Type                 | ``chart.chart_type = '3d_area'``                 | ``chart.chart_type = xw.constants.ChartType.xl3DArea``             |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Add picture to sheet       | ``wb.sheets[0].pictures.add('path/to/pic')``     | ``xw.Picture.add('path/to/pic', sheet=1, wkb=wb)``                 |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Existing picture           | ``wb.sheets['Sheet 1'].pictures[0]``             | ``xw.Picture('Sheet 1', 1)``                                       |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Matplotlib                 | ``sht.pictures.add(fig, name='x', update=True)`` | ``xw.Plot(fig).show('MyPlot', sheet=sht, wkb=wb)``                 |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Table expansion            | ``sht.range('A1').expand('table')``              | ``xw.Range(sht, 'A1', wkb=wb).table``                              |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Vertical expansion         | ``sht.range('A1').expand('down')``               | ``xw.Range(sht, 'A1', wkb=wb).vertical``                           |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Horizontal expansion       | ``sht.range('A1').expand('right')``              | ``xw.Range(sht, 'A1', wkb=wb).horizontal``                         |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Set name of range          | ``sht.range('A1').name = 'name'``                | ``xw.Range(sht, 'A1', wkb=wb).name = 'name'``                      |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| Get name of range          | ``sht.range('A1').name.name``                    | ``xw.Range(sht, 'A1', wkb=wb).name``                               |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
|                            |                                                  |                                                                    |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+
| mock caller                | ``xw.Book('file.xlsm').set_mock_caller()``       | ``xw.Workbook.set_mock_caller('file.xlsm')``                       |
+----------------------------+--------------------------------------------------+--------------------------------------------------------------------+

v0.7.2 (May 18, 2016)
---------------------

Bug Fixes
*********
* [Win] UDFs returning Pandas DataFrames/Series containing ``nan`` were failing (:issue:`446`).
* [Win] ``RunFrozenPython`` was not finding the executable (:issue:`452`).
* The xlwings VBA module was not finding the Python interpreter if ``PYTHON_WIN`` or ``PYTHON_MAC`` contained spaces (:issue:`461`).


v0.7.1 (April 3, 2016)
----------------------

Enhancements
************
* [Win]: User Defined Functions (UDFs) support now optional/default arguments (:issue:`363`)
* [Win]: User Defined Functions (UDFs) support now multiple source files, see also under API changes below. For example
  (VBA settings): ``UDF_MODULES="common;myproject"``
* VBA Subs & Functions are now callable from Python:

    As an example, this VBA function:

    .. code-block:: basic

        Function MySum(x, y)
            MySum = x + y
        End Function

    can be accessed like this:

    >>> import xlwings as xw
    >>> wb = xw.Workbook.active()
    >>> my_sum = wb.macro('MySum')
    >>> my_sum(1, 2)
    3.0
* New ``xw.view`` method: This opens a new workbook and displays an object on its first sheet. E.g.:

    >>> import xlwings as xw
    >>> import pandas as pd
    >>> import numpy as np
    >>> df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
    >>> xw.view(df)

* New docs about :ref:`matplotlib` and :ref:`custom_converter`
* New method: :meth:`xlwings.Range.formula_array` (:issue:`411`)

API changes
***********

* VBA settings: ``PYTHON_WIN`` and ``PYTHON_MAC`` must now include the interpreter if you are not using the default
  (``PYTHON_WIN = ""``) (:issue:`289`). E.g.::

    PYTHON_WIN: "C:\Python35\pythonw.exe"
    PYTHON_MAC: "/usr/local/bin/python3.5"

* [Win]: VBA settings: ``UDF_PATH`` has been replaced with ``UDF_MODULES``. The default behaviour doesn't change though
  (i.e. if ``UDF_MODULES = ""``, then a Python source file with the same name as the Excel file, but with ``.py`` ending
  will be imported from the same directory as the Excel file).

  **New**:

  .. code-block:: basic

    UDF_MODULES: "mymodule"
    PYTHONPATH: "C:\path\to"

  **Old**:

  .. code-block:: basic

    UDF_PATH: "C:\path\to\mymodule.py"


Bug Fixes
*********
* Numpy scalars issues were resolved (:issue:`415`)
* [Win]: xlwings was failing with freezers like cx_Freeze (:issue:`413`)
* [Win]: UDFs were failing if they were returning ``None`` or ``np.nan`` (:issue:`390`)
* Multiindex Pandas Series have been fixed (:issue:`383`)
* [Mac]: ``xlwings runpython install`` was failing (:issue:`424`)

v0.7.0 (March 4, 2016)
----------------------

This version marks an important first step on our path towards a stable release. It introduces **converters**, a new and powerful
concept that brings a consistent experience for how Excel Ranges and their values are treated both when **reading** and **writing** but
also across **xlwings.Range** objects and **User Defined Functions** (UDFs).

As a result, a few highlights of this release include:

* Pandas DataFrames and Series are now supported for reading and writing, both via Range object and UDFs
* New Range converter options: ``transpose``, ``dates``, ``numbers``, ``empty``, ``expand``
* New dictionary converter
* New UDF debug server
* No more pyc files when using ``RunPython``

Converters are accessed via the new ``options`` method when dealing with ``xlwings.Range`` objects or via the ``@xw.arg``
and ``@xw.ret`` decorators when using UDFs. As an introductory sample, let's look at how to read and write Pandas DataFrames:

.. figure:: images/df_converter.png
  :scale: 55%

**Range object**::

    >>> import xlwings as xw
    >>> import pandas as pd
    >>> wb = xw.Workbook()
    >>> df = xw.Range('A1:D5').options(pd.DataFrame, header=2).value
    >>> df
        a     b
        c  d  e
    ix
    10  1  2  3
    20  4  5  6
    30  7  8  9

    # Writing back using the defaults:
    >>> Range('A1').value = df

    # Writing back and changing some of the options, e.g. getting rid of the index:
    >>> Range('B7').options(index=False).value = df

**UDFs**:

This is the same sample as above (starting in ``Range('A13')`` on screenshot). If you wanted to return a DataFrame with
the defaults, the ``@xw.ret`` decorator can be left away. ::

    @xw.func
    @xw.arg('x', pd.DataFrame, header=2)
    @xw.ret(index=False)
    def myfunction(x):
       # x is a DataFrame, do something with it
       return x


Enhancements
************

* Dictionary (``dict``) converter:

  .. figure:: images/dict_converter.png
    :scale: 80%

  ::

    >>> Range('A1:B2').options(dict).value
    {'a': 1.0, 'b': 2.0}
    >>> Range('A4:B5').options(dict, transpose=True).value
    {'a': 1.0, 'b': 2.0}

* ``transpose`` option: This works in both directions and finally allows us to e.g. write a list in column
  orientation to Excel (:issue:`11`)::

    Range('A1').options(transpose=True).value = [1, 2, 3]

* ``dates`` option: This allows us to read Excel date-formatted cells in specific formats:

    >>> import datetime as dt
    >>> Range('A1').value
    datetime.datetime(2015, 1, 13, 0, 0)
    >>> Range('A1').options(dates=dt.date).value
    datetime.date(2015, 1, 13)

* ``empty`` option: This allows us to override the default behavior for empty cells:

   >>> Range('A1:B1').value
   [None, None]
   >>> Range('A1:B1').options(empty='NA')
   ['NA', 'NA']

* ``numbers`` option: This transforms all numbers into the indicated type.

    >>> xw.Range('A1').value = 1
    >>> type(xw.Range('A1').value)  # Excel stores all numbers interally as floats
    float
    >>> type(xw.Range('A1').options(numbers=int).value)
    int

* ``expand`` option: This works the same as the Range properties ``table``, ``vertical`` and ``horizontal`` but is
  only evaluated when getting the values of a Range::

    >>> import xlwings as xw
    >>> wb = xw.Workbook()
    >>> xw.Range('A1').value = [[1,2], [3,4]]
    >>> rng1 = xw.Range('A1').table
    >>> rng2 = xw.Range('A1').options(expand='table')
    >>> rng1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> rng2.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> xw.Range('A3').value = [5, 6]
    >>> rng1.value
    [[1.0, 2.0], [3.0, 4.0]]
    >>> rng2.value
    [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]

All these options work the same with decorators for UDFs, e.g. for transpose::

  @xw.arg('x', transpose=True)
  @xw.ret(transpose=True)
  def myfunction(x):
      # x will be returned unchanged as transposed both when reading and writing
      return x


**Note**: These options (``dates``, ``empty``, ``numbers``) currently apply to the whole Range and can't be selectively
applied to e.g. only certain columns.

* UDF debug server

  The new UDF debug server allows you to easily debug UDFs: just set ``UDF_DEBUG_SERVER = True`` in the VBA Settings,
  at the top of the xlwings VBA module (make sure to update it to the latest version!). Then add the following lines
  to your Python source file and run it::


    if __name__ == '__main__':
        xw.serve()

  When you recalculate the Sheet, the code will stop at breakpoints or print any statements that you may have. For
  details, see: :ref:`debugging`.

* pyc files: The creation of pyc files has been disabled when using ``RunPython``, leaving your directory in an
  uncluttered state when having the Python source file next to the Excel workbook (:issue:`326`).


API changes
***********

* UDF decorator changes (it is assumed that xlwings is imported as ``xw`` and numpy as ``np``):

  ==============================  =========================
  **New**                         **Old**
  ==============================  =========================
  ``@xw.func``                    ``@xw.xlfunc``
  ``@xw.arg``                     ``@xw.xlarg``
  ``@xw.ret``                     ``@xw.xlret``
  ``@xw.sub``                     ``@xw.xlsub``
  ==============================  =========================

  Pay attention to the following subtle change:

  ==============================  =========================
  **New**                         **Old**
  ==============================  =========================
  ``@xw.arg("x", np.array)``      ``@xw.xlarg("x", "nparray")``
  ==============================  =========================

* Samples of how the new options method replaces the old Range keyword arguments:

  =============================================================   ===========================
  **New**                                                         **Old**
  =============================================================   ===========================
  ``Range('A1:A2').options(ndim=2)``                              ``Range('A1:A2', atleast_2d=True)``
  ``Range('A1:B2').options(np.array)``                            ``Range('A1:B2', asarray=True)``
  ``Range('A1').options(index=False, header=False).value = df``   ``Range('A1', index=False, header=False).value = df``
  =============================================================   ===========================

* Upon writing, Pandas Series are now shown by default with their name and index name, if they exist. This can be
  changed using the same options as for DataFrames (:issue:`276`)::

    import pandas as pd

    # unchanged behaviour
    Range('A1').value = pd.Series([1,2,3])

    # Changed behaviour: This will print a header row in Excel
    s = pd.Series([1,2,3], name='myseries', index=pd.Index([0,1,2], name='myindex'))
    Range('A1').value = s

    # Control this behaviour like so (as with DataFrames):
    Range('A1').options(header=False, index=True).value = s

* NumPy scalar values

  Previously, NumPy scalar values were returned as ``np.atleast_1d``. To keep the same behaviour, this now has to be
  set explicitly using ``ndim=1``. Otherwise they're returned as numpy scalar values.

  ===============================================                  =========================
  **New**                                                          **Old**
  ===============================================                  =========================
  ``Range('A1').options(np.array, ndim=1).value``                  ``Range('A1', asarray=True).value``
  ===============================================                  =========================

Bug Fixes
*********

A few bugfixes were made: :issue:`352`, :issue:`359`.


v0.6.4 (January 6, 2016)
------------------------

API changes
***********
None

Enhancements
************

* Quickstart: It's now easier than ever to start a new xlwings project, simply use the commmand line client (:issue:`306`):

  ``xlwings quickstart myproject`` will produce a folder with the following files, ready to be used (see :ref:`command_line`)::

    myproject
      |--myproject.xlsm
      |--myproject.py


* New documentation about how to use xlwings with other languages like R and Julia, see :ref:`r_and_julia`.

Bug Fixes
*********

* [Win]: Importing UDFs with the add-in was throwing an error if the filename was including characters like spaces or dashes (:issue:`331`).
  To fix this, close Excel completely and run ``xlwings addin update``.

* [Win]: ``Workbook.caller()`` is now also accessible within functions that are decorated with ``@xlfunc``. Previously,
  it was only available with functions that used the ``@xlsub`` decorator (:issue:`316`).

* Writing a Pandas DataFrame failed in case the index was named the same as a column (:issue:`334`).


v0.6.3 (December 18, 2015)
--------------------------

Bug Fixes
*********

* [Mac]: This fixes a bug introduced in v0.6.2: When using ``RunPython`` from VBA, errors were not shown in a pop-up window (:issue:`330`).


v0.6.2 (December 15, 2015)
--------------------------

API changes
***********

* LOG_FILE: So far, the log file has been placed next to the Excel file per default (VBA settings). This has been changed as it was
  causing issues for files on SharePoint/OneDrive and Mac Excel 2016: The place where ``LOG_FILE = ""`` refers to depends on the OS and the Excel version.

Enhancements
************
* [Mac]: This version adds support for the VBA module on Mac Excel 2016 (i.e. the ``RunPython`` command) and is now feature equivalent
  with Mac Excel 2011 (:issue:`206`).

Bug Fixes
*********
* [Win]: On certain systems, the xlwings dlls weren't found (:issue:`323`).


v0.6.1 (December 4, 2015)
-------------------------

Bug Fixes
*********

* [Python 3]: The command line client has been fixed (:issue:`319`).
* [Mac]: It now works correctly with ``psutil>=3.0.0`` (:issue:`315`).


v0.6.0 (November 30, 2015)
--------------------------

API changes
***********
None

Enhancements
************

* **User Defined Functions (UDFs) - currently Windows only**

  The `ExcelPython <https://github.com/ericremoreynolds/excelpython/>`_ project has been fully merged into xlwings. This means
  that on Windows, UDF's are now supported via decorator syntax. A simple example::

    from xlwings import xlfunc

    @xlfunc
    def double_sum(x, y):
        """Returns twice the sum of the two arguments"""
        return 2 * (x + y)

  For **array formulas** with or without **NumPy**, see the docs: :ref:`udfs`

* **Command Line Client**

  The new xlwings command line client makes it easy to work with the xlwings **template** and the developer **add-in**
  (the add-in is currently Windows-only). E.g. to create a new Excel spreadsheet from the template, run::

      xlwings template open

  For all commands, see the docs: :ref:`command_line`

* **Other enhancements**:

  - New method: :meth:`xlwings.Sheet.delete`
  - New method: :meth:`xlwings.Range.top`
  - New method: :meth:`xlwings.Range.left`


v0.5.0 (November 10, 2015)
--------------------------

API changes
***********
None

Enhancements
************
This version adds support for Matplotlib! Matplotlib figures can be shown in Excel as pictures in just 2 lines of code:

.. figure:: images/matplotlib.png
  :scale: 80%

1) Get a matplotlib ``figure`` object:

* via PyPlot interface::

    import matplotlib.pyplot as plt
    fig = plt.figure()
    plt.plot([1, 2, 3, 4, 5])

* via object oriented interface::

    from matplotlib.figure import Figure
    fig = Figure(figsize=(8, 6))
    ax = fig.add_subplot(111)
    ax.plot([1, 2, 3, 4, 5])

* via Pandas::

    import pandas as pd
    import numpy as np

    df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
    ax = df.plot(kind='bar')
    fig = ax.get_figure()

2) Show it in Excel as picture::

    plot = Plot(fig)
    plot.show('Plot1')

See the full API: :meth:`xlwings.Plot`. There's also a new example available both on
`GitHub <https://github.com/xlwings/xlwings/tree/master/examples/matplotlib/>`_ and as download on the
`homepage <http://www.xlwings.org/examples>`_.

**Other enhancements**:

* New :meth:`xlwings.Shape` class
* New :meth:`xlwings.Picture` class
* The ``PYTHONPATH`` in the VBA settings now accepts multiple directories, separated by ``;`` (:issue:`258`)
* An explicit exception is raised when ``Range`` is called with 0-based indices (:issue:`106`)

Bug Fixes
*********
* ``Sheet.add`` was not always acting on the correct workbook (:issue:`287`)
* Iteration over a ``Range`` only worked the first time (:issue:`272`)
* [Win]: Sometimes, an error was raised when Excel was not running (:issue:`269`)
* [Win]: Non-default Python interpreters (as specified in the VBA settings under ``PYTHON_WIN``) were not found
  if the path contained a space (:issue:`257`)


v0.4.1 (September 27, 2015)
---------------------------

API changes
***********
None

Enhancements
************

This release makes it easier than ever to connect to Excel from Python! In addition to the existing ways, you can now
connect to the active Workbook (on Windows across all instances) and if the Workbook is already open, it's good enough
to refer to it by name (instead of having to use the full path). Accordingly, this is how you make a connection to...
(:issue:`30` and :issue:`226`):

* a new workbook: ``wb = Workbook()``
* the active workbook [New!]: ``wb = Workbook.active()``
* an unsaved workbook: ``wb = Workbook('Book1')``
* a saved (open) workbook by name (incl. xlsx etc.) [New!]: ``wb = Workbook('MyWorkbook.xlsx')``
* a saved (open or closed) workbook by path: ``wb = Workbook(r'C:\\path\\to\\file.xlsx')``

Also, there are some new docs:

* :ref:`connect_to_workbook`
* :ref:`missing_features`

Bug Fixes
*********

* The Excel template was updated to the latest VBA code (:issue:`234`).
* Connections to files that are saved on OneDrive/SharePoint are now working correctly (:issue:`215`).
* Various issues with timezone-aware objects were fixed (:issue:`195`).
* [Mac]: A certain range of integers were not written to Excel (:issue:`227`).


v0.4.0 (September 13, 2015)
---------------------------

API changes
***********
None

Enhancements
************
The most important update with this release was made on Windows: The methodology used to make a connection
to Workbooks has been completely replaced. This finally allows xlwings to reliably connect to multiple instances of
Excel even if the Workbooks are opened from untrusted locations (network drives or files downloaded from the internet).
This gets rid of the dreaded ``Filename is already open...`` error message that was sometimes shown in this
context. It also allows the VBA hooks (``RunPython``) to work correctly if the very same file is opened in various instances of
Excel.

Note that you will need to update the VBA module and that apart from ``pywin32`` there is now a new dependency for the
Windows version: ``comtypes``. It should be installed automatically though when installing/upgrading xlwings with
``pip``.


Other updates:

* Added support to manipulate named Ranges (:issue:`92`):

    >>> wb = Workbook()
    >>> Range('A1').name = 'Name1'
    >>> Range('A1').name
    >>> 'Name1'
    >>> del wb.names['Name1']

* New ``Range`` properties (:issue:`81`):
    * :meth:`xlwings.Range.column_width`
    * :meth:`xlwings.Range.row_height`
    * :meth:`xlwings.Range.width`
    * :meth:`xlwings.Range.height`

* ``Range`` now also accepts ``Sheet`` objects, the following 3 ways are hence all valid (:issue:`92`)::

    r = Range(1, 'A1')
    r = Range('Sheet1', 'A1')
    sheet1 = Sheet(1)
    r = Range(sheet1, 'A1')

* [Win]: Error pop-ups show now the full error message that can also be copied with ``Ctrl-C`` (:issue:`221`).


Bug Fixes
*********
* The VBA module was not accepting lower case drive letters (:issue:`205`).
* Fixed an error when adding a new Sheet that was already existing (:issue:`211`).

v0.3.6 (July 14, 2015)
----------------------

API changes
***********

``Application`` as attribute of a ``Workbook`` has been removed (``wb`` is a ``Workbook`` object):

==============================  =========================
**Correct Syntax (as before)**  **Removed**
==============================  =========================
``Application(wkb=wb)``         ``wb.application``
==============================  =========================

Enhancements
************

**Excel 2016 for Mac Support** (:issue:`170`)

Excel 2016 for Mac is finally supported (Python side). The VBA hooks (``RunPython``) are currently not yet supported.
In more details:

* This release allows Excel 2011 and Excel 2016 to be installed in parallel.
* ``Workbook()`` will open the default Excel installation (usually Excel 2016).
* The new keyword argument ``app_target`` allows to connect to a different Excel installation, e.g.::

    Workbook(app_target='/Applications/Microsoft Office 2011/Microsoft Excel')

  Note that ``app_target`` is only available on Mac. On Windows, if you want to change the version of Excel that
  xlwings talks to, go to ``Control Panel > Programs and Features`` and ``Repair`` the Office version that you want
  as default.

* The ``RunPython`` calls in VBA are not yet available through Excel 2016 but Excel 2011 doesn't get confused anymore if
  Excel 2016 is installed on the same system - make sure to update your VBA module!

**Other enhancements**

* New method: :meth:`xlwings.Application.calculate` (:issue:`207`)

Bug Fixes
*********

* [Win]: When using the ``OPTIMIZED_CONNECTION`` on Windows, Excel left an orphaned process running after
  closing (:issue:`193`).

Various improvements regarding unicode file path handling, including:

* [Mac]: Excel 2011 for Mac now supports unicode characters in the filename when called via VBA's ``RunPython``
  (but not in the path - this is a limitation of Excel 2011 that will be resolved in Excel 2016) (:issue:`154`).
* [Win]: Excel on Windows now handles unicode file paths correctly with untrusted documents.
  (:issue:`154`).

v0.3.5 (April 26, 2015)
-----------------------

API changes
***********

``Sheet.autofit()`` and ``Range.autofit()``: The integer argument for the axis has been removed (:issue:`186`).
Use string arguments ``rows`` or ``r`` for autofitting rows and ``columns`` or ``c`` for autofitting columns
(as before).

Enhancements
************
New methods:

* :meth:`xlwings.Range.row` (:issue:`143`)
* :meth:`xlwings.Range.column` (:issue:`143`)
* :meth:`xlwings.Range.last_cell` (:issue:`142`)

Example::

    >>> rng = Range('A1').table
    >>> rng.row, rng.column
    (1, 1)
    >>> rng.last_cell.row, rng.last_cell.column
    (4, 5)

Bug Fixes
*********
* The ``unicode`` bug on Windows/Python3 has been fixed (:issue:`161`)

v0.3.4 (March 9, 2015)
----------------------

Bug Fixes
*********
* The installation error on Windows has been fixed (:issue:`160`)

v0.3.3 (March 8, 2015)
----------------------

API changes
***********

None

Enhancements
************

* New class ``Application`` with ``quit`` method and properties ``screen_updating`` und ``calculation`` (:issue:`101`,
  :issue:`158`, :issue:`159`). It can be
  conveniently accessed from within a Workbook (on Windows, ``Application`` is instance dependent). A few examples:

  >>> from xlwings import Workbook, Calculation
  >>> wb = Workbook()
  >>> wb.application.screen_updating = False
  >>> wb.application.calculation = Calculation.xlCalculationManual
  >>> wb.application.quit()

* New headless mode: The Excel application can be hidden either during ``Workbook`` instantiation or through the
  ``application`` object:

  >>> wb = Workbook(app_visible=False)
  >>> wb.application.visible
  False
  >>> wb.application.visible = True

* Newly included Excel template which includes the xlwings VBA module and boilerplate code. This is currently
  accessible from an interactive interpreter session only:

  >>> from xlwings import Workbook
  >>> Workbook.open_template()

Bug Fixes
*********

* [Win]: ``datetime.date`` objects were causing an error (:issue:`44`).

* Depending on how it was instantiated, Workbook was sometimes missing the ``fullname`` attribute (:issue:`76`).

* ``Range.hyperlink`` was failing if the hyperlink had been set as formula (:issue:`132`).

* A bug introduced in v0.3.0 caused frozen versions (eg. with ``cx_Freeze``) to fail (:issue:`133`).

* [Mac]: Sometimes, xlwings was causing an error when quitting the Python interpreter (:issue:`136`).

v0.3.2 (January 17, 2015)
-------------------------

API changes
***********

None

Enhancements
************

None

Bug Fixes
*********

* The :meth:`xlwings.Workbook.save` method has been fixed to show the expected behavior (:issue:`138`): Previously,
  calling `save()` without a `path` argument would always create a new file in the current working directory. This is
  now only happening if the file hasn't been previously saved.



v0.3.1 (January 16, 2015)
-------------------------

API changes
***********

None

Enhancements
************

* New method :meth:`xlwings.Workbook.save` (:issue:`110`).

* New method :meth:`xlwings.Workbook.set_mock_caller` (:issue:`129`). This makes calling files from both
  Excel and Python much easier::

    import os
    from xlwings import Workbook, Range

    def my_macro():
        wb = Workbook.caller()
        Range('A1').value = 1

    if __name__ == '__main__':
        # To run from Python, not needed when called from Excel.
        # Expects the Excel file next to this source file, adjust accordingly.
        path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'myfile.xlsm'))
        Workbook.set_mock_caller(path)
        my_macro()

* The ``simulation`` example on the homepage works now also on Mac.

Bug Fixes
*********

* [Win]: A long-standing bug that caused the Excel file to close and reopen under certain circumstances has been
  fixed (:issue:`10`): Depending on your security settings (Trust Center) and in connection with files downloaded from
  the internet or possibly in connection with some add-ins, Excel was either closing the file and reopening it or giving
  a "file already open" warning. This has now been fixed which means that the examples downloaded from the homepage should
  work right away after downloading and unzipping.


v0.3.0 (November 26, 2014)
--------------------------

API changes
***********

* To reference the calling Workbook when running code from VBA, you now have to use ``Workbook.caller()``. This means
  that ``wb = Workbook()`` is now consistently creating a new Workbook, whether the code is called interactively or
  from VBA.

  ==============================  =========================
  **New**                         **Old**
  ==============================  =========================
  ``Workbook.caller()``           ``Workbook()``
  ==============================  =========================

Enhancements
************
This version adds two exciting but still **experimental** features from
`ExcelPython` (**Windows only!**):

* Optimized connection: Set the ``OPTIMIZED_CONNECTION = True`` in the VBA settings. This will use a COM server that
  will keep the connection to Python alive between different calls and is therefore much more efficient. However,
  changes in the Python code are not being picked up until the ``pythonw.exe`` process is restarted by killing it
  manually in the Windows Task Manager. The suggested workflow is hence to set ``OPTIMIZED_CONNECTION = False`` for
  development and only set it to ``True`` for production - keep in mind though that this feature is still experimental!

* User Defined Functions (UDFs): Using ExcelPython's wrapper syntax in VBA, you can expose Python functions as UDFs, see
  :ref:`udfs` for details.

**Note:** ExcelPython's developer add-in that autogenerates the VBA wrapper code by simply using Python decorators
isn't available through xlwings yet.


Further enhancements include:

* New method :meth:`xlwings.Range.resize` (:issue:`90`).
* New method :meth:`xlwings.Range.offset` (:issue:`89`).
* New property :attr:`xlwings.Range.shape` (:issue:`109`).
* New property :attr:`xlwings.Range.size` (:issue:`109`).
* New property :attr:`xlwings.Range.hyperlink` and new method :meth:`xlwings.Range.add_hyperlink` (:issue:`104`).
* New property :attr:`xlwings.Range.color` (:issue:`97`).
* The ``len`` built-in function can now be used on ``Range`` (:issue:`109`):

    >>> len(Range('A1:B5'))
    5

* The ``Range`` object is now iterable (:issue:`108`)::

    for cell in Range('A1:B2'):
        if cell.value < 2:
            cell.color = (255, 0, 0)

* [Mac]: The VBA module finds now automatically the default Python installation as per ``PATH`` variable on
  ``.bash_profile`` when ``PYTHON_MAC = ""`` (the default in the VBA settings) (:issue:`95`).
* The VBA error pop-up can now be muted by setting ``SHOW_LOG = False`` in the VBA settings. To be used with
  care, but it can be useful on Mac, as the pop-up window is currently showing printed log messages even if no error
  occurred(:issue:`94`).

Bug Fixes
*********

* [Mac]: Environment variables from ``.bash_profile`` are now available when called from VBA, e.g. by using:
  ``os.environ['USERNAME']`` (:issue:`95`)


v0.2.3 (October 17, 2014)
-------------------------

API changes
***********

None

Enhancements
************

* New method ``Sheet.add()`` (:issue:`71`)::

    >>> Sheet.add()  # Place at end with default name
    >>> Sheet.add('NewSheet', before='Sheet1')  # Include name and position
    >>> new_sheet = Sheet.add(after=3)
    >>> new_sheet.index
    4

* New method ``Sheet.count()``::

    >>> Sheet.count()
    3

* ``autofit()`` works now also on ``Sheet`` objects, not only on ``Range`` objects (:issue:`66`)::

    >>> Sheet(1).autofit()  # autofit columns and rows
    >>> Sheet('Sheet1').autofit('c')  # autofit columns

* New property ``number_format`` for ``Range`` objects (:issue:`60`)::

    >>> Range('A1').number_format
    'General'
    >>> Range('A1:C3').number_format = '0.00%'
    >>> Range('A1:C3').number_format
    '0.00%'

  Works also with the ``Range`` properties ``table``, ``vertical``, ``horizontal``::

    >>> Range('A1').value = [1,2,3,4,5]
    >>> Range('A1').table.number_format = '0.00%'

* New method ``get_address`` for ``Range`` objects (:issue:`7`)::

    >>> Range((1,1)).get_address()
    '$A$1'
    >>> Range((1,1)).get_address(False, False)
    'A1'
    >>> Range('Sheet1', (1,1), (3,3)).get_address(True, False, include_sheetname=True)
    'Sheet1!A$1:C$3'
    >>> Range('Sheet1', (1,1), (3,3)).get_address(True, False, external=True)
    '[Workbook1]Sheet1!A$1:C$3'

* New method ``Sheet.all()`` returning a list with all Sheet objects::

    >>> Sheet.all()
    [<Sheet 'Sheet1' of Workbook 'Book1'>, <Sheet 'Sheet2' of Workbook 'Book1'>]
    >>> [i.name.lower() for i in Sheet.all()]
    ['sheet1', 'sheet2']
    >>> [i.autofit() for i in Sheet.all()]

Bug Fixes
*********

* xlwings works now also with NumPy < 1.7.0. Before, doing something like ``Range('A1').value = 'Foo'`` was causing
  a ``NotImplementedError: Not implemented for this type`` error when NumPy < 1.7.0 was installed (:issue:`73`).

* [Win]: The VBA module caused an error on the 64bit version of Excel (:issue:`72`).

* [Mac]: The error pop-up wasn't shown on Python 3 (:issue:`85`).

* [Mac]: Autofitting bigger Ranges, e.g. ``Range('A:D').autofit()`` was causing a time out (:issue:`74`).

* [Mac]: Sometimes, calling xlwings from Python was causing Excel to show old errors as pop-up alert (:issue:`70`).


v0.2.2 (September 23, 2014)
---------------------------

API changes
***********

* The ``Workbook`` qualification changed: It now has to be specified as keyword argument. Assume we have instantiated
  two Workbooks like so: ``wb1 = Workbook()`` and ``wb2 = Workbook()``. ``Sheet``, ``Range`` and ``Chart`` classes will
  default to ``wb2`` as it was instantiated last. To target ``wb1``, use the new ``wkb`` keyword argument:

  ==============================  =========================
  **New**                         **Old**
  ==============================  =========================
  ``Range('A1', wkb=wb1).value``  ``wb1.range('A1').value``
  ``Chart('Chart1', wkb=wb1)``    ``wb1.chart('Chart1')``
  ==============================  =========================

  Alternatively, simply set the current Workbook before using the ``Sheet``, ``Range`` or ``Chart`` classes::

    wb1.set_current()
    Range('A1').value

* Through the introduction of the ``Sheet`` class (see Enhancements), a few methods moved from the ``Workbook``
  to the ``Sheet`` class. Assume the current Workbook is: ``wb = Workbook()``:

  ====================================  ====================================
  **New**                               **Old**
  ====================================  ====================================
  ``Sheet('Sheet1').activate()``        ``wb.activate('Sheet1')``
  ``Sheet('Sheet1').clear()``           ``wb.clear('Sheet1')``
  ``Sheet('Sheet1').clear_contents()``  ``wb.clear_contents('Sheet1')``
  ``Sheet.active().clear_contents()``   ``wb.clear_contents()``
  ====================================  ====================================

* The syntax to add a new Chart has been slightly changed (it is a class method now):

  ===============================  ====================================
  **New**                          **Old**
  ===============================  ====================================
  ``Chart.add()``                  ``Chart().add()``
  ===============================  ====================================

Enhancements
************

* [Mac]: Python errors are now also shown in a Message Box. This makes the Mac version feature equivalent with the
  Windows version (:issue:`57`):

  .. figure:: images/mac_error.png
    :scale: 75%

* New ``Sheet`` class: The new class handles everything directly related to a Sheet. See the Python API section about
  ``Sheet`` for details (:issue:`62`). A few examples::

    >>> Sheet(1).name
    'Sheet1'
    >>> Sheet('Sheet1').clear_contents()
    >>> Sheet.active()
    <Sheet 'Sheet1' of Workbook 'Book1'>

* The ``Range`` class has a new method ``autofit()`` that autofits the width/height of either columns, rows or both
  (:issue:`33`).

  *Arguments*::

    axis : string or integer, default None
        - To autofit rows, use one of the following: 'rows' or 'r'
        - To autofit columns, use one of the following: 'columns' or 'c'
        - To autofit rows and columns, provide no arguments

  *Examples*::

    # Autofit column A
    Range('A:A').autofit()
    # Autofit row 1
    Range('1:1').autofit()
    # Autofit columns and rows, taking into account Range('A1:E4')
    Range('A1:E4').autofit()
    # AutoFit rows, taking into account Range('A1:E4')
    Range('A1:E4').autofit('rows')

* The ``Workbook`` class has the following additional methods: ``current()`` and ``set_current()``. They determine the
  default Workbook for ``Sheet``, ``Range`` or ``Chart``. On Windows, in case there are various Excel instances, when
  creating new or opening existing Workbooks,
  they are being created in the same instance as the current Workbook.

    >>> wb1 = Workbook()
    >>> wb2 = Workbook()
    >>> Workbook.current()
    <Workbook 'Book2'>
    >>> wb1.set_current()
    >>> Workbook.current()
    <Workbook 'Book1'>

* If a ``Sheet``, ``Range`` or ``Chart`` object is instantiated without an existing ``Workbook`` object, a user-friendly
  error message is raised (:issue:`58`).

* New docs about :ref:`debugging` and :ref:`datastructures`.


Bug Fixes
*********

* The ``atleast_2d`` keyword had no effect on Ranges consisting of a single cell and was raising an error when used in
  combination with the ``asarray`` keyword. Both have been fixed (:issue:`53`)::

    >>> Range('A1').value = 1
    >>> Range('A1', atleast_2d=True).value
    [[1.0]]
    >>> Range('A1', atleast_2d=True, asarray=True).value
    array([[1.]])

* [Mac]: After creating two new unsaved Workbooks with ``Workbook()``, any ``Sheet``, ``Range`` or ``Chart``
  object would always just access the latest one, even if the Workbook had been specified (:issue:`63`).

* [Mac]: When xlwings was imported without ever instantiating a ``Workbook`` object, Excel would start upon
  quitting the Python interpreter (:issue:`51`).

* [Mac]: When installing xlwings, it now requires ``psutil`` to be at least version ``2.0.0`` (:issue:`48`).


v0.2.1 (August 7, 2014)
-----------------------

API changes
***********

None

Enhancements
************

* All VBA user settings have been reorganized into a section at the top of the VBA xlwings module::

    PYTHON_WIN = ""
    PYTHON_MAC = GetMacDir("Home") & "/anaconda/bin"
    PYTHON_FROZEN = ThisWorkbook.Path & "\build\exe.win32-2.7"
    PYTHONPATH = ThisWorkbook.Path
    LOG_FILE = ThisWorkbook.Path & "\xlwings_log.txt"

* Calling Python from within Excel VBA is now also supported on Mac, i.e. Python functions can be called like
  this: ``RunPython("import bar; bar.foo()")``. Running frozen executables (``RunFrozenPython``) isn't available
  yet on Mac though.

Note that there is a slight difference in the way that this functionality behaves on Windows and Mac:

* **Windows**: After calling the Macro (e.g. by pressing a button), Excel waits until Python is done. In case there's an
  error in the Python code, a pop-up message is being shown with the traceback.

* **Mac**: After calling the Macro, the call returns instantly but Excel's Status Bar turns into "Running..." during the
  duration of the Python call. Python errors are currently not shown as a pop-up, but need to be checked in the
  log file. I.e. if the Status Bar returns to its default ("Ready") but nothing has happened, check out the log file
  for the Python traceback.

Bug Fixes
*********

None

Special thanks go to Georgi Petrov for helping with this release.

v0.2.0 (July 29, 2014)
----------------------

API changes
***********

None

Enhancements
************

* Cross-platform: xlwings is now additionally supporting Microsoft Excel for Mac. The only functionality that is not
  yet available is the possibility to call the Python code from within Excel via VBA macros.
* The ``clear`` and ``clear_contents`` methods of the ``Workbook`` object now default to the active
  sheet (:issue:`5`)::

    wb = Workbook()
    wb.clear_contents()  # Clears contents of the entire active sheet

Bug Fixes
*********

* DataFrames with MultiHeaders were sometimes getting truncated (:issue:`41`).


v0.1.1 (June 27, 2014)
----------------------

API Changes
***********

* If ``asarray=True``, NumPy arrays are now always at least 1d arrays, even in the case of a single cell (:issue:`14`)::

    >>> Range('A1', asarray=True).value
    array([34.])

* Similar to NumPy's logic, 1d Ranges in Excel, i.e. rows or columns, are now being read in as flat lists or 1d arrays.
  If you want the same behavior as before, you can use the ``atleast_2d`` keyword (:issue:`13`).

  .. note:: The ``table`` property is also delivering a 1d array/list, if the table Range is really a column or row.

  .. figure:: images/1d_ranges.png

  ::

    >>> Range('A1').vertical.value
    [1.0, 2.0, 3.0, 4.0]
    >>> Range('A1', atleast_2d=True).vertical.value
    [[1.0], [2.0], [3.0], [4.0]]
    >>> Range('C1').horizontal.value
    [1.0, 2.0, 3.0, 4.0]
    >>> Range('C1', atleast_2d=True).horizontal.value
    [[1.0, 2.0, 3.0, 4.0]]
    >>> Range('A1', asarray=True).table.value
    array([ 1.,  2.,  3.,  4.])
    >>> Range('A1', asarray=True, atleast_2d=True).table.value
    array([[ 1.],
           [ 2.],
           [ 3.],
           [ 4.]])

* The single file approach has been dropped. xlwings is now a traditional Python package.

Enhancements
************

* xlwings is now officially suppported on Python 2.6-2.7 and 3.1-3.4
* Support for Pandas ``Series`` has been added (:issue:`24`)::

    >>> import numpy as np
    >>> import pandas as pd
    >>> from xlwings import Workbook, Range
    >>> wb = Workbook()
    >>> s = pd.Series([1.1, 3.3, 5., np.nan, 6., 8.])
    >>> s
    0    1.1
    1    3.3
    2    5.0
    3    NaN
    4    6.0
    5    8.0
    dtype: float64
    >>> Range('A1').value = s
    >>> Range('D1', index=False).value = s

  .. figure:: images/pandas_series.png

* Excel constants have been added under their original Excel name, but categorized under their enum (:issue:`18`),
  e.g.::

    # Extra long version
    import xlwings as xl
    xl.constants.ChartType.xlArea

    # Long version
    from xlwings import constants
    constants.ChartType.xlArea

    # Short version
    from xlwings import ChartType
    ChartType.xlArea

* Slightly enhanced Chart support to control the ``ChartType`` (:issue:`1`)::

    >>> from xlwings import Workbook, Range, Chart, ChartType
    >>> wb = Workbook()
    >>> Range('A1').value = [['one', 'two'],[10, 20]]
    >>> my_chart = Chart().add(chart_type=ChartType.xlLine,
                               name='My Chart',
                               source_data=Range('A1').table)

  alternatively, the properties can also be set like this::

    >>> my_chart = Chart().add()  # Existing Charts: my_chart = Chart('My Chart')
    >>> my_chart.name = 'My Chart'
    >>> my_chart.chart_type = ChartType.xlLine
    >>> my_chart.set_source_data(Range('A1').table)

  .. figure:: images/chart_type.png
    :scale: 70%

* ``pytz`` is no longer a dependency as ``datetime`` object are now being read in from Excel as time-zone naive (Excel
  doesn't know timezones). Before, ``datetime`` objects got the UTC timezone attached.

* The ``Workbook`` class has the following additional methods: ``close()``
* The ``Range`` class has the following additional methods: ``is_cell()``, ``is_column()``, ``is_row()``,
  ``is_table()``


Bug Fixes
*********

* Writing ``None`` or ``np.nan`` to Excel works now (:issue:`16` & :issue:`15`).
* The import error on Python 3 has been fixed (:issue:`26`).
* Python 3 now handles Pandas DataFrames with MultiIndex headers correctly (:issue:`39`).
* Sometimes, a Pandas DataFrame was not handling ``nan`` correctly in Excel or numbers were being truncated
  (:issue:`31`) & (:issue:`35`).
* Installation is now putting all files in the correct place (:issue:`20`).


v0.1.0 (March 19, 2014)
-----------------------

Initial release of xlwings.
