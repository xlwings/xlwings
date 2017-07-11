.. _migrate_to_0.11:

Migrate to v0.11 (Add-in)
=========================

This migration guide shows you how you can start using the new xlwings add-in as opposed to the old xlwings VBA module
(and the old add-in that consisted of just a single import button).

Upgrade the xlwings Python package
----------------------------------

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
------------------

1. If you have the old xlwings addin installed, find the location and remove it or overwrite it with the new version (see next step).
   If you installed it via the xlwings command line client, you should be able to do: ``xlwings addin remove``.
2. Close Excel. Run ``xlwings addin install`` from a command prompt. Reopen Excel and check if the xlwings Ribbon
   appears. If not, copy ``xlwings.xlam`` (from your xlwings installation folder under ``addin\xlwings.xlam`` manually
   into the ``XLSTART`` folder.
   You can find the location of this folder under Options > Trust Center > Trust Center Settings... > Trusted Locations,
   under the description ``Excel default location: User StartUp``. Restart Excel and you should see the add-in.


Upgrade existing workbooks
--------------------------

1. Make a backup of your Excel file
2. Open the file and go to the VBA Editor (``Alt-F11``)
3. Remove the xlwings VBA module
4. Add a reference to the xlwings addin, see :ref:`addin_installation`
5. If you want to use workbook specific settings, add a sheet ``xlwings.conf``, see :ref:`addin_wb_settings`


**Note**: To import UDFs, you need to have the reference to the xlwings add-in set!
