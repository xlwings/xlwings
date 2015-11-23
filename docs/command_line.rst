.. _command_line:

Command Line Client
===================

.. versionadded:: 0.6.0

xlwings comes with a command line client that allows an easy handling of template and add-in.


Template
--------

After installing, the templates are accessible via Excel's Menu:

* Win Excel 2007, 2010: ``File > New > My templates``
* Win Excel 2013, 2016: There's an additional step needed as explained `here <https://support.office.com/en-us/article/Where-are-my-custom-templates-88ed77ca-df34-49e9-9087-3f01ae296e6e/>`_
* Mac: ``File > New from template``

**Commands**

* ``xlwings template open``: Opens a new Workbook with the xlwings VBA module

* ``xlwings template install``: Copies the xlwings template file to the correct Excel folder

* ``xlwings template update``: Replaces the current xlwings template with the latest one

* ``xlwings template remove``: Removes the template from Excel's template folder

* ``xlwings template status``: Shows if the template is installed together with the installation path.


Add-in (Currently Windows-only)
-------------------------------

After installing the add-in, it will be available as xlwings tab on the Ribbon.
This is currently not available on Mac.

.. note:: Excel needs to be closed before installing/updating the add-in!

**Commands**

* ``xlwings addin install``: Copies the xlwings add-in to the XLSTART folder

* ``xlwings addin update``: Replaces the current add-in with the latest one

* ``xlwings addin remove``: Removes the add-in from the XLSTART folder

* ``xlwings addin status``: Shows if the add-in is installed together with the installation path.


