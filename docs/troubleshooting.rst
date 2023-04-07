.. _troubleshooting:

Troubleshooting
===============

Issue: dll not found
--------------------

Solution:

1) ``xlwings32-<version>.dll`` and ``xlwings64-<version>.dll`` are both in the same directory as your ``python.exe``. If not, something went wrong
   with your installation. Reinstall it with ``pip`` or ``conda``, see :ref:`installation:installation`.
2) Check your ``Interpreter`` in the add-in or config sheet. If it is empty, then you need to be able to open a windows command prompt and type
   ``python`` to start an interactive Python session. If you get the error ``'python' is not recognized as an internal or external command,
   operable program or batch file.``, then you have two options: Either add the path of where your ``python.exe`` lives to your Windows path
   (see https://www.computerhope.com/issues/ch000549.htm) or set the full path to your interpreter in the add-in or your config sheet, e.g.
   ``C:\Users\MyUser\anaconda\pythonw.exe``

Issue: Files that are saved on OneDrive or SharePoint cause an error to pop up
------------------------------------------------------------------------------

Solution:

See the dedicated page about how to configure OneDrive and Sharepoint: :ref:`onedrive_sharepoint`.

Issue: Python was not found; run without arguments to install from the Microsoft Store, or disable this shortcut from Settings > Manage App Execution Aliases.
--------------------------------------------------------------------------------------------------------------------------------------------------------------

Cause:

The Python interpreter is not correctly installed or the configuration does not point to the Python interpreter.

Solution:

1) Verifiy that an interpreter is installed. This can be an installation provided by conda, virtual environment or the xlwings-installer(requires pro)
2) Check the configuration of xlwings accoring to its hierachy :ref:`_config_hierarchy:`