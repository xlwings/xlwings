.. _troubleshooting:

Troubleshooting
===============

Issue: dll not found
--------------------

Solution:

1) ``xlwings32-<version>.dll`` and ``xlwings64-<version>.dll`` are both in the same directory as your ``python.exe``. If not, something went wrong
   with your installation. Reinstall it with ``pip`` or ``conda``, see :ref:`installation`.
2) Check your ``Interpreter`` in the add-in or config sheet. If it is empty, then you need to be able to open a windows command prompt and type
   ``python`` to start an interactive Python session. If you get the error ``'python' is not recognized as an internal or external command,
   operable program or batch file.``, then you have two options: Either add the path of where your ``python.exe`` lives to your Windows path
   (see https://www.computerhope.com/issues/ch000549.htm) or set the full path to your interpreter in the add-in or your config sheet, e.g.
   ``C:\Users\MyUser\anaconda\pythonw.exe``

Issue: Couldn't find the local location of your OneDrive or SharePoint
----------------------------------------------------------------------

Solution:

On either the ``xlwings.conf`` sheet or on the ``xlwings.conf`` file under your home folder (for location see :ref:`user_config`), add the following setting::

    "ONEDRIVE_WIN", "C:\path\to\OneDrive"

Note: Don't use quotes on the ``xlwings.conf`` sheet and if you are on macOS, use ``ONEDRIVE_MAC`` instead. You need to use the ONEDRIVE setting, even if you use SharePoint.