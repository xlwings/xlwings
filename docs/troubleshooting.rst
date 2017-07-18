.. _troubleshooting:

Troubleshooting
===============

xlwingsXX.dll not found
-----------------------

Check:

1) ``xlwings32.dll`` and ``xlwings64.dll`` are both in the same directory as your ``python.exe``. If not, something went wrong
   with your installation. Reinstall it with ``pip`` or ``conda``, see :ref:`installation`.
2) Check your ``Interpreter`` in the add-in or config sheet. If it is empty, then you need to be able to open a windows command prompt and type
   ``python`` to start an interactive Python session. If you get the error ``'python' is not recognized as an internal or external command,
   operable program or batch file.``, then you have two options: Either add the path of where your python.exe lives to your Windows path
   (see https://www.computerhope.com/issues/ch000549.htm) or set the full path to your interpreter in the add-in or your config sheet, e.g.
   ``C:\Users\MyUser\anaconda\pythonw.exe``