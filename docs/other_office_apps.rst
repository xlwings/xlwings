.. _other_office_apps:

xlwings with other Office Apps
==============================

xlwings can also be used to call Python functions from VBA within Office apps other than Excel (like Outlook, Access etc.).

.. note::
    New in v0.12.0 and still in a somewhat early stage that involves a bit of manual work.
    Currently, this functionality is only available on Windows for UDFs. The ``RunPython`` functionality
    is currently not supported.


How To
------

1) As usual, write your Python function and import it into Excel (see :ref:`udfs`).
2) Press ``Alt-F11`` to get into the VBA editor, then right-click on the ``xlwings_udfs`` VBA module and select ``Export File...``.
   Save the ``xlwings_udfs.bas`` file somewhere.
3) Switch into the other Office app, e.g. Microsoft Access and click again ``Alt-F11`` to get into the VBA editor. Right-click on the
   VBA Project and ``Import File...``, then select the file that you exported in the previous step. Once imported, replace the app
   name in the first line to the one that you are using, i.e. ``Microsoft Access`` or ``Microsoft Outlook`` etc. so that the first 
   line then reads: ``#Const App = "Microsoft Access"``
4) Now import the standalone xlwings VBA module (``xlwings.bas``). You can find it in your xlwings installation folder. To know where that is, do::

    >>> import xlwings as xw
    >>> xlwings.__path__

   And finally do the same as in the previous step and replace the App name in the first line with the name of the
   corresponding app that you are using. You are now able to call the Python function from VBA.

Config
------

The other Office apps will use the same global config file as you are editing via the Excel ribbon add-in. When it makes sense,
you'll be able to use the directory config file (e.g. you can put it next to your Access or Word file) or you can hardcode
the path to the config file in the VBA standalone module, e.g. in the function ``GetDirectoryConfigFilePath``
(e.g. suggested when using Outlook that doesn't really have the same concept of files like the other Office apps).
NOTE: For Office apps without file concept, you need to make sure that the ``PYTHONPATH`` points to the directory with the
Python source file.
For details on the different config options, see :ref:`Config <config_file>`.