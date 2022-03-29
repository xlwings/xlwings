.. _onedrive_sharepoint:

OneDrive and SharePoint
=======================

Since v0.27.4, xlwings works with locally synced files on OneDrive, OneDrive for Business, and SharePoint. Some constellations will work out-of-the-box, while others require you to edit the configuration via the ``xlwings.conf`` file (see :ref:`User Config<user_config>`) or the workbook's ``xlwings.conf`` sheet (see :ref:`Workbook Config<addin_wb_settings>`).

.. note:: This documentation is for OneDrive and SharePoint files that are synced to a local folder. This means that both, the Excel and Python file, need to show the green check mark in the File Explorer/Finder as status---a cloud icon will not work. If, in turn, you are looking for the documentation to run xlwings with Excel on the web, see :ref:`remote_interpreter`.

An easy workaround if you run into issues is to:

* Disable the ``ADD_WORKBOOK_TO_PYTHONPATH`` setting (either via the checkbox on the Ribbon or via the settings in the ``xlwings.conf`` sheet).
* Add the directory of your Python source file to the ``PYTHONPATH``---again, either via Ribbon or ``xlwings.conf`` sheet.

If you are using the PRO version, you could instead also embed your code to get around these issues.

For a bit more flexibility, follow the solutions below.

OneDrive (Personal)
-------------------

Default setups work out-of-the-box on Windows and macOS. If you get an error message, add the following setting with the correct path to the local root directory of your OneDrive. If possible, make use of environment variables (as shown in the examples) so the configuration will work across different users with the same setup:

* **Windows** (Example):

  +-------------------------+--------------------------+
  +``ONEDRIVE_CONSUMER_WIN``|``%USERPROFILE%\OneDrive``+
  +-------------------------+--------------------------+

* **macOS** (Example):

  +-------------------------+--------------------------+
  +``ONEDRIVE_CONSUMER_MAC``|``$HOME/OneDrive``        +
  +-------------------------+--------------------------+

OneDrive for Business
---------------------

* **Windows**: Default setups work out-of-the-box. If you get an error message, add the following setting with the correct path to the local root directory of your OneDrive for Business. If possible, make use of environment variables (as shown in the examples) so the configuration will work across different users with the same setup:

  +---------------------------+-------------------------------------------+
  +``ONEDRIVE_COMMERCIAL_WIN``|``%USERPROFILE%\OneDrive - My Company LLC``+
  +---------------------------+-------------------------------------------+

* **macOS**: macOS *always* requires the following setting with the correct path to the local root directory of your OneDrive for Business. If possible, make use of environment variables (as shown in the examples) so the configuration will work across different users with the same setup:

  +---------------------------+-------------------------------------------+
  +``ONEDRIVE_COMMERCIAL_MAC``|``$HOME/OneDrive - My Company LLC``        +
  +---------------------------+-------------------------------------------+

SharePoint (Online and On-Premises)
-----------------------------------

On Windows, the location of the local root folder of SharePoint can *sometimes* be derived from the OneDrive environment variables. Most of the time though, you'll have to provide the following setting (on macOS this is a must):

* **Windows**:

  +-------------------------+--------------------------------+
  +``SHAREPOINT_WIN``       |``%USERPROFILE%\My Company LLC``+
  +-------------------------+--------------------------------+

* **macOS**:

  +-------------------------+------------------------+
  +``SHAREPOINT_MAC``       |``$HOME/My Company LLC``+
  +-------------------------+------------------------+

Implementation Details & Limitations
------------------------------------

A lot of the xlwings functionality depends on the workbook's ``FullName`` property (via VBA/COM) that returns the local path of the file unless it is saved on OneDrive, OneDrive for Business or SharePoint **with AutoSave enabled**. In this case, it returns a URL instead.

URLs for OneDrive and OneDrive for Business can be translated fairly straight forward to the local equivalent. You will need to know the root directory of the local drive though: on Windows, these are usually provided via environment variables for OneDrive. On macOS they don't exist, which is the reason why you need to provide the root directory for OneDrive. On Windows, the root directory for SharePoint can sometimes be derived from the env vars, too, but this is not guaranteed. On macOS, you'll need to provide it always anyway.

SharePoint, unfortunately, allows you to map the drives locally in any way you want and there's no way to reliably get the local path for these files. On Windows, xlwings first checks the registry for the mapping. If this doesn't work, xlwings checks if the local path is mapped by using the defaults and if the file can't be found, it checks all existing local files on SharePoint. If it finds one with the same name, it'll use this. If, however, it finds more than one with the same name, you will get an error message. In this case, you can either rename the file to something unique across all the locally synced SharePoint files or you can change the ``SHAREPOINT_WIN/MAC`` setting to not stop at the root folder but include additional folders. As an example, assume you have the following file structure on your local SharePoint:

.. code-block:: text

    My Company LLC/
    └── sitename1/
        └── myfile.xlsx
    └── sitename2 - Documents/
        └── myfile.xlsx

In this case, you could either rename one of the files, or you could add a path that goes beyond the root folder (preferably under the ``xlwings.conf`` sheet):

+-------------------------+------------------------------------------------------+
+``SHAREPOINT_WIN``       |``%USERPROFILE%/My Company LLC/sitename2 - Documents``+
+-------------------------+------------------------------------------------------+