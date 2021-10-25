.. _onedrive_sharepoint:

OneDrive and SharePoint
=======================

Since v0.25.0, xlwings works with files that are stored on OneDrive, OneDrive for Business, and SharePoint---as long as they are synced locally. Some constellations will work out-of-the-box, while others require you to edit the configuration via the ``xlwings.conf`` file (see :ref:`User Config<user_config>`) or the workbook's ``xlwings.conf`` sheet (see :ref:`Workbook Config<addin_wb_settings>`).



.. warning:: **xlwings only works with OneDrive and SharePoint files that are synced to a local folder!**

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
