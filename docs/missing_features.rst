.. _missing_features:

Missing Features
================

If you're missing a feature in xlwings, do the following:

1) Most importantly, open an issue on `GitHub <https://github.com/ZoomerAnalytics/xlwings/issues>`_.
   If it's something bigger or if you want support from other users, consider opening a
   `feature request <https://zoomeranalytics.uservoice.com/>`_. Only feature requests that are logged
   will be implemented eventually.

2) Workaround: in essence, xlwings is just a smart wrapper around `pywin32 <http://sourceforge.net/projects/pywin32/>`_ on
   Windows and `appscript <http://appscript.sourceforge.net/>`_ on Mac. You can access the underlying objects by doing:

   .. code-block:: python

        >>> wb = Workbook('Book1')
        >>> wb.xl_workbook
        <COMObject <unknown>>  # Windows/pywin32
        app('/Applications/Microsoft Excel.app').workbooks['Book1']  # Mac/appscript

   This works accordingly for ``Range.xl_range``, ``Sheet.xl_sheet``, ``Chart.xl_chart`` etc.

   The underlying objects will offer you pretty much everything you can do with VBA. But apart from looking ugly,
   keep in mind that **it makes your code platform specific (!)**, i.e. even if you go for option 2), you should still
   do option 1) and open an issue so the feature finds it's way into the library (cross-platform and with a Pythonic
   syntax).

Example: Workaround to use VBA's ``Range.WrapText``
---------------------------------------------------
::

    # Windows:
    Range('A1').xl_range.WrapText = True

    # Mac:
    Range('A1').xl_range.wrap_text.set(True)
