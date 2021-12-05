.. _missing_features:

Missing Features
================

If you're missing a feature in xlwings, do the following:

1) Most importantly, open an issue on `GitHub <https://github.com/xlwings/xlwings/issues>`_.
   Adding functionality should be user driven, so only if you tell us about what you're missing,
   it's eventually going to find its way into the library. By the way, we also appreciate pull requests!

2) Workaround: in essence, xlwings is just a smart wrapper around `pywin32 <https://github.com/mhammond/pywin32/>`_ on
   Windows and `appscript <http://appscript.sourceforge.net/>`_ on Mac. You can access the underlying objects by calling
   the ``api`` property:

   .. code-block:: python

        >>> sheet = xw.Book().sheets[0]
        >>> sheet.api
        <COMObject <unknown>>  # Windows/pywin32
        app(pid=2319).workbooks['Workbook1'].worksheets[1]  # Mac/appscript

   This works accordingly for the other objects like ``sheet.range('A1').api`` etc.

   The underlying objects will offer you pretty much everything you can do with VBA, using the syntax of pywin32 (which
   pretty much feels like VBA) and appscript (which doesn't feel like VBA).
   But apart from looking ugly, keep in mind that **it makes your code platform specific (!)**, i.e. even if you go for
   option 2), you should still follow option 1) and open an issue so the feature finds it's way into the library
   (cross-platform and with a Pythonic syntax).

Example: Workaround to use VBA's ``Range.WrapText``
---------------------------------------------------
::

    # Windows
    sheet.range('A1').api.WrapText = True

    # Mac
    sheet.range('A1').api.wrap_text.set(True)
