"""
xlwings - Make Excel fly with Python!

Homepage and documentation: https://www.xlwings.org

Copyright (C) 2014-present, Zoomer Analytics GmbH.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import numbers
import os
import re
import sys
import warnings
from contextlib import contextmanager
from pathlib import Path

import xlwings

from . import LicenseError, ShapeAlreadyExists, XlwingsError, utils

# Optional imports
try:
    import matplotlib as mpl
    from matplotlib.backends.backend_agg import FigureCanvas
except ImportError:
    mpl = None
    FigureCanvas = None

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import PIL
except ImportError:
    PIL = None


class Collection:
    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl.api

    def __call__(self, name_or_index):
        return self._wrap(impl=self.impl(name_or_index))

    def __len__(self):
        return len(self.impl)

    @property
    def count(self):
        """
        Returns the number of objects in the collection.
        """
        return len(self)

    def __iter__(self):
        for impl in self.impl:
            yield self._wrap(impl=impl)

    def __getitem__(self, key):
        if isinstance(key, numbers.Number):
            length = len(self)
            if key >= length:
                raise IndexError("Index %s out of range (%s elements)" % (key, length))
            if key < 0:
                if key < -length:
                    raise IndexError(
                        "Index %s out of range (%s elements)" % (key, length)
                    )
                key += length
            return self(key + 1)
        elif isinstance(key, slice):
            raise ValueError(
                self.impl.__class__.__name__ + " object does not support slicing"
            )
        else:
            return self(key)

    def __contains__(self, key):
        return key in self.impl

    # used by repr - by default the name of the collection class, but can be overridden
    @property
    def _name(self):
        return self.__class__.__name__

    def __repr__(self):
        r = []
        for i, x in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(x))

        return "{}({})".format(self._name, "[" + ", ".join(r) + "]")


class Engines:
    def __init__(self):
        self.active = None
        self.engines = []
        self.engines_by_name = {}

    def add(self, engine):
        self.engines.append(engine)
        self.engines_by_name[engine.name] = engine

    @property
    def count(self):
        return len(self)

    def __call__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return self.engines[name_or_index - 1]
        else:
            return self.engines_by_name[name_or_index]

    def __getitem__(self, name_or_index):
        if isinstance(name_or_index, numbers.Number):
            return self.engines[name_or_index]
        else:
            try:
                return self.engines_by_name[name_or_index]
            except KeyError:
                engine_to_product_name = {
                    "remote": "xlwings Server",
                    "calamine": "xlwings Reader",
                }
                if not xlwings.__pro__ and name_or_index != "excel":
                    raise LicenseError(
                        f"{engine_to_product_name.get(name_or_index, name_or_index)} "
                        "requires a license key. Install one by running 'xlwings "
                        "license update -k your-key-here' or by setting the "
                        "XLWINGS_LICENSE_KEY environment variable."
                    )
                else:
                    raise

    def __len__(self):
        return len(self.engines)

    def __iter__(self):
        for engine in self.engines:
            yield engine

    def __repr__(self):
        return "{}({})".format(self.__class__.__name__, repr(list(self)))


class Engine:
    def __init__(self, impl):
        self.impl = impl

    @property
    def apps(self):
        return Apps(impl=self.impl.apps)

    @property
    def name(self):
        return self.impl.name

    @property
    def type(self):
        return self.impl.type

    def activate(self):
        engines.active = self

    def __repr__(self):
        return f"<Engine {self.name}>"


class Apps:
    """
    A collection of all :meth:`app <App>` objects:

    >>> import xlwings as xw
    >>> xw.apps
    Apps([<Excel App 1668>, <Excel App 1644>])
    """

    def __init__(self, impl):
        self.impl = impl

    def keys(self):
        """
        Provides the PIDs of the Excel instances
        that act as keys in the Apps collection.

        .. versionadded:: 0.13.0
        """
        return self.impl.keys()

    def add(self, **kwargs):
        """
        Creates a new App. The new App becomes the active one. Returns an App object.
        """
        return App(impl=self.impl.add(**kwargs))

    @property
    def active(self):
        """
        Returns the active app.

        .. versionadded:: 0.9.0
        """
        for app in self.impl:
            return App(impl=app)
        return None

    def __call__(self, i):
        return self[i]

    def __repr__(self):
        return "{}({})".format(
            getattr(self.__class__, "_name", self.__class__.__name__), repr(list(self))
        )

    def __getitem__(self, item):
        return App(impl=self.impl[item])

    def __len__(self):
        return len(self.impl)

    @property
    def count(self):
        """
        Returns the number of apps.

        .. versionadded:: 0.9.0
        """
        return len(self)

    def cleanup(self):
        """
        Removes Excel zombie processes (Windows-only). Note that this is automatically
        called with ``App.quit()`` and ``App.kill()`` and when the Python interpreter
        exits.

        .. versionadded:: 0.30.2
        """
        self.impl.cleanup()

    def __iter__(self):
        for app in self.impl:
            yield App(impl=app)


engines = Engines()


class App:
    """
    An app corresponds to an Excel instance and should normally be used as context
    manager to make sure that everything is properly cleaned up again and to prevent
    zombie processes. New Excel instances can be fired up like so::

        import xlwings as xw

        with xw.App() as app:
            print(app.books)

    An app object is a member of the :meth:`apps <xlwings.main.Apps>` collection:

    >>> xw.apps
    Apps([<Excel App 1668>, <Excel App 1644>])
    >>> xw.apps[1668]  # get the available PIDs via xw.apps.keys()
    <Excel App 1668>
    >>> xw.apps.active
    <Excel App 1668>

    Parameters
    ----------
    visible : bool, default None
        Returns or sets a boolean value that determines whether the app is visible. The
        default leaves the state unchanged or sets visible=True if the object doesn't
        exist yet.

    spec : str, default None
        Mac-only, use the full path to the Excel application,
        e.g. ``/Applications/Microsoft Office 2011/Microsoft Excel`` or
        ``/Applications/Microsoft Excel``

        On Windows, if you want to change the version of Excel that xlwings talks to, go
        to ``Control Panel > Programs and Features`` and ``Repair`` the Office version
        that you want as default.


    .. note::
        On Mac, while xlwings allows you to run multiple instances of Excel, it's a
        feature that is not officially supported by Excel for Mac: Unlike on Windows,
        Excel will not ask you to open a read-only version of a file if it is already
        open in another instance. This means that you need to watch out yourself so
        that the same file is not being overwritten from different instances.
    """

    def __init__(self, visible=None, spec=None, add_book=True, impl=None):
        if impl is None:
            self.impl = engines.active.apps.add(
                spec=spec, add_book=add_book, visible=visible
            ).impl
            if visible or visible is None:
                self.visible = True
        else:
            self.impl = impl
            if visible:
                self.visible = True
        self._pid = self.pid

    @property
    def engine(self):
        return Engine(impl=self.impl.engine)

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    @property
    def version(self):
        """
        Returns the Excel version number object.

        Examples
        --------
        >>> import xlwings as xw
        >>> xw.App().version
        VersionNumber('15.24')
        >>> xw.apps[10559].version.major
        15

        .. versionchanged:: 0.9.0
        """
        return utils.VersionNumber(self.impl.version)

    @property
    def selection(self):
        """
        Returns the selected cells as Range.

        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.selection) if self.impl.selection else None

    def activate(self, steal_focus=False):
        """
        Activates the Excel app.

        Parameters
        ----------
        steal_focus : bool, default False
            If True, make frontmost application
            and hand over focus from Python to Excel.


        .. versionadded:: 0.9.0
        """
        # Win Excel >= 2013 fails if visible=False...
        # we may somehow not be using the correct HWND
        self.impl.activate(steal_focus)
        if self.engine.name != "remote":
            if self != apps.active:
                raise Exception(
                    "Could not activate App! "
                    "Try to instantiate the App with visible=True."
                )

    @property
    def visible(self):
        """
        Gets or sets the visibility of Excel to ``True`` or  ``False``.

        .. versionadded:: 0.3.3
        """
        return self.impl.visible

    @visible.setter
    def visible(self, value):
        self.impl.visible = value

    def quit(self):
        """
        Quits the application without saving any workbooks.

        .. versionadded:: 0.3.3

        """
        return self.impl.quit()

    def kill(self):
        """
        Forces the Excel app to quit by killing its process.

        .. versionadded:: 0.9.0
        """
        return self.impl.kill()

    @property
    def screen_updating(self):
        """
        Turn screen updating off to speed up your script. You won't be able to see what
        the script is doing, but it will run faster. Remember to set the screen_updating
        property back to True when your script ends.

        .. versionadded:: 0.3.3
        """
        return self.impl.screen_updating

    @screen_updating.setter
    def screen_updating(self, value):
        self.impl.screen_updating = value

    @property
    def display_alerts(self):
        """
        The default value is True. Set this property to False to suppress prompts and
        alert messages while code is running; when a message requires a response, Excel
        chooses the default response.

        .. versionadded:: 0.9.0
        """
        return self.impl.display_alerts

    @display_alerts.setter
    def display_alerts(self, value):
        self.impl.display_alerts = value

    @property
    def enable_events(self):
        """
        ``True`` if events are enabled. Read/write boolean.

        .. versionadded:: 0.24.4
        """
        return self.impl.enable_events

    @enable_events.setter
    def enable_events(self, value):
        self.impl.enable_events = value

    @property
    def interactive(self):
        """
        ``True`` if Excel is in interactive mode. If you set this property to ``False``,
        Excel blocks all input from the keyboard and mouse (except input to dialog boxes
        that are displayed by your code). Read/write Boolean.
        NOTE: Not supported on macOS.

        .. versionadded:: 0.24.4
        """
        return self.impl.interactive

    @interactive.setter
    def interactive(self, value):
        self.impl.interactive = value

    @property
    def startup_path(self):
        """
        Returns the path to ``XLSTART`` which is where the xlwings add-in gets
        copied to by doing ``xlwings addin install``.

        .. versionadded:: 0.19.4
        """
        return self.impl.startup_path

    @property
    def calculation(self):
        """
        Returns or sets a calculation value that represents the calculation mode.
        Modes: ``'manual'``, ``'automatic'``, ``'semiautomatic'``

        Examples
        --------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> wb.app.calculation = 'manual'

        .. versionchanged:: 0.9.0
        """
        return self.impl.calculation

    @calculation.setter
    def calculation(self, value):
        self.impl.calculation = value

    def calculate(self):
        """
        Calculates all open books.

        .. versionadded:: 0.3.6

        """
        self.impl.calculate()

    @property
    def books(self):
        """
        A collection of all Book objects that are currently open.

        .. versionadded:: 0.9.0
        """
        return Books(impl=self.impl.books)

    @property
    def hwnd(self):
        """
        Returns the Window handle (Windows-only).

        .. versionadded:: 0.9.0
        """
        return self.impl.hwnd

    @property
    def path(self):
        """
        Returns the path to where the App is installed.

        .. versionadded:: 0.28.4
        """
        return self.impl.path

    @property
    def pid(self):
        """
        Returns the PID of the app.

        .. versionadded:: 0.9.0
        """
        return self.impl.pid

    def range(self, cell1, cell2=None):
        """
        Range object from the active sheet of the active book, see :meth:`Range`.

        .. versionadded:: 0.9.0
        """
        return self.books.active.sheets.active.range(cell1, cell2)

    def macro(self, name):
        """
        Runs a Sub or Function in Excel VBA that are not part of a specific workbook
        but e.g. are part of an add-in.

        Arguments
        ---------
        name : Name of Sub or Function with or without module name,
               e.g., ``'Module1.MyMacro'`` or ``'MyMacro'``

        Examples
        --------
        This VBA function:

        .. code-block:: vb.net

            Function MySum(x, y)
                MySum = x + y
            End Function

        can be accessed like this:

        >>> import xlwings as xw
        >>> app = xw.App()
        >>> my_sum = app.macro('MySum')
        >>> my_sum(1, 2)
        3

        Types are supported too:

        .. code-block:: vb.net

            Function MySum(x as integer, y as integer)
                MySum = x + y
            End Function

        >>> import xlwings as xw
        >>> app = xw.App()
        >>> my_sum = app.macro('MySum')
        >>> my_sum(1, 2)
        3

        However typed arrays are not supported. So the following won't work

        .. code-block:: vb.net

            Function MySum(arr() as integer)
                ' code here
            End Function

        See also: :meth:`Book.macro`

        .. versionadded:: 0.9.0
        """
        return Macro(self, name)

    @property
    def status_bar(self):
        """
        Gets or sets the value of the status bar.
        Returns ``False`` if Excel has control of it.

        .. versionadded:: 0.20.0
        """
        return self.impl.status_bar

    @status_bar.setter
    def status_bar(self, value):
        self.impl.status_bar = value

    @property
    def cut_copy_mode(self):
        """
        Gets or sets the status of the cut or copy mode.
        Accepts ``False`` for setting and returns ``None``,
        ``copy`` or ``cut`` when getting the status.

        .. versionadded:: 0.24.0
        """
        return self.impl.cut_copy_mode

    @cut_copy_mode.setter
    def cut_copy_mode(self, value):
        self.impl.cut_copy_mode = value

    @contextmanager
    def properties(self, **kwargs):
        """
        Context manager that allows you to easily change the app's properties
        temporarily. Once the code leaves the with block, the properties are changed
        back to their previous state.
        Note: Must be used as context manager or else will have no effect. Also, you can
        only use app properties that you can both read and write.

        Examples
        --------
        ::

            import xlwings as xw
            app = App()

            # Sets app.display_alerts = False
            with app.properties(display_alerts=False):
                # do stuff

            # Sets app.calculation = 'manual' and app.enable_events = True
            with app.properties(calculation='manual', enable_events=True):
                # do stuff

            # Makes sure the status bar is reset even if an error happens in the with block
            with app.properties(status_bar='Calculating...'):
                # do stuff

        .. versionadded:: 0.24.4
        """
        initial_state = {}
        for attribute, value in kwargs.items():
            initial_state[attribute] = getattr(self, attribute, value)
            setattr(self, attribute, value)
        try:
            yield self
        finally:
            for attribute, value in initial_state.items():
                setattr(self, attribute, value)

    def create_report(self, template=None, output=None, book_settings=None, **data):
        warnings.warn("Deprecated. Use render_template instead.")
        return self.render_template(
            template=template, output=output, book_settings=book_settings, **data
        )

    def render_template(self, template=None, output=None, book_settings=None, **data):
        """
        This function requires xlwings :bdg-secondary:`PRO`.

        This is a convenience wrapper around :meth:`mysheet.render_template
        <xlwings.Sheet.render_template>`

        Writes the values of all key word arguments to the ``output`` file according to
        the ``template`` and the variables contained in there (Jinja variable syntax).
        Following variable types are supported:

        strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames, pictures
        and Matplotlib/Plotly figures.

        Parameters
        ----------
        template: str or path-like object
            Path to your Excel template, e.g. ``r'C:\\Path\\to\\my_template.xlsx'``

        output: str or path-like object
            Path to your Report, e.g. ``r'C:\\Path\\to\\my_report.xlsx'``

        book_settings: dict, default None
            A dictionary of ``xlwings.Book`` parameters, for details see:
            :attr:`xlwings.Book`.
            For example: ``book_settings={'update_links': False}``.

        data: kwargs
            All key/value pairs that are used in the template.

        Returns
        -------
        wb: xlwings Book


        .. versionadded:: 0.24.4
        """
        from .pro.reports import render_template

        return render_template(
            template=template,
            output=output,
            book_settings=book_settings,
            app=self,
            **data,
        )

    def alert(self, prompt, title=None, buttons="ok", mode=None, callback=None):
        """
        This corresponds to ``MsgBox`` in VBA, shows an alert/message box and returns
        the value of the pressed button. For xlwings Server, instead of
        returning a value, the function accepts the name of a callback to which it will
        supply the value of the pressed button.

        Parameters
        ----------

        prompt : str, default None
            The message to be displayed.

        title : str, default None
            The title of the alert.

        buttons : str, default ``"ok"``
            Can be either ``"ok"``, ``"ok_cancel"``, ``"yes_no"``, or
            ``"yes_no_cancel"``.

        mode : str, default None
            Can be ``"info"`` or ``"critical"``. Not supported by Google Sheets.

        callback : str, default None
            Only used by xlwings Server: you can provide the name of a
            function that will be called with the value of the pressed button as
            argument. The function has to exist on the client side, i.e., in VBA or
            JavaScript.

        Returns
        -------
        button_value: str or None
            Returns ``None`` when used with xlwings Server, otherwise the value
            of the pressed button in lowercase: ``"ok"``, ``"cancel"``, ``"yes"``,
            ``"no"``.


        .. versionadded:: 0.27.13
        """
        return self.impl.alert(prompt, title, buttons, mode, callback)

    def __repr__(self):
        return f"<App [{self.engine.name}] {self.pid}>"

    def __eq__(self, other):
        return type(other) is App and other.pid == self.pid

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.pid)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        self.quit()
        if sys.platform.startswith("win"):
            try:
                self.kill()
            except:  # noqa: E722
                pass


class Book:
    """
    A book object is a member of the :meth:`books <xlwings.main.Books>` collection:

    >>> import xlwings as xw
    >>> xw.books[0]
    <Book [Book1]>


    The easiest way to connect to a book is offered by ``xw.Book``: it looks for the
    book in all app instances and returns an error, should the same book be open in
    multiple instances. To connect to a book in the active app instance, use
    ``xw.books`` and to refer to a specific app, use:

    >>> app = xw.App()  # or xw.apps[10559] (get the PIDs via xw.apps.keys())
    >>> app.books['Book1']

    +--------------------+--------------------------------------+--------------------------------------------+
    |                    | xw.Book                              | xw.books                                   |
    +====================+======================================+============================================+
    | New book           | ``xw.Book()``                        | ``xw.books.add()``                         |
    +--------------------+--------------------------------------+--------------------------------------------+
    | Unsaved book       | ``xw.Book('Book1')``                 | ``xw.books['Book1']``                      |
    +--------------------+--------------------------------------+--------------------------------------------+
    | Book by (full)name | ``xw.Book(r'C:/path/to/file.xlsx')`` | ``xw.books.open(r'C:/path/to/file.xlsx')`` |
    +--------------------+--------------------------------------+--------------------------------------------+

    Parameters
    ----------
    fullname : str or path-like object, default None
        Full path or name (incl. xlsx, xlsm etc.) of existing workbook or name of an
        unsaved workbook. Without a full path, it looks for the file in the current
        working directory.
    update_links : bool, default None
        If this argument is omitted, the user is prompted to specify how links will be
        updated
    read_only : bool, default False
        True to open workbook in read-only mode
    format : str
        If opening a text file, this specifies the delimiter character
    password : str
        Password to open a protected workbook
    write_res_password : str
        Password to write to a write-reserved workbook
    ignore_read_only_recommended : bool, default False
        Set to ``True`` to mute the read-only recommended message
    origin : int
        For text files only. Specifies where it originated. Use Platform constants.
    delimiter : str
        If format argument is 6, this specifies the delimiter.
    editable : bool, default False
        This option is only for legacy Microsoft Excel 4.0 addins.
    notify : bool, default False
        Notify the user when a file becomes available If the file cannot be opened in
        read/write mode.
    converter : int
        The index of the first file converter to try when opening the file.
    add_to_mru : bool, default False
        Add this workbook to the list of recently added workbooks.
    local : bool, default False
        If ``True``, saves files against the language of Excel, otherwise against the
        language of VBA. Not supported on macOS.
    corrupt_load : int, default xlNormalLoad
        Can be one of xlNormalLoad, xlRepairFile or xlExtractData.
        Not supported on macOS.
    json : dict
        A JSON object as delivered by the MS Office Scripts or Google Apps Script
        xlwings module but in a deserialized form, i.e., as dictionary.

        .. versionadded:: 0.26.0

    mode : str, default None
        Either ``"i"`` (interactive (default)) or ``"r"`` (read). In interactive mode,
        xlwings opens the workbook in Excel, i.e., Excel needs to be installed. In read
        mode, xlwings reads from the file directly, without requiring Excel to be
        installed. Read mode requires xlwings :bdg-secondary:`PRO`.

        .. versionadded:: 0.28.0
    """

    def __init__(
        self,
        fullname=None,
        update_links=None,
        read_only=None,
        format=None,
        password=None,
        write_res_password=None,
        ignore_read_only_recommended=None,
        origin=None,
        delimiter=None,
        editable=None,
        notify=None,
        converter=None,
        add_to_mru=None,
        local=None,
        corrupt_load=None,
        impl=None,
        json=None,
        mode=None,
        engine=None,
    ):
        if not impl:
            if json:
                engine = engine if engine else "remote"
                impl = engines[engine].apps.active.books.open(json=json).impl
            elif fullname and mode == "r":
                engine = engine if engine else "calamine"
                impl = engines[engine].apps.active.books.open(fullname=fullname).impl
            elif fullname:
                fullname = utils.fspath(fullname)

                candidates = []
                for app in apps:
                    for wb in app.books:
                        # Comparing by name first saves us from having to compare the
                        # fullname for non-candidates, which can get around issues in
                        # case the fullname is a problematic URL (GH 1946)
                        if wb.name.lower() == os.path.split(fullname)[1].lower() and (
                            wb.fullname.lower() == fullname.lower()
                            or wb.name.lower() == fullname.lower()
                        ):
                            candidates.append((app, wb))

                app = apps.active
                if len(candidates) == 0:
                    if not app:
                        app = App(add_book=False)
                    impl = app.books.open(
                        fullname,
                        update_links,
                        read_only,
                        format,
                        password,
                        write_res_password,
                        ignore_read_only_recommended,
                        origin,
                        delimiter,
                        editable,
                        notify,
                        converter,
                        add_to_mru,
                        local,
                        corrupt_load,
                    ).impl
                elif len(candidates) > 1:
                    raise Exception(
                        "Workbook '%s' is open in more than one Excel instance."
                        % fullname
                    )
                else:
                    impl = candidates[0][1].impl
            else:
                # Open Excel if necessary and create a new workbook
                if apps.active:
                    impl = apps.active.books.add().impl
                else:
                    app = App()
                    impl = app.books[0].impl

        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine being
        used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def json(self):
        """
        Returns a JSON serializable object as expected by the MS Office Scripts or
        Google Apps Script xlwings module. Only available with book objects that have
        been instantiated via ``xw.Book(json=...)``.

        .. versionadded:: 0.26.0
        """
        return self.impl.json()

    def __eq__(self, other):
        return (
            isinstance(other, Book)
            and self.app == other.app
            and self.name == other.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash((self.app, self.name))

    @classmethod
    def caller(cls):
        """
        References the calling book when the Python function is called from Excel via
        ``RunPython``. Pack it into the function being called from Excel, e.g.::

            import xlwings as xw

             def my_macro():
                wb = xw.Book.caller()
                wb.sheets[0].range('A1').value = 1

        To be able to easily invoke such code from Python for debugging, use
        ``xw.Book.set_mock_caller()``.

        .. versionadded:: 0.3.0
        """
        wb, from_xl, hwnd = None, None, None
        for arg in sys.argv:
            if arg.startswith("--wb="):
                wb = arg.split("=")[1].strip()
            elif arg.startswith("--from_xl"):
                from_xl = arg.split("=")[1].strip()
            elif arg.startswith("--hwnd"):
                hwnd = arg.split("=")[1].strip()
        if hasattr(Book, "_mock_caller"):
            # Use mocking Book, see Book.set_mock_caller()
            return cls(impl=Book._mock_caller.impl)
        elif from_xl == "1":
            name = wb.lower()
            if sys.platform.startswith("win"):
                app = App(impl=xlwings._xlwindows.App(xl=int(hwnd)))
                return cls(impl=app.books[name].impl)
            else:
                # On Mac, the same file open in two instances is not supported
                if apps.active.version < 15:
                    name = name.encode("utf-8", "surrogateescape").decode("mac_latin2")
                return cls(impl=Book(name).impl)
        elif xlwings._xlwindows.BOOK_CALLER:
            # Called via OPTIMIZED_CONNECTION = True
            return cls(impl=xlwings._xlwindows.Book(xlwings._xlwindows.BOOK_CALLER))
        else:
            raise Exception(
                "Book.caller() must not be called directly. Call through Excel "
                "or set a mock caller first with Book.set_mock_caller()."
            )

    def set_mock_caller(self):
        """
        Sets the Excel file which is used to mock ``xw.Book.caller()`` when the code is
        called from Python and not from Excel via ``RunPython``.

        Examples
        --------
        ::

            # This code runs unchanged from Excel via RunPython and from Python directly
            import os
            import xlwings as xw

            def my_macro():
                sht = xw.Book.caller().sheets[0]
                sht.range('A1').value = 'Hello xlwings!'

            if __name__ == '__main__':
                xw.Book('file.xlsm').set_mock_caller()
                my_macro()

        .. versionadded:: 0.3.1
        """
        Book._mock_caller = self

    def macro(self, name):
        """
        Runs a Sub or Function in Excel VBA.

        Arguments
        ---------
        name : Name of Sub or Function with or without module name, e.g.,
        ``'Module1.MyMacro'`` or ``'MyMacro'``

        Examples
        --------
        This VBA function:

        .. code-block:: vb.net

            Function MySum(x, y)
                MySum = x + y
            End Function

        can be accessed like this:

        >>> import xlwings as xw
        >>> wb = xw.books.active
        >>> my_sum = wb.macro('MySum')
        >>> my_sum(1, 2)
        3

        See also: :meth:`App.macro`

        .. versionadded:: 0.7.1
        """
        return self.app.macro("'{0}'!{1}".format(self.name, name))

    @property
    def name(self):
        """
        Returns the name of the book as str.
        """
        return self.impl.name

    @property
    def sheets(self):
        """
        Returns a sheets collection that represents all the sheets in the book.

        .. versionadded:: 0.9.0
        """
        return Sheets(impl=self.impl.sheets)

    @property
    def app(self):
        """
        Returns an app object that represents the creator of the book.

        .. versionadded:: 0.9.0
        """
        return App(impl=self.impl.app)

    def close(self):
        """
        Closes the book without saving it.

        .. versionadded:: 0.1.1
        """
        self.impl.close()

    def save(self, path=None, password=None):
        """
        Saves the Workbook. If a path is provided, this works like SaveAs() in
        Excel. If no path is specified and if the file hasn't been saved previously,
        it's saved in the current working directory with the current filename.
        Existing files are overwritten without prompting. To change the file type,
        provide the appropriate extension, e.g. to save ``myfile.xlsx`` in the ``xlsb``
        format, provide ``myfile.xlsb`` as path.

        Arguments
        ---------
        path : str or path-like object, default None
            Path where you want to save the Book.
        password : str, default None
            Protection password with max. 15 characters

            .. versionadded :: 0.25.1

        Example
        -------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> wb.save()
        >>> wb.save(r'C:\\path\\to\\new_file_name.xlsx')


        .. versionadded:: 0.3.1
        """
        if path:
            path = utils.fspath(path)
        with self.app.properties(display_alerts=False):
            self.impl.save(path, password=password)

    @property
    def fullname(self):
        """
        Returns the name of the object, including its path on disk, as a string.
        Read-only String.

        """
        return self.impl.fullname

    @property
    def names(self):
        """
        Returns a names collection that represents all the names in the specified book
        (including all sheet-specific names).

        .. versionchanged:: 0.9.0

        """
        return Names(impl=self.impl.names)

    def activate(self, steal_focus=False):
        """
        Activates the book.

        Parameters
        ----------
        steal_focus : bool, default False
            If True, make frontmost window and hand over focus from Python to Excel.

        """
        self.app.activate(steal_focus=steal_focus)
        self.impl.activate()

    @property
    def selection(self):
        """
        Returns the selected cells as Range.

        .. versionadded:: 0.9.0
        """
        return Range(impl=self.app.selection.impl) if self.app.selection else None

    def to_pdf(
        self,
        path=None,
        include=None,
        exclude=None,
        layout=None,
        exclude_start_string="#",
        show=False,
        quality="standard",
    ):
        """
        Exports the whole Excel workbook or a subset of the sheets to a PDF file.
        If you want to print hidden sheets, you will need to list them explicitely
        under ``include``.

        Parameters
        ----------
        path : str or path-like object, default None
            Path to the PDF file, defaults to the same name as the workbook, in the same
            directory. For unsaved workbooks, it defaults to the current working
            directory instead.

        include : int or str or list, default None
            Which sheets to include: provide a selection of sheets in the form of sheet
            indices (1-based like in Excel) or sheet names. Can be an int/str for a
            single sheet or a list of int/str for multiple sheets.

        exclude : int or str or list, default None
            Which sheets to exclude: provide a selection of sheets in the form of sheet
            indices (1-based like in Excel) or sheet names. Can be an int/str for a
            single sheet or a list of int/str for multiple sheets.

        layout : str or path-like object, default None
            This argument requires xlwings :bdg-secondary:`PRO`.

            Path to a PDF file on which the report will be printed. This is ideal for
            headers and footers as well as borderless printing of graphics/artwork. The
            PDF file either needs to have only 1 page (every report page uses the same
            layout) or otherwise needs the same amount of pages as the report (each
            report page is printed on the respective page in the layout PDF).

            .. versionadded:: 0.24.3

        exclude_start_string : str, default ``'#'``
            Sheet names that start with this character/string will not be printed.

            .. versionadded:: 0.24.4

        show : bool, default False
            Once created, open the PDF file with the default application.

            .. versionadded:: 0.24.6

        quality : str, default ``'standard'``
            Quality of the PDF file. Can either be ``'standard'`` or ``'minimum'``.

            .. versionadded:: 0.26.2

        Examples
        --------
        >>> wb = xw.Book()
        >>> wb.sheets[0]['A1'].value = 'PDF'
        >>> wb.to_pdf()

        See also :meth:`xlwings.Sheet.to_pdf`

        .. versionadded:: 0.21.1
        """
        return utils.to_pdf(
            self,
            path=path,
            include=include,
            exclude=exclude,
            layout=layout,
            exclude_start_string=exclude_start_string,
            show=show,
            quality=quality,
        )

    def __repr__(self):
        return "<Book [{0}]>".format(self.name)

    def render_template(self, **data):
        """
        This method requires xlwings :bdg-secondary:`PRO`.

        Replaces all Jinja variables (e.g ``{{ myvar }}``) in the book
        with the keyword argument of the same name.

        .. versionadded:: 0.25.0

        Parameters
        ----------
        data: kwargs
            All key/value pairs that are used in the template.

        Examples
        --------

        >>> import xlwings as xw
        >>> book = xw.Book()
        >>> book.sheets[0]['A1:A2'].value = '{{ myvar }}'
        >>> book.render_template(myvar='test')
        """
        for sheet in reversed(self.sheets):
            sheet.render_template(**data)

    @property
    def sheet_names(self):
        """
        Returns
        -------

        sheet_names : List
            List of sheet names in order of appearance.


        .. versionadded:: 0.28.1
        """
        return [sheet.name for sheet in self.sheets]

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        self.close()


class Sheet:
    """
    A sheet object is a member of the :meth:`sheets <xlwings.main.Sheets>` collection:

    >>> import xlwings as xw
    >>> wb = xw.Book()
    >>> wb.sheets[0]
    <Sheet [Book1]Sheet1>
    >>> wb.sheets['Sheet1']
    <Sheet [Book1]Sheet1>
    >>> wb.sheets.add()
    <Sheet [Book1]Sheet2>

    .. versionchanged:: 0.9.0
    """

    def __init__(self, sheet=None, impl=None):
        if impl is None:
            self.impl = books.active.sheets(sheet).impl
        else:
            self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __eq__(self, other):
        return (
            isinstance(other, Sheet)
            and self.book == other.book
            and self.name == other.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash((self.book, self.name))

    @property
    def name(self):
        """Gets or sets the name of the Sheet."""
        return self.impl.name

    @name.setter
    def name(self, value):
        if value in [None, ""]:
            raise ValueError("A sheet name can't be empty.")
        elif any(char in value for char in ["\\", "/", "?", "*", "[", "]"]):
            raise ValueError(
                "A sheet name must not contain any "
                "of the following characters: \\, /, ?, *, [, ]"
            )
        elif len(value) > 31:
            raise ValueError(
                f"The max. length of a sheet name is 31 characters. "
                f"Yours is {len(value)}."
            )
        else:
            self.impl.name = value

    @property
    def names(self):
        """
        Returns a names collection that represents all the sheet-specific names
        (names defined with the "SheetName!" prefix).

        .. versionadded:: 0.9.0

        """
        return Names(impl=self.impl.names)

    @property
    def book(self):
        """Returns the Book of the specified Sheet. Read-only."""
        return Book(impl=self.impl.book)

    @property
    def index(self):
        """Returns the index of the Sheet (1-based as in Excel)."""
        return self.impl.index

    def range(self, cell1, cell2=None):
        """
        Returns a Range object from the active sheet of the active book,
        see :meth:`Range`.

        .. versionadded:: 0.9.0
        """
        if isinstance(cell1, Range):
            if cell1.sheet != self:
                raise ValueError("First range is not on this sheet")
            cell1 = cell1.impl
        if isinstance(cell2, Range):
            if cell2.sheet != self:
                raise ValueError("Second range is not on this sheet")
            cell2 = cell2.impl
        return Range(impl=self.impl.range(cell1, cell2))

    @property
    def cells(self):
        """
        Returns a Range object that represents all the cells on the Sheet
        (not just the cells that are currently in use).

        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.cells)

    def activate(self):
        """Activates the Sheet and returns it."""
        self.book.activate()
        return self.impl.activate()

    def select(self):
        """
        Selects the Sheet. Select only works on the active book.

        .. versionadded:: 0.9.0
        """
        return self.impl.select()

    def clear_contents(self):
        """Clears the content of the whole sheet but leaves the formatting."""
        return self.impl.clear_contents()

    def clear_formats(self):
        """Clears the format of the whole sheet but leaves the content.

        .. versionadded:: 0.26.2
        """
        return self.impl.clear_formats()

    def clear(self):
        """Clears the content and formatting of the whole sheet."""
        return self.impl.clear()

    def autofit(self, axis=None):
        """
        Autofits the width of either columns, rows or both on a whole Sheet.

        Arguments
        ---------
        axis : string, default None
            - To autofit rows, use one of the following: ``rows`` or ``r``
            - To autofit columns, use one of the following: ``columns`` or ``c``
            - To autofit rows and columns, provide no arguments

        Examples
        --------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> wb.sheets['Sheet1'].autofit('c')
        >>> wb.sheets['Sheet1'].autofit('r')
        >>> wb.sheets['Sheet1'].autofit()

        .. versionadded:: 0.2.3
        """
        return self.impl.autofit(axis)

    def delete(self):
        """
        Deletes the Sheet.

        .. versionadded:: 0.6.0
        """
        return self.impl.delete()

    def to_html(self, path=None):
        """
        Export a Sheet as HTML page.

        Parameters
        ----------

        path : str or path-like, default None
            Path where you want to save the HTML file. Defaults to Sheet name in the
            current working directory.


        .. versionadded:: 0.28.1
        """
        path = utils.fspath(path)
        self.impl.to_html(self.name + ".html" if path is None else path)

    def to_pdf(self, path=None, layout=None, show=False, quality="standard"):
        """
        Exports the sheet to a PDF file.

        Parameters
        ----------
        path : str or path-like object, default None
            Path to the PDF file, defaults to the name of the sheet in the same
            directory of the workbook. For unsaved workbooks, it defaults to the current
            working directory instead.

        layout : str or path-like object, default None
            This argument requires xlwings :bdg-secondary:`PRO`.

            Path to a PDF file on which the report will be printed. This is ideal for
            headers and footers as well as borderless printing of graphics/artwork. The
            PDF file either needs to have only 1 page (every report page uses the same
            layout) or otherwise needs the same amount of pages as the report (each
            report page is printed on the respective page in the layout PDF).

            .. versionadded:: 0.24.3

        show : bool, default False
            Once created, open the PDF file with the default application.

            .. versionadded:: 0.24.6

        quality : str, default ``'standard'``
            Quality of the PDF file. Can either be ``'standard'`` or ``'minimum'``.

            .. versionadded:: 0.26.2

        Examples
        --------
        >>> wb = xw.Book()
        >>> sheet = wb.sheets[0]
        >>> sheet['A1'].value = 'PDF'
        >>> sheet.to_pdf()

        See also :meth:`xlwings.Book.to_pdf`

        .. versionadded:: 0.22.3
        """
        return self.book.to_pdf(
            self.name + ".pdf" if path is None else path,
            include=self.index,
            layout=layout,
            show=show,
            quality=quality,
        )

    def copy(self, before=None, after=None, name=None):
        """
        Copy a sheet to the current or a new Book. By default, it places the copied
        sheet after all existing sheets in the current Book. Returns the copied sheet.

        .. versionadded:: 0.22.0

        Arguments
        ---------
        before : sheet object, default None
            The sheet object before which you want to place the sheet

        after : sheet object, default None
            The sheet object after which you want to place the sheet,
            by default it is placed after all existing sheets

        name : str, default None
            The sheet name of the copy

        Returns
        -------
        Sheet object: Sheet
            The copied sheet

        Examples
        --------

        .. code-block:: python

            # Create two books and add a value to the first sheet of the first book
            first_book = xw.Book()
            second_book = xw.Book()
            first_book.sheets[0]['A1'].value = 'some value'

            # Copy to same Book with the default location and name
            first_book.sheets[0].copy()

            # Copy to same Book with custom sheet name
            first_book.sheets[0].copy(name='copied')

            # Copy to second Book requires to use before or after
            first_book.sheets[0].copy(after=second_book.sheets[0])
        """
        # copy() doesn't return the copied sheet object and has an awkward default
        # (copy it to a new workbook if neither before or after are provided),
        # so we're not taking that behavior over
        assert (before is None) or (
            after is None
        ), "you must provide either before or after but not both"
        if (before is None) and (after is None):
            after = self.book.sheets[-1]
        if before:
            target_book = before.book
            before = before.impl
        if after:
            target_book = after.book
            after = after.impl
        if name:
            if name.lower() in (s.name.lower() for s in target_book.sheets):
                raise ValueError(f"Sheet named '{name}' already present in workbook")
        sheet_names_before = {sheet.name for sheet in target_book.sheets}
        self.impl.copy(before=before, after=after)
        sheet_names_after = {sheet.name for sheet in target_book.sheets}
        new_sheet_name = sheet_names_after.difference(sheet_names_before).pop()
        copied_sheet = target_book.sheets[new_sheet_name]
        if name:
            copied_sheet.name = name
        return copied_sheet

    def render_template(self, **data):
        """
        This method requires xlwings :bdg-secondary:`PRO`.

        Replaces all Jinja variables (e.g ``{{ myvar }}``) in the sheet with the keyword
        argument that has the same name. Following variable types are supported:

        strings, numbers, lists, simple dicts, NumPy arrays, Pandas DataFrames,
        PIL Image objects that have a filename and Matplotlib figures.

        .. versionadded:: 0.22.0

        Parameters
        ----------
        data: kwargs
            All key/value pairs that are used in the template.

        Examples
        --------

        >>> import xlwings as xw
        >>> book = xw.Book()
        >>> book.sheets[0]['A1:A2'].value = '{{ myvar }}'
        >>> book.sheets[0].render_template(myvar='test')
        """
        from .pro.reports.main import render_sheet

        render_sheet(self, **data)

    @property
    def charts(self):
        """
        See :meth:`Charts <xlwings.main.Charts>`

        .. versionadded:: 0.9.0
        """
        return Charts(impl=self.impl.charts)

    @property
    def shapes(self):
        """
        See :meth:`Shapes <xlwings.main.Shapes>`

        .. versionadded:: 0.9.0
        """
        return Shapes(impl=self.impl.shapes)

    @property
    def tables(self):
        """
        See :meth:`Tables <xlwings.main.Tables>`

        .. versionadded:: 0.21.0
        """
        return Tables(impl=self.impl.tables)

    @property
    def pictures(self):
        """
        See :meth:`Pictures <xlwings.main.Pictures>`

        .. versionadded:: 0.9.0
        """
        return Pictures(impl=self.impl.pictures)

    @property
    def used_range(self):
        """
        Used Range of Sheet.

        Returns
        -------
        xw.Range


        .. versionadded:: 0.13.0
        """
        return Range(impl=self.impl.used_range)

    @property
    def visible(self):
        """Gets or sets the visibility of the Sheet (bool).

        .. versionadded:: 0.21.1
        """
        return self.impl.visible

    @visible.setter
    def visible(self, value):
        self.impl.visible = value

    @property
    def page_setup(self):
        """
        Returns a PageSetup object.

        .. versionadded:: 0.24.2
        """
        return PageSetup(self.impl.page_setup)

    def __getitem__(self, item):
        if isinstance(item, str):
            return self.range(item)
        else:
            return self.cells[item]

    def __repr__(self):
        return "<Sheet [{1}]{0}>".format(self.name, self.book.name)


class Range:
    """
    Returns a Range object that represents a cell or a range of cells.

    Arguments
    ---------
    cell1 : str or tuple or Range
        Name of the range in the upper-left corner in A1 notation or as index-tuple or
        as name or as xw.Range object. It can also specify a range using the range
        operator (a colon), .e.g. 'A1:B2'

    cell2 : str or tuple or Range, default None
        Name of the range in the lower-right corner in A1 notation or as index-tuple or
        as name or as xw.Range object.

    Examples
    --------

    .. code-block:: python

        import xlwings as xw
        sheet1 = xw.Book("MyBook.xlsx").sheets[0]

        sheet1.range("A1")
        sheet1.range("A1:C3")
        sheet1.range((1,1))
        sheet1.range((1,1), (3,3))
        sheet1.range("NamedRange")

        # Or using index/slice notation
        sheet1["A1"]
        sheet1["A1:C3"]
        sheet1[0, 0]
        sheet1[0:4, 0:4]
        sheet1["NamedRange"]
    """

    def __init__(self, cell1=None, cell2=None, **options):
        # Arguments
        impl = options.pop("impl", None)
        if impl is None:
            if (
                cell2 is not None
                and isinstance(cell1, Range)
                and isinstance(cell2, Range)
            ):
                if cell1.sheet != cell2.sheet:
                    raise ValueError("Ranges are not on the same sheet")
                impl = cell1.sheet.range(cell1, cell2).impl
            elif cell2 is None and isinstance(cell1, str):
                impl = apps.active.range(cell1).impl
            elif cell2 is None and isinstance(cell1, tuple):
                impl = sheets.active.range(cell1, cell2).impl
            elif (
                cell2 is not None
                and isinstance(cell1, tuple)
                and isinstance(cell2, tuple)
            ):
                impl = sheets.active.range(cell1, cell2).impl
            else:
                raise ValueError("Invalid arguments")

        self._impl = impl

        # Keyword Arguments
        self._impl.options = options
        self._options = options

    @property
    def impl(self):
        return self._impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __eq__(self, other):
        return (
            isinstance(other, Range)
            and self.sheet == other.sheet
            and self.row == other.row
            and self.column == other.column
            and self.shape == other.shape
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash((self.sheet, self.row, self.column, self.shape))

    def __iter__(self):
        # Iterator object that returns cell Ranges: (1, 1), (1, 2) etc.
        for i in range(len(self)):
            yield self(i + 1)

    def options(self, convert=None, **options):
        """
        Allows you to set a converter and their options. Converters define how Excel
        Ranges and their values are being converted both during reading and writing
        operations. If no explicit converter is specified, the base converter is being
        applied, see :ref:`converters`.

        Arguments
        ---------
        ``convert`` : object, default None
            A converter, e.g. ``dict``, ``np.array``, ``pd.DataFrame``, ``pd.Series``,
            defaults to default converter

        Keyword Arguments
        -----------------
        ndim : int, default None
            number of dimensions

        numbers : type, default None
            type of numbers, e.g. ``int``

        dates : type, default None
            e.g. ``datetime.date`` defaults to ``datetime.datetime``

        empty : object, default None
            transformation of empty cells

        transpose : Boolean, default False
            transpose values

        expand : str, default None
            One of ``'table'``, ``'down'``, ``'right'``

        chunksize : int
            Use a chunksize, e.g. ``10000`` to prevent timeout or memory issues when
            reading or writing large amounts of data. Works with all formats, including
            DataFrames, NumPy arrays, and list of lists.

        err_to_str : Boolean, default False
            If ``True``, will include cell errors such as ``#N/A`` as strings. By
            default, they will be converted to ``None``.

            .. versionadded:: 0.28.0

        => For converter-specific options, see :ref:`converters`.

        Returns
        -------
        Range object

        """
        options["convert"] = convert
        return Range(impl=self.impl, **options)

    @property
    def sheet(self):
        """
        Returns the Sheet object to which the Range belongs.

        .. versionadded:: 0.9.0
        """
        return Sheet(impl=self.impl.sheet)

    def __len__(self):
        return len(self.impl)

    @property
    def count(self):
        """
        Returns the number of cells.

        """
        return len(self)

    @property
    def row(self):
        """
        Returns the number of the first row in the specified range. Read-only.

        Returns
        -------
        Integer


        .. versionadded:: 0.3.5
        """
        return self.impl.row

    @property
    def column(self):
        """
        Returns the number of the first column in the in the specified range. Read-only.

        Returns
        -------
        Integer


        .. versionadded:: 0.3.5
        """
        return self.impl.column

    @property
    def raw_value(self):
        """
        Gets and sets the values directly as delivered from/accepted by the engine that
        s being used (``pywin32`` or ``appscript``) without going through any of
        xlwings' data cleaning/converting. This can be helpful if speed is an issue but
        naturally will be engine specific, i.e. might remove the cross-platform
        compatibility.
        """
        return self.impl.raw_value

    @raw_value.setter
    def raw_value(self, data):
        self.impl.raw_value = data

    def clear_contents(self):
        """Clears the content of a Range but leaves the formatting."""
        return self.impl.clear_contents()

    def clear_formats(self):
        """Clears the format of a Range but leaves the content.

        .. versionadded:: 0.26.2
        """
        return self.impl.clear_formats()

    def clear(self):
        """Clears the content and the formatting of a Range."""
        return self.impl.clear()

    @property
    def has_array(self):
        """
        ``True`` if the range is part of a legacy CSE Array formula
        and ``False`` otherwise.
        """
        return self.impl.has_array

    def end(self, direction):
        """
        Returns a Range object that represents the cell at the end of the region that
        contains the source range. Equivalent to pressing Ctrl+Up, Ctrl+down,
        Ctrl+left, or Ctrl+right.

        Parameters
        ----------
        direction : One of 'up', 'down', 'right', 'left'

        Examples
        --------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> sheet1 = xw.sheets[0]
        >>> sheet1.range('A1:B2').value = 1
        >>> sheet1.range('A1').end('down')
        <Range [Book1]Sheet1!$A$2>
        >>> sheet1.range('B2').end('right')
        <Range [Book1]Sheet1!$B$2>

        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.end(direction))

    @property
    def formula(self):
        """Gets or sets the formula for the given Range."""
        return self.impl.formula

    @formula.setter
    def formula(self, value):
        self.impl.formula = value

    @property
    def formula2(self):
        """Gets or sets the formula2 for the given Range."""
        return self.impl.formula2

    @formula2.setter
    def formula2(self, value):
        self.impl.formula2 = value

    @property
    def formula_array(self):
        """
        Gets or sets an  array formula for the given Range.

        .. versionadded:: 0.7.1
        """
        return self.impl.formula_array

    @formula_array.setter
    def formula_array(self, value):
        self.impl.formula_array = value

    @property
    def font(self):
        return Font(impl=self.impl.font)

    @property
    def characters(self):
        return Characters(impl=self.impl.characters)

    @property
    def column_width(self):
        """
        Gets or sets the width, in characters, of a Range.
        One unit of column width is equal to the width of one character in the Normal
        style. For proportional fonts, the width of the character 0 (zero) is used.

        If all columns in the Range have the same width, returns the width.
        If columns in the Range have different widths, returns None.

        column_width must be in the range:
        0 <= column_width <= 255

        Note: If the Range is outside the used range of the Worksheet, and columns in
        the Range have different widths, returns the width of the first column.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return self.impl.column_width

    @column_width.setter
    def column_width(self, value):
        self.impl.column_width = value

    @property
    def row_height(self):
        """
        Gets or sets the height, in points, of a Range.
        If all rows in the Range have the same height, returns the height.
        If rows in the Range have different heights, returns None.

        row_height must be in the range:
        0 <= row_height <= 409.5

        Note: If the Range is outside the used range of the Worksheet, and rows in the
        Range have different heights, returns the height of the first row.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return self.impl.row_height

    @row_height.setter
    def row_height(self, value):
        self.impl.row_height = value

    @property
    def width(self):
        """
        Returns the width, in points, of a Range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return self.impl.width

    @property
    def height(self):
        """
        Returns the height, in points, of a Range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.4.0
        """
        return self.impl.height

    @property
    def left(self):
        """
        Returns the distance, in points, from the left edge of column A to the left
        edge of the range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.6.0
        """
        return self.impl.left

    @property
    def top(self):
        """
        Returns the distance, in points, from the top edge of row 1 to the top edge of
        the range. Read-only.

        Returns
        -------
        float


        .. versionadded:: 0.6.0
        """
        return self.impl.top

    @property
    def number_format(self):
        """
        Gets and sets the number_format of a Range.

        Examples
        --------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> sheet1 = wb.sheets[0]
        >>> sheet1.range('A1').number_format
        'General'
        >>> sheet1.range('A1:C3').number_format = '0.00%'
        >>> sheet1.range('A1:C3').number_format
        '0.00%'

        .. versionadded:: 0.2.3
        """
        return self.impl.number_format

    @number_format.setter
    def number_format(self, value):
        self.impl.number_format = value

    def get_address(
        self,
        row_absolute=True,
        column_absolute=True,
        include_sheetname=False,
        external=False,
    ):
        """
        Returns the address of the range in the specified format. ``address`` can be
        used instead if none of the defaults need to be changed.

        Arguments
        ---------
        row_absolute : bool, default True
            Set to True to return the row part of the reference as an absolute
            reference.

        column_absolute : bool, default True
            Set to True to return the column part of the reference as an absolute
            reference.

        include_sheetname : bool, default False
            Set to True to include the Sheet name in the address. Ignored if
            external=True.

        external : bool, default False
            Set to True to return an external reference with workbook and worksheet
            name.

        Returns
        -------
        str

        Examples
        --------
        ::

            >>> import xlwings as xw
            >>> wb = xw.Book()
            >>> sheet1 = wb.sheets[0]
            >>> sheet1.range((1,1)).get_address()
            '$A$1'
            >>> sheet1.range((1,1)).get_address(False, False)
            'A1'
            >>> sheet1.range((1,1), (3,3)).get_address(True, False, True)
            'Sheet1!A$1:C$3'
            >>> sheet1.range((1,1), (3,3)).get_address(True, False, external=True)
            '[Book1]Sheet1!A$1:C$3'

        .. versionadded:: 0.2.3
        """

        if include_sheetname and not external:
            # TODO: when the Workbook name contains spaces but not the Worksheet name,
            #  it will still be surrounded
            # by '' when include_sheetname=True. Also, should probably changed to regex
            temp_str = self.impl.get_address(row_absolute, column_absolute, True)

            if temp_str.find("[") > -1:
                results_address = temp_str[temp_str.rfind("]") + 1 :]
                if results_address.find("'") > -1:
                    results_address = "'" + results_address
                return results_address
            else:
                return temp_str

        else:
            return self.impl.get_address(row_absolute, column_absolute, external)

    @property
    def address(self):
        """
        Returns a string value that represents the range reference.
        Use ``get_address()`` to be able to provide parameters.

        .. versionadded:: 0.9.0
        """
        return self.impl.address

    @property
    def current_region(self):
        """
        This property returns a Range object representing a range bounded by (but not
        including) any combination of blank rows and blank columns or the edges of the
        worksheet. It corresponds to ``Ctrl-*`` on Windows and ``Shift-Ctrl-Space`` on
        Mac.

        Returns
        -------
        Range object
        """

        return Range(impl=self.impl.current_region)

    def autofit(self):
        """
        Autofits the width and height of all cells in the range.

        * To autofit only the width of the columns use
          ``myrange.columns.autofit()``
        * To autofit only the height of the rows use
          ``myrange.rows.autofit()``

        .. versionchanged:: 0.9.0
        """
        return self.impl.autofit()

    @property
    def color(self):
        """
        Gets and sets the background color of the specified Range.

        To set the color, either use an RGB tuple ``(0, 0, 0)`` or a hex string
        like ``#efefef`` or an Excel color constant.
        To remove the background, set the color to ``None``, see Examples.

        Returns
        -------
        RGB : tuple

        Examples
        --------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> sheet1 = xw.sheets[0]
        >>> sheet1.range('A1').color = (255, 255, 255)  # or '#ffffff'
        >>> sheet1.range('A2').color
        (255, 255, 255)
        >>> sheet1.range('A2').color = None
        >>> sheet1.range('A2').color is None
        True

        .. versionadded:: 0.3.0
        """
        return self.impl.color

    @color.setter
    def color(self, color_or_rgb):
        self.impl.color = color_or_rgb

    @property
    def name(self):
        """
        Sets or gets the name of a Range.

        .. versionadded:: 0.4.0
        """
        impl = self.impl.name
        return impl and Name(impl=impl)

    @name.setter
    def name(self, value):
        self.impl.name = value

    def __call__(self, *args):
        return Range(impl=self.impl(*args))

    @property
    def rows(self):
        """
        Returns a :class:`RangeRows` object that represents the rows in the specified
        range.

        .. versionadded:: 0.9.0
        """
        return RangeRows(self)

    @property
    def columns(self):
        """
        Returns a :class:`RangeColumns` object that represents the columns in the
        specified range.

        .. versionadded:: 0.9.0
        """
        return RangeColumns(self)

    @property
    def shape(self):
        """
        Tuple of Range dimensions.

        .. versionadded:: 0.3.0
        """
        return self.impl.shape

    @property
    def size(self):
        """
        Number of elements in the Range.

        .. versionadded:: 0.3.0
        """
        a, b = self.shape
        return a * b

    @property
    def value(self):
        """
        Gets and sets the values for the given Range. See :meth:`xlwings.Range.options`
        about how to set options, e.g., to transform it into a DataFrame or how to set
        a chunksize.

        Returns
        -------
        object : returned object depends on the converter being used,
                 see :meth:`xlwings.Range.options`
        """
        return conversion.read(self, None, self._options)

    @value.setter
    def value(self, data):
        conversion.write(data, self, self._options)

    def expand(self, mode="table"):
        """
        Expands the range according to the mode provided. Ignores empty top-left cells
        (unlike ``Range.end()``).

        Parameters
        ----------
        mode : str, default 'table'
            One of ``'table'`` (=down and right), ``'down'``, ``'right'``.

        Returns
        -------
        Range

        Examples
        --------

        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> sheet1 = wb.sheets[0]
        >>> sheet1.range('A1').value = [[None, 1], [2, 3]]
        >>> sheet1.range('A1').expand().address
        $A$1:$B$2
        >>> sheet1.range('A1').expand('right').address
        $A$1:$B$1

        .. versionadded:: 0.9.0
        """
        return expansion.expanders.get(mode, mode).expand(self)

    def __getitem__(self, key):
        if type(key) is tuple:
            row, col = key

            n = self.shape[0]
            if isinstance(row, slice):
                row1, row2, step = row.indices(n)
                if step != 1:
                    raise ValueError("Slice steps not supported.")
                row2 -= 1
            elif isinstance(row, int):
                if row < 0:
                    row += n
                if row < 0 or row >= n:
                    raise IndexError("Row index %s out of range (%s rows)." % (row, n))
                row1 = row2 = row
            else:
                raise TypeError(
                    "Row indices must be integers or slices, not %s"
                    % type(row).__name__
                )

            n = self.shape[1]
            if isinstance(col, slice):
                col1, col2, step = col.indices(n)
                if step != 1:
                    raise ValueError("Slice steps not supported.")
                col2 -= 1
            elif isinstance(col, int):
                if col < 0:
                    col += n
                if col < 0 or col >= n:
                    raise IndexError(
                        "Column index %s out of range (%s columns)." % (col, n)
                    )
                col1 = col2 = col
            else:
                raise TypeError(
                    "Column indices must be integers or slices, not %s"
                    % type(col).__name__
                )

            return self.sheet.range(
                (
                    self.row + row1,
                    self.column + col1,
                    max(0, row2 - row1 + 1),
                    max(0, col2 - col1 + 1),
                )
            )

        elif isinstance(key, slice):
            if self.shape[0] > 1 and self.shape[1] > 1:
                raise IndexError(
                    "One-dimensional slicing is not allowed on two-dimensional ranges"
                )

            if self.shape[0] > 1:
                return self[key, :]
            else:
                return self[:, key]

        elif isinstance(key, int):
            n = len(self)
            k = key + n if key < 0 else key
            if k < 0 or k >= n:
                raise IndexError("Index %s out of range (%s elements)." % (key, n))
            else:
                return self(k + 1)

        else:
            raise TypeError(
                "Cell indices must be integers or slices, not %s" % type(key).__name__
            )

    def __repr__(self):
        return "<Range [{1}]{0}!{2}>".format(
            self.sheet.name, self.sheet.book.name, self.address
        )

    def insert(self, shift, copy_origin="format_from_left_or_above"):
        """
        Insert a cell or range of cells into the sheet.

        Parameters
        ----------
        shift : str
            Use ``right`` or ``down``.
        copy_origin : str, default format_from_left_or_above
            Use ``format_from_left_or_above`` or ``format_from_right_or_below``.
            Note that copy_origin is only supported on Windows.

        Returns
        -------
        None


        .. versionchanged:: 0.30.3
            ``shift`` is now a required argument.

        """
        self.impl.insert(shift, copy_origin)

    def delete(self, shift=None):
        """
        Deletes a cell or range of cells.

        Parameters
        ----------
        shift : str, default None
            Use ``left`` or ``up``. If omitted, Excel decides based on the shape of
            the range.

        Returns
        -------
        None

        """
        self.impl.delete(shift)

    def copy(self, destination=None):
        """
        Copy a range to a destination range or clipboard.

        Parameters
        ----------
        destination : xlwings.Range
            xlwings Range to which the specified range will be copied. If omitted,
            the range is copied to the clipboard.

        Returns
        -------
        None

        """
        self.impl.copy(destination)

    def paste(self, paste=None, operation=None, skip_blanks=False, transpose=False):
        """
        Pastes a range from the clipboard into the specified range.

        Parameters
        ----------
        paste : str, default None
            One of ``all_merging_conditional_formats``, ``all``, ``all_except_borders``,
            ``all_using_source_theme``, ``column_widths``, ``comments``, ``formats``,
            ``formulas``, ``formulas_and_number_formats``, ``validation``, ``values``,
            ``values_and_number_formats``.
        operation : str, default None
            One of "add", "divide", "multiply", "subtract".
        skip_blanks : bool, default False
            Set to ``True`` to skip over blank cells
        transpose : bool, default False
            Set to ``True`` to transpose rows and columns.

        Returns
        -------
        None

        """
        self.impl.paste(
            paste=paste,
            operation=operation,
            skip_blanks=skip_blanks,
            transpose=transpose,
        )

    @property
    def hyperlink(self):
        """
        Returns the hyperlink address of the specified Range (single Cell only)

        Examples
        --------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> sheet1 = wb.sheets[0]
        >>> sheet1.range('A1').value
        'www.xlwings.org'
        >>> sheet1.range('A1').hyperlink
        'http://www.xlwings.org'

        .. versionadded:: 0.3.0
        """
        if self.formula.lower().startswith("="):
            # If it's a formula, extract the URL from the formula string
            formula = self.formula
            try:
                return re.compile(r"\"(.+?)\"").search(formula).group(1)
            except AttributeError:
                raise Exception("The cell doesn't seem to contain a hyperlink!")
        else:
            # If it has been set pragmatically
            return self.impl.hyperlink

    def add_hyperlink(self, address, text_to_display=None, screen_tip=None):
        """
        Adds a hyperlink to the specified Range (single Cell)

        Arguments
        ---------
        address : str
            The address of the hyperlink.
        text_to_display : str, default None
            The text to be displayed for the hyperlink. Defaults to the hyperlink
            address.
        screen_tip: str, default None
            The screen tip to be displayed when the mouse pointer is paused over the
            hyperlink. Default is set to '<address> - Click once to follow. Click and
            hold to select this cell.'


        .. versionadded:: 0.3.0
        """
        if text_to_display is None:
            text_to_display = address
        if address[:4] == "www.":
            address = "http://" + address
        if screen_tip is None:
            screen_tip = (
                address + " - Click once to follow. Click and hold to select this cell."
            )
        self.impl.add_hyperlink(address, text_to_display, screen_tip)

    def resize(self, row_size=None, column_size=None):
        """
        Resizes the specified Range

        Arguments
        ---------
        row_size: int > 0
            The number of rows in the new range (if None, the number of rows in the
            range is unchanged).
        column_size: int > 0
            The number of columns in the new range (if None, the number of columns in
            the range is unchanged).

        Returns
        -------
        Range object: Range


        .. versionadded:: 0.3.0
        """

        if row_size is not None:
            assert row_size > 0
        else:
            row_size = self.shape[0]
        if column_size is not None:
            assert column_size > 0
        else:
            column_size = self.shape[1]

        return Range(self(1, 1), self(row_size, column_size)).options(**self._options)

    def offset(self, row_offset=0, column_offset=0):
        """
        Returns a Range object that represents a Range that's offset from the
        specified range.

        Returns
        -------
        Range object : Range


        .. versionadded:: 0.3.0
        """
        return Range(
            self(row_offset + 1, column_offset + 1),
            self(row_offset + self.shape[0], column_offset + self.shape[1]),
        ).options(**self._options)

    @property
    def last_cell(self):
        """
        Returns the bottom right cell of the specified range. Read-only.

        Returns
        -------
        Range

        Example
        -------
        >>> import xlwings as xw
        >>> wb = xw.Book()
        >>> sheet1 = wb.sheets[0]
        >>> myrange = sheet1.range('A1:E4')
        >>> myrange.last_cell.row, myrange.last_cell.column
        (4, 5)

        .. versionadded:: 0.3.5
        """
        return self(self.shape[0], self.shape[1]).options(**self._options)

    def select(self):
        """
        Selects the range. Select only works on the active book.

        .. versionadded:: 0.9.0
        """
        self.impl.select()

    @property
    def merge_area(self):
        """
        Returns a Range object that represents the merged Range containing the
        specified cell. If the specified cell isn't in a merged range, this property
        returns the specified cell.

        """
        return Range(impl=self.impl.merge_area)

    @property
    def merge_cells(self):
        """
        Returns ``True`` if the Range contains merged cells, otherwise ``False``
        """
        return self.impl.merge_cells

    def merge(self, across=False):
        """
        Creates a merged cell from the specified Range object.

        Parameters
        ----------
        across : bool, default False
            True to merge cells in each row of the specified Range as separate merged
            cells.
        """
        with self.sheet.book.app.properties(display_alerts=False):
            self.impl.merge(across)

    def unmerge(self):
        """
        Separates a merged area into individual cells.
        """
        self.impl.unmerge()

    @property
    def table(self):
        """
        Returns a Table object if the range is part of one, otherwise ``None``.

        .. versionadded:: 0.21.0
        """
        if self.impl.table:
            return Table(impl=self.impl.table)
        else:
            return None

    @property
    def wrap_text(self):
        """
        Returns ``True`` if the wrap_text property is enabled and ``False`` if it's
        disabled. If not all cells have the same value in a range, on Windows it returns
        ``None`` and on macOS ``False``.

        .. versionadded:: 0.23.2
        """
        return self.impl.wrap_text

    @wrap_text.setter
    def wrap_text(self, value):
        self.impl.wrap_text = value

    @property
    def note(self):
        """
        Returns a Note object.
        Before the introduction of threaded comments, a Note was called a Comment.

        .. versionadded:: 0.24.2
        """
        return Note(impl=self.impl.note) if self.impl.note else None

    def copy_picture(self, appearance="screen", format="picture"):
        """
        Copies the range to the clipboard as picture.

        Parameters
        ----------
        appearance : str, default 'screen'
            Either 'screen' or 'printer'.

        format : str, default 'picture'
            Either 'picture' or 'bitmap'.


        .. versionadded:: 0.24.8
        """
        self.impl.copy_picture(appearance, format)

    def to_png(self, path=None):
        """
        Exports the range as PNG picture.

        Parameters
        ----------

        path : str or path-like, default None
            Path where you want to store the picture. Defaults to the name of the range
            in the same directory as the Excel file if the Excel file is stored and to
            the current working directory otherwise.


        .. versionadded:: 0.24.8
        """
        if not PIL:
            raise XlwingsError("Range.to_png() requires an installation of Pillow.")
        path = utils.fspath(path)
        if path is None:
            # TODO: factor this out as it's used in multiple locations
            directory, _ = os.path.split(self.sheet.book.fullname)
            default_name = (
                str(self)
                .replace("<", "")
                .replace(">", "")
                .replace(":", "_")
                .replace(" ", "")
            )
            if directory:
                path = os.path.join(directory, default_name + ".png")
            else:
                path = str(Path.cwd() / default_name) + ".png"
        self.impl.to_png(path)

    def to_pdf(self, path=None, layout=None, show=None, quality="standard"):
        """
        Exports the range as PDF.

        Parameters
        ----------

        path : str or path-like, default None
            Path where you want to store the pdf. Defaults to the address of the range
            in the same directory as the Excel file if the Excel file is stored and to
            the current working directory otherwise.

        layout : str or path-like object, default None
            This argument requires xlwings :bdg-secondary:`PRO`.

            Path to a PDF file on which the report will be printed. This is ideal for
            headers and footers as well as borderless printing of graphics/artwork. The
            PDF file either needs to have only 1 page (every report page uses the same
            layout) or otherwise needs the same amount of pages as the report (each
            report page is printed on the respective page in the layout PDF).

        show : bool, default False
            Once created, open the PDF file with the default application.

        quality : str, default ``'standard'``
            Quality of the PDF file. Can either be ``'standard'`` or ``'minimum'``.


        .. versionadded:: 0.26.2
        """
        return utils.to_pdf(self, path=path, layout=layout, show=show, quality=quality)

    def autofill(self, destination, type_="fill_default"):
        """
        Autofills the destination Range. Note that the destination Range must include
        the origin Range.

        Arguments
        ---------

        destination : Range
            The origin.

        type_ : str, default ``"fill_default"``
            One of the following strings: ``"fill_copy"``, ``"fill_days"``,
            ``"fill_default"``, ``"fill_formats"``, ``"fill_months"``,
            ``"fill_series"``, ``"fill_values"``, ``"fill_weekdays"``, ``"fill_years"``,
            ``"growth_trend"``, ``"linear_trend"``, ``"flash_fill``


        .. versionadded:: 0.30.1
        """
        self.impl.autofill(destination=destination, type_=type_)


# These have to be after definition of Range to resolve circular reference
from . import conversion, expansion


class Ranges:
    pass


class RangeRows(Ranges):
    """
    Represents the rows of a range. Do not construct this class directly, use
    :attr:`Range.rows` instead.

    Example
    -------

    .. code-block:: python

        import xlwings as xw

        wb = xw.Book("MyBook.xlsx")
        sheet1 = wb.sheets[0]
        myrange = sheet1.range('A1:C4')

        assert len(myrange.rows) == 4  # or myrange.rows.count

        myrange.rows[0].value = 'a'

        assert myrange.rows[2] == sheet1.range('A3:C3')
        assert myrange.rows(2) == sheet1.range('A2:C2')

        for r in myrange.rows:
            print(r.address)
    """

    def __init__(self, rng):
        self.rng = rng

    def __len__(self):
        """
        Returns the number of rows.

        .. versionadded:: 0.9.0
        """
        return self.rng.shape[0]

    count = property(__len__)

    def autofit(self):
        """
        Autofits the height of the rows.
        """
        self.rng.impl.autofit(axis="r")

    def __iter__(self):
        for i in range(0, self.rng.shape[0]):
            yield self.rng[i, :]

    def __call__(self, key):
        return self.rng[key - 1, :]

    def __getitem__(self, key):
        if isinstance(key, slice):
            return RangeRows(rng=self.rng[key, :])
        elif isinstance(key, int):
            return self.rng[key, :]
        else:
            raise TypeError(
                "Indices must be integers or slices, not %s" % type(key).__name__
            )

    def __repr__(self):
        return "{}({})".format(self.__class__.__name__, repr(self.rng))


class RangeColumns(Ranges):
    """
    Represents the columns of a range. Do not construct this class directly, use
    :attr:`Range.columns` instead.

    Example
    -------

    .. code-block:: python

        import xlwings as xw

        wb = xw.Book("MyFile.xlsx")
        sheet1 = wb.sheets[0]
        myrange = sheet1.range('A1:C4')

        assert len(myrange.columns) == 3  # or myrange.columns.count

        myrange.columns[0].value = 'a'

        assert myrange.columns[2] == sheet1.range('C1:C4')
        assert myrange.columns(2) == sheet1.range('B1:B4')

        for c in myrange.columns:
            print(c.address)
    """

    def __init__(self, rng):
        self.rng = rng

    def __len__(self):
        """
        Returns the number of columns.

        .. versionadded:: 0.9.0
        """
        return self.rng.shape[1]

    count = property(__len__)

    def autofit(self):
        """
        Autofits the width of the columns.
        """
        self.rng.impl.autofit(axis="c")

    def __iter__(self):
        for j in range(0, self.rng.shape[1]):
            yield self.rng[:, j]

    def __call__(self, key):
        return self.rng[:, key - 1]

    def __getitem__(self, key):
        if isinstance(key, slice):
            return RangeColumns(rng=self.rng[:, key])
        elif isinstance(key, int):
            return self.rng[:, key]
        else:
            raise TypeError(
                "Indices must be integers or slices, not %s" % type(key).__name__
            )

    def __repr__(self):
        return "{}({})".format(self.__class__.__name__, repr(self.rng))


class Shape:
    """
    The shape object is a member of the :meth:`shapes <xlwings.main.Shapes>` collection:

    >>> import xlwings as xw
    >>> sht = xw.books['Book1'].sheets[0]
    >>> sht.shapes[0]  # or sht.shapes['ShapeName']
    <Shape 'Rectangle 1' in <Sheet [Book1]Sheet1>>

    .. versionchanged:: 0.9.0
    """

    def __init__(self, *args, **options):
        impl = options.pop("impl", None)
        if impl is None:
            if len(args) == 1:
                impl = sheets.active.shapes(args[0]).impl

            else:
                raise ValueError("Invalid arguments")

        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine
        being used.

        .. versionadded:: 0.19.2
        """
        return self.impl.api

    @property
    def name(self):
        """
        Returns or sets the name of the shape.

        .. versionadded:: 0.5.0
        """
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def type(self):
        """
        Returns the type of the shape.

        .. versionadded:: 0.9.0
        """
        return self.impl.type

    @property
    def left(self):
        """
        Returns or sets the number of points that represent the horizontal position of
        the shape.

        .. versionadded:: 0.5.0
        """
        return self.impl.left

    @left.setter
    def left(self, value):
        self.impl.left = value

    @property
    def top(self):
        """
        Returns or sets the number of points that represent the vertical position of
        the shape.

        .. versionadded:: 0.5.0
        """
        return self.impl.top

    @top.setter
    def top(self, value):
        self.impl.top = value

    @property
    def width(self):
        """
        Returns or sets the number of points that represent the width of the shape.

        .. versionadded:: 0.5.0
        """
        return self.impl.width

    @width.setter
    def width(self, value):
        self.impl.width = value

    @property
    def height(self):
        """
        Returns or sets the number of points that represent the height of the shape.

        .. versionadded:: 0.5.0
        """
        return self.impl.height

    @height.setter
    def height(self, value):
        self.impl.height = value

    def delete(self):
        """
        Deletes the shape.

        .. versionadded:: 0.5.0
        """
        self.impl.delete()

    def activate(self):
        """
        Activates the shape.

        .. versionadded:: 0.5.0
        """
        self.impl.activate()

    def scale_height(
        self, factor, relative_to_original_size=False, scale="scale_from_top_left"
    ):
        """
        factor : float
            For example 1.5 to scale it up to 150%

        relative_to_original_size : bool, optional
            If ``False``, it scales relative to current height (default).
            For ``True`` must be a picture or OLE object.

        scale : str, optional
            One of ``scale_from_top_left`` (default), ``scale_from_bottom_right``,
            ``scale_from_middle``

        .. versionadded:: 0.19.2
        """
        self.impl.scale_height(
            factor=factor,
            relative_to_original_size=relative_to_original_size,
            scale=scale,
        )

    def scale_width(
        self, factor, relative_to_original_size=False, scale="scale_from_top_left"
    ):
        """
        factor : float
            For example 1.5 to scale it up to 150%

        relative_to_original_size : bool, optional
            If ``False``, it scales relative to current width (default).
            For ``True`` must be a picture or OLE object.

        scale : str, optional
            One of ``scale_from_top_left`` (default), ``scale_from_bottom_right``,
            ``scale_from_middle``

        .. versionadded:: 0.19.2
        """
        self.impl.scale_width(
            factor=factor,
            relative_to_original_size=relative_to_original_size,
            scale=scale,
        )

    @property
    def text(self):
        """
        Returns or sets the text of a shape.

        .. versionadded:: 0.21.4
        """
        return self.impl.text

    @text.setter
    def text(self, value):
        if xlwings.__pro__:
            from xlwings.pro import Markdown
            from xlwings.pro.reports.markdown import format_text, render_text

            if isinstance(value, Markdown):
                self.impl.text = render_text(value.text, value.style)
                format_text(self, value.text, value.style)
            else:
                self.impl.text = value
        else:
            self.impl.text = value

    @property
    def font(self):
        return Font(impl=self.impl.font)

    @property
    def characters(self):
        return Characters(impl=self.impl.characters)

    @property
    def parent(self):
        """
        Returns the parent of the shape.

        .. versionadded:: 0.9.0
        """
        return Sheet(impl=self.impl.parent)

    def __eq__(self, other):
        return (
            isinstance(other, Shape)
            and other.parent == self.parent
            and other.name == self.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __repr__(self):
        return "<Shape '{0}' in {1}>".format(self.name, self.parent)


class Shapes(Collection):
    """
    A collection of all :meth:`shape <Shape>` objects on the specified sheet:

    >>> import xlwings as xw
    >>> xw.books['Book1'].sheets[0].shapes
    Shapes([<Shape 'Oval 1' in <Sheet [Book1]Sheet1>>,
            <Shape 'Rectangle 1' in <Sheet [Book1]Sheet1>>])

    .. versionadded:: 0.9.0
    """

    _wrap = Shape


class PageSetup:
    def __init__(self, impl):
        """
        Represents a PageSetup object.

        .. versionadded:: 0.24.2
        """
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.24.2
        """
        return self.impl.api

    @property
    def print_area(self):
        """
        Gets or sets the range address that defines the print area.

        Examples
        --------

        >>> mysheet.page_setup.print_area = '$A$1:$B$3'
        >>> mysheet.page_setup.print_area
        '$A$1:$B$3'
        >>> mysheet.page_setup.print_area = None  # clear the print_area

        .. versionadded:: 0.24.2
        """
        return self.impl.print_area

    @print_area.setter
    def print_area(self, value):
        self.impl.print_area = value


class Note:
    def __init__(self, impl):
        """
        Represents a cell Note.
        Before the introduction of threaded comments, a Note was called a Comment.

        .. versionadded:: 0.24.2
        """
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.24.2
        """
        return self.impl.api

    @property
    def text(self):
        """
        Gets or sets the text of a note. Keep in mind that the note must already exist!

        Examples
        --------

        >>> sheet = xw.Book(...).sheets[0]
        >>> sheet['A1'].note.text = 'mynote'
        >>> sheet['A1'].note.text
        >>> 'mynote'

        .. versionadded:: 0.24.2
        """
        return self.impl.text

    @text.setter
    def text(self, value):
        self.impl.text = value

    def delete(self):
        """
        Delete the note.

        .. versionadded:: 0.24.2
        """
        self.impl.delete()


class Table:
    """
    The table object is a member of the :meth:`tables <xlwings.main.Tables>` collection:

    >>> import xlwings as xw
    >>> sht = xw.books['Book1'].sheets[0]
    >>> sht.tables[0]  # or sht.tables['TableName']
    <Table 'Table 1' in <Sheet [Book1]Sheet1>>

    .. versionadded:: 0.21.0
    """

    def __init__(self, *args, **options):
        impl = options.pop("impl", None)
        if impl is None:
            if len(args) == 1:
                impl = sheets.active.tables(args[0]).impl
            else:
                raise ValueError("Invalid arguments")
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl.api

    @property
    def parent(self):
        """
        Returns the parent of the table.
        """
        return Sheet(impl=self.impl.parent)

    @property
    def name(self):
        """
        Returns or sets the name of the Table.
        """
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def data_body_range(self):
        """Returns an xlwings range object that represents the range of values,
        excluding the header row
        """
        return (
            Range(impl=self.impl.data_body_range) if self.impl.data_body_range else None
        )

    @property
    def display_name(self):
        """Returns or sets the display name for the specified Table object"""
        return self.impl.display_name

    @display_name.setter
    def display_name(self, value):
        self.impl.display_name = value

    @property
    def header_row_range(self):
        """Returns an xlwings range object that represents the range of the header row"""
        if self.impl.header_row_range:
            return Range(impl=self.impl.header_row_range)
        else:
            return None

    @property
    def insert_row_range(self):
        """Returns an xlwings range object representing the row where data is going to
        be inserted. This is only available for empty tables, otherwise it'll return
        ``None``
        """
        if self.impl.insert_row_range:
            return Range(impl=self.impl.insert_row_range)
        else:
            return None

    @property
    def range(self):
        """Returns an xlwings range object of the table."""
        return Range(impl=self.impl.range)

    @property
    def show_autofilter(self):
        """Turn the autofilter on or off by setting it to ``True`` or ``False``
        (read/write boolean)
        """
        return self.impl.show_autofilter

    @show_autofilter.setter
    def show_autofilter(self, value):
        self.impl.show_autofilter = value

    @property
    def show_headers(self):
        """Show or hide the header (read/write)"""
        return self.impl.show_headers

    @show_headers.setter
    def show_headers(self, value):
        self.impl.show_headers = value

    @property
    def show_table_style_column_stripes(self):
        """Returns or sets if the Column Stripes table style is used for
        (read/write boolean)
        """
        return self.impl.show_table_style_column_stripes

    @show_table_style_column_stripes.setter
    def show_table_style_column_stripes(self, value):
        self.impl.show_table_style_column_stripes = value

    @property
    def show_table_style_first_column(self):
        """Returns or sets if the first column is formatted (read/write boolean)"""
        return self.impl.show_table_style_first_column

    @show_table_style_first_column.setter
    def show_table_style_first_column(self, value):
        self.impl.show_table_style_first_column = value

    @property
    def show_table_style_last_column(self):
        """Returns or sets if the last column is displayed (read/write boolean)"""
        return self.impl.show_table_style_last_column

    @show_table_style_last_column.setter
    def show_table_style_last_column(self, value):
        self.impl.show_table_style_last_column = value

    @property
    def show_table_style_row_stripes(self):
        """Returns or sets if the Row Stripes table style is used
        (read/write boolean)
        """
        return self.impl.show_table_style_row_stripes

    @show_table_style_row_stripes.setter
    def show_table_style_row_stripes(self, value):
        self.impl.show_table_style_row_stripes = value

    @property
    def show_totals(self):
        """Gets or sets a boolean to show/hide the Total row."""
        return self.impl.show_totals

    @show_totals.setter
    def show_totals(self, value):
        self.impl.show_totals = value

    @property
    def table_style(self):
        """Gets or sets the table style.
        See :meth:`Tables.add <xlwings.main.Tables.add>` for possible values.
        """
        return self.impl.table_style

    @table_style.setter
    def table_style(self, value):
        self.impl.table_style = value

    @property
    def totals_row_range(self):
        """Returns an xlwings range object representing the Total row"""
        if self.impl.totals_row_range:
            return Range(impl=self.impl.totals_row_range)
        else:
            return None

    def update(self, data, index=True):
        """
        Updates the Excel table with the provided data.
        Currently restricted to DataFrames.

        .. versionchanged:: 0.24.0

        Arguments
        ---------

        data : pandas DataFrame
            Currently restricted to pandas DataFrames.
        index : bool, default True
            Whether or not the index of a pandas DataFrame should be written to the
            Excel table.

        Returns
        -------
        Table

        Examples
        --------

        .. code-block:: python

            import pandas as pd
            import xlwings as xw

            sheet = xw.Book('Book1.xlsx').sheets[0]
            table_name = 'mytable'

            # Sample DataFrame
            nrows, ncols = 3, 3
            df = pd.DataFrame(data=nrows * [ncols * ['test']],
                              columns=['col ' + str(i) for i in range(ncols)])

            # Hide the index, then insert a new table if it doesn't exist yet,
            # otherwise update the existing one
            df = df.set_index('col 0')
            if table_name in [table.name for table in sheet.tables]:
                sheet.tables[table_name].update(df)
            else:
                mytable = sheet.tables.add(source=sheet['A1'],
                                           name=table_name).update(df)
        """
        type_error_msg = "Currently, only pandas DataFrames are supported by update"
        if pd:
            if not isinstance(data, pd.DataFrame):
                raise TypeError(type_error_msg)
            if data.empty:
                nrows_data = 1
            else:
                nrows_data = len(data)
            nrows_table = len(self.data_body_range.rows) if self.data_body_range else 1
            row_diff = nrows_table - nrows_data
            if data.empty:
                ncols_data = 1
            else:
                ncols_data = (
                    len(data.columns)
                    if not index
                    else len(data.columns) + len(data.index.names)
                )
            ncols_table = len(self.range.columns)
            col_diff = ncols_table - ncols_data
            cols_to_be_cleared = None
            if col_diff > 0:
                cols_to_be_cleared = self.range[:, ncols_table - col_diff :]
            rows_to_be_cleared = None
            if row_diff > 0 and self.data_body_range:
                rows_to_be_cleared = self.data_body_range[nrows_table - row_diff :, :]
            self.resize(
                self.range[0, 0].resize(
                    row_size=nrows_data + 1 if self.header_row_range else nrows_data,
                    column_size=ncols_data,
                )
            )
            # Clearing must happen after resizing as table headers will be replaced
            # with Column1 etc. if deleted while still being part of table
            if cols_to_be_cleared:
                cols_to_be_cleared.clear_contents()
            if rows_to_be_cleared:
                rows_to_be_cleared.clear_contents()
            if self.header_row_range:
                # Tables with 'Header Row' checked
                header = (
                    (list(data.index.names) + list(data.columns))
                    if index
                    else list(data.columns)
                )
                # Replace None in the header with a unique number of spaces
                n_empty = len([i for i in header if isinstance(i, str) and i.isspace()])
                header = [
                    " " * (i + n_empty + 1) if name is None else name
                    for i, name in enumerate(header)
                ]
                self.header_row_range.value = header
                self.range[1, 0].options(index=index, header=False).value = data
            else:
                # Tables with 'Header Row' unchecked
                self.resize(self.range[0, 0])  # Otherwise the table will be deleted
                self.range[0, 0].options(index=index, header=False).value = data
                # If the top-left cell isn't empty, it doesn't manage to resize the
                # columns automatically
                self.resize(
                    self.range[0, 0].resize(row_size=nrows_data, column_size=ncols_data)
                )
            return self
        else:
            raise TypeError(type_error_msg)

    def resize(self, range):
        """Resize a Table by providing an xlwings range object

        .. versionadded:: 0.24.4
        """
        self.impl.resize(range)

    def __eq__(self, other):
        return (
            isinstance(other, Table)
            and other.parent == self.parent
            and other.name == self.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __repr__(self):
        return "<Table '{0}' in {1}>".format(self.name, self.parent)


class Tables(Collection):
    """A collection of all :meth:`table <Table>` objects on the specified sheet:

    >>> import xlwings as xw
    >>> xw.books['Book1'].sheets[0].tables
    Tables([<Table 'Table1' in <Sheet [Book11]Sheet1>>,
            <Table 'Table2' in <Sheet [Book11]Sheet1>>])

    .. versionadded:: 0.21.0
    """

    _wrap = Table

    def add(
        self,
        source=None,
        name=None,
        source_type=None,
        link_source=None,
        has_headers=True,
        destination=None,
        table_style_name="TableStyleMedium2",
    ):
        """
        Creates a Table to the specified sheet.

        Arguments
        ---------

        source : xlwings range, default None
            An xlwings range object, representing the data source.

        name : str, default None
            The name of the Table. By default, it uses the autogenerated name that is
            assigned by Excel.

        source_type : str, default None
            This currently defaults to ``xlSrcRange``, i.e. expects an xlwings range
            object. No other options are allowed at the moment.

        link_source : bool, default None
            Currently not implemented as this is only in case ``source_type`` is
            ``xlSrcExternal``.

        has_headers : bool or str, default True
            Indicates whether the data being imported has column labels. Defaults to
            ``True``. Possible values: ``True``, ``False``, ``'guess'``

        destination : xlwings range, default None
            Currently not implemented as this is used in case ``source_type`` is
            ``xlSrcExternal``.

        table_style_name : str, default 'TableStyleMedium2'
            Possible strings: ``'TableStyleLightN'`` (where N is 1-21),
            ``'TableStyleMediumN'`` (where N is 1-28),
            ``'TableStyleDarkN'`` (where N is 1-11)

        Returns
        -------
        Table

        Examples
        --------

        >>> import xlwings as xw
        >>> sheet = xw.Book().sheets[0]
        >>> sheet['A1'].value = [['a', 'b'], [1, 2]]
        >>> table = sheet.tables.add(source=sheet['A1'].expand(), name='MyTable')
        >>> table
        <Table 'MyTable' in <Sheet [Book1]Sheet1>>
        """

        impl = self.impl.add(
            source_type=source_type,
            source=source,
            link_source=link_source,
            has_headers=has_headers,
            destination=destination,
            table_style_name=table_style_name,
            name=name,
        )

        return Table(impl=impl)


class Chart:
    """
    The chart object is a member of the :meth:`charts <xlwings.main.Charts>` collection:

    >>> import xlwings as xw
    >>> sht = xw.books['Book1'].sheets[0]
    >>> sht.charts[0]  # or sht.charts['ChartName']
    <Chart 'Chart 1' in <Sheet [Book1]Sheet1>>
    """

    def __init__(self, name_or_index=None, impl=None):
        if impl is not None:
            self.impl = impl
        elif name_or_index is not None:
            self.impl = sheets.active.charts(name_or_index).impl
        else:
            self.impl = sheets.active.charts.add().impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    @property
    def name(self):
        """
        Returns or sets the name of the chart.
        """
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def parent(self):
        """
        Returns the parent of the chart.

        .. versionadded:: 0.9.0
        """
        # Chart sheet (parent is Book) is not supported
        return Sheet(impl=self.impl.parent)

    @property
    def chart_type(self):
        """
        Returns and sets the chart type of the chart.
        The following chart types are available:

        ``3d_area``,
        ``3d_area_stacked``,
        ``3d_area_stacked_100``,
        ``3d_bar_clustered``,
        ``3d_bar_stacked``,
        ``3d_bar_stacked_100``,
        ``3d_column``,
        ``3d_column_clustered``,
        ``3d_column_stacked``,
        ``3d_column_stacked_100``,
        ``3d_line``,
        ``3d_pie``,
        ``3d_pie_exploded``,
        ``area``,
        ``area_stacked``,
        ``area_stacked_100``,
        ``bar_clustered``,
        ``bar_of_pie``,
        ``bar_stacked``,
        ``bar_stacked_100``,
        ``bubble``,
        ``bubble_3d_effect``,
        ``column_clustered``,
        ``column_stacked``,
        ``column_stacked_100``,
        ``combination``,
        ``cone_bar_clustered``,
        ``cone_bar_stacked``,
        ``cone_bar_stacked_100``,
        ``cone_col``,
        ``cone_col_clustered``,
        ``cone_col_stacked``,
        ``cone_col_stacked_100``,
        ``cylinder_bar_clustered``,
        ``cylinder_bar_stacked``,
        ``cylinder_bar_stacked_100``,
        ``cylinder_col``,
        ``cylinder_col_clustered``,
        ``cylinder_col_stacked``,
        ``cylinder_col_stacked_100``,
        ``doughnut``,
        ``doughnut_exploded``,
        ``line``,
        ``line_markers``,
        ``line_markers_stacked``,
        ``line_markers_stacked_100``,
        ``line_stacked``,
        ``line_stacked_100``,
        ``pie``,
        ``pie_exploded``,
        ``pie_of_pie``,
        ``pyramid_bar_clustered``,
        ``pyramid_bar_stacked``,
        ``pyramid_bar_stacked_100``,
        ``pyramid_col``,
        ``pyramid_col_clustered``,
        ``pyramid_col_stacked``,
        ``pyramid_col_stacked_100``,
        ``radar``,
        ``radar_filled``,
        ``radar_markers``,
        ``stock_hlc``,
        ``stock_ohlc``,
        ``stock_vhlc``,
        ``stock_vohlc``,
        ``surface``,
        ``surface_top_view``,
        ``surface_top_view_wireframe``,
        ``surface_wireframe``,
        ``xy_scatter``,
        ``xy_scatter_lines``,
        ``xy_scatter_lines_no_markers``,
        ``xy_scatter_smooth``,
        ``xy_scatter_smooth_no_markers``

        .. versionadded:: 0.1.1
        """
        return self.impl.chart_type

    @chart_type.setter
    def chart_type(self, value):
        self.impl.chart_type = value

    def set_source_data(self, source):
        """
        Sets the source data range for the chart.

        Arguments
        ---------
        source : Range
            Range object, e.g. ``xw.books['Book1'].sheets[0].range('A1')``
        """
        self.impl.set_source_data(source.impl)

    @property
    def left(self):
        """
        Returns or sets the number of points that represent the horizontal position
        of the chart.
        """
        return self.impl.left

    @left.setter
    def left(self, value):
        self.impl.left = value

    @property
    def top(self):
        """
        Returns or sets the number of points that represent the vertical position
        of the chart.
        """
        return self.impl.top

    @top.setter
    def top(self, value):
        self.impl.top = value

    @property
    def width(self):
        """
        Returns or sets the number of points that represent the width of the chart.
        """
        return self.impl.width

    @width.setter
    def width(self, value):
        self.impl.width = value

    @property
    def height(self):
        """
        Returns or sets the number of points that represent the height of the chart.
        """
        return self.impl.height

    @height.setter
    def height(self, value):
        self.impl.height = value

    def delete(self):
        """
        Deletes the chart.
        """
        self.impl.delete()

    def to_png(self, path=None):
        """
        Exports the chart as PNG picture.

        Parameters
        ----------

        path : str or path-like, default None
            Path where you want to store the picture. Defaults to the name of the chart
            in the same directory as the Excel file if the Excel file is stored and to
            the current working directory otherwise.


        .. versionadded:: 0.24.8
        """
        path = utils.fspath(path)
        if path is None:
            directory, _ = os.path.split(self.parent.book.fullname)
            if directory:
                path = os.path.join(directory, self.name + ".png")
            else:
                path = str(Path.cwd() / self.name) + ".png"
        self.impl.to_png(path)

    def to_pdf(self, path=None, show=None, quality="standard"):
        """
        Exports the chart as PDF.

        Parameters
        ----------

        path : str or path-like, default None
            Path where you want to store the pdf. Defaults to the name of the chart in
            the same directory as the Excel file if the Excel file is stored and to the
            current working directory otherwise.

        show : bool, default False
            Once created, open the PDF file with the default application.

        quality : str, default ``'standard'``
            Quality of the PDF file. Can either be ``'standard'`` or ``'minimum'``.


        .. versionadded:: 0.26.2
        """
        return utils.to_pdf(self, path=path, show=show, quality=quality)

    def __repr__(self):
        return "<Chart '{0}' in {1}>".format(self.name, self.parent)


class Charts(Collection):
    """
    A collection of all :meth:`chart <Chart>` objects on the specified sheet:

    >>> import xlwings as xw
    >>> xw.books['Book1'].sheets[0].charts
    Charts([<Chart 'Chart 1' in <Sheet [Book1]Sheet1>>,
            <Chart 'Chart 1' in <Sheet [Book1]Sheet1>>])

    .. versionadded:: 0.9.0
    """

    _wrap = Chart

    def add(self, left=0, top=0, width=355, height=211):
        """
        Creates a new chart on the specified sheet.

        Arguments
        ---------

        left : float, default 0
            left position in points

        top : float, default 0
            top position in points

        width : float, default 355
            width in points

        height : float, default 211
            height in points

        Returns
        -------
        Chart

        Examples
        --------

        >>> import xlwings as xw
        >>> sht = xw.Book().sheets[0]
        >>> sht.range('A1').value = [['Foo1', 'Foo2'], [1, 2]]
        >>> chart = sht.charts.add()
        >>> chart.set_source_data(sht.range('A1').expand())
        >>> chart.chart_type = 'line'
        >>> chart.name
        'Chart1'
        """

        impl = self.impl.add(left, top, width, height)

        return Chart(impl=impl)


class Picture:
    """
    The picture object is a member of the :meth:`pictures <xlwings.main.Pictures>`
    collection:

    >>> import xlwings as xw
    >>> sht = xw.books['Book1'].sheets[0]
    >>> sht.pictures[0]  # or sht.charts['PictureName']
    <Picture 'Picture 1' in <Sheet [Book1]Sheet1>>

    .. versionchanged:: 0.9.0
    """

    def __init__(self, impl=None):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine
        being used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    @property
    def parent(self):
        """
        Returns the parent of the picture.

        .. versionadded:: 0.9.0
        """
        return Sheet(impl=self.impl.parent)

    @property
    def name(self):
        """
        Returns or sets the name of the picture.

        .. versionadded:: 0.5.0
        """
        return self.impl.name

    @name.setter
    def name(self, value):
        if value in self.parent.pictures:
            if value == self.name:
                return
            else:
                raise ShapeAlreadyExists(
                    f"'{value}' is already present on {self.parent.name}."
                )

        self.impl.name = value

    @property
    def left(self):
        """
        Returns or sets the number of points that represent the horizontal position
        of the picture.

        .. versionadded:: 0.5.0
        """
        return self.impl.left

    @left.setter
    def left(self, value):
        self.impl.left = value

    @property
    def top(self):
        """
        Returns or sets the number of points that represent the vertical position
        of the picture.

        .. versionadded:: 0.5.0
        """
        return self.impl.top

    @top.setter
    def top(self, value):
        self.impl.top = value

    @property
    def width(self):
        """
        Returns or sets the number of points that represent the width of the picture.

        .. versionadded:: 0.5.0
        """
        return self.impl.width

    @width.setter
    def width(self, value):
        self.impl.width = value

    @property
    def height(self):
        """
        Returns or sets the number of points that represent the height of the picture.

        .. versionadded:: 0.5.0
        """
        return self.impl.height

    @height.setter
    def height(self, value):
        self.impl.height = value

    def delete(self):
        """
        Deletes the picture.

        .. versionadded:: 0.5.0
        """
        self.impl.delete()

    def __eq__(self, other):
        return (
            isinstance(other, Picture)
            and other.parent == self.parent
            and other.name == self.name
        )

    def __ne__(self, other):
        return not self.__eq__(other)

    def __repr__(self):
        return "<Picture '{0}' in {1}>".format(self.name, self.parent)

    def update(self, image, format=None, export_options=None):
        """
        Replaces an existing picture with a new one, taking over the attributes of the
        existing picture.

        Arguments
        ---------

        image : str or path-like object or matplotlib.figure.Figure
            Either a filepath or a Matplotlib figure object.

        format : str, default None
            See under ``Pictures.add()``

        export_options : dict, default None
            See under ``Pictures.add()``


        .. versionadded:: 0.5.0
        """

        filename, is_temp_file = utils.process_image(
            image,
            format="png" if not format else format,
            export_options=export_options,
        )

        picture = Picture(impl=self.impl.update(filename))

        # Cleanup temp file
        if is_temp_file:
            try:
                os.unlink(filename)
            except:  # noqa: E722
                pass

        return picture

    @property
    def lock_aspect_ratio(self):
        """
        ``True`` will keep the original proportion,
        ``False`` will allow you to change height and width independently of each other
        (read/write).

        .. versionadded:: 0.24.0
        """
        return self.impl.lock_aspect_ratio

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value):
        self.impl.lock_aspect_ratio = value


class Pictures(Collection):
    """
    A collection of all :meth:`picture <Picture>` objects on the specified sheet:

    >>> import xlwings as xw
    >>> xw.books['Book1'].sheets[0].pictures
    Pictures([<Picture 'Picture 1' in <Sheet [Book1]Sheet1>>,
              <Picture 'Picture 2' in <Sheet [Book1]Sheet1>>])

    .. versionadded:: 0.9.0
    """

    _wrap = Picture

    @property
    def parent(self):
        return Sheet(impl=self.impl.parent)

    def add(
        self,
        image,
        link_to_file=False,
        save_with_document=True,
        left=None,
        top=None,
        width=None,
        height=None,
        name=None,
        update=False,
        scale=None,
        format=None,
        anchor=None,
        export_options=None,
    ):
        """
        Adds a picture to the specified sheet.

        Arguments
        ---------

        image : str or path-like object or matplotlib.figure.Figure
            Either a filepath or a Matplotlib figure object.

        left : float, default None
            Left position in points, defaults to 0. If you use ``top``/``left``, you
            must not provide a value for ``anchor``.

        top : float, default None
            Top position in points, defaults to 0. If you use ``top``/``left``,
            you must not provide a value for ``anchor``.

        width : float, default None
            Width in points. Defaults to original width.

        height : float, default None
            Height in points. Defaults to original height.

        name : str, default None
            Excel picture name. Defaults to Excel standard name if not provided,
            e.g., 'Picture 1'.

        update : bool, default False
            Replace an existing picture with the same name. Requires ``name`` to be set.

        scale : float, default None
            Scales your picture by the provided factor.

        format : str, default None
            Only used if image is a Matplotlib or Plotly plot. By default, the plot is
            inserted in the "png" format, but you may want to change this to a
            vector-based format like "svg" on Windows (may require Microsoft 365) or
            "eps" on macOS for better print quality. If you use ``'vector'``, it will be
            using ``'svg'`` on Windows and ``'eps'`` on macOS. To find out which formats
            your version of Excel supports, see:
            https://support.microsoft.com/en-us/topic/support-for-eps-images-has-been-turned-off-in-office-a069d664-4bcf-415e-a1b5-cbb0c334a840

        anchor: xw.Range, default None
            The xlwings Range object of where you want to insert the picture. If you use
            ``anchor``, you must not provide values for ``top``/``left``.

            .. versionadded:: 0.24.3

        export_options : dict, default None
            For Matplotlib plots, this dictionary is passed on to ``image.savefig()``
            with the following defaults: ``{"bbox_inches": "tight", "dpi": 200}``, so
            if you want to leave the picture uncropped and increase dpi to 300, use:
            ``export_options={"dpi": 300}``. For Plotly, the options are passed to
            ``write_image()``.

            .. versionadded:: 0.27.7

        Returns
        -------
        Picture

        Examples
        --------

        1. Picture

        >>> import xlwings as xw
        >>> sht = xw.Book().sheets[0]
        >>> sht.pictures.add(r'C:\\path\\to\\file.png')
        <Picture 'Picture 1' in <Sheet [Book1]Sheet1>>

        2. Matplotlib

        >>> import matplotlib.pyplot as plt
        >>> fig = plt.figure()
        >>> plt.plot([1, 2, 3, 4, 5])
        >>> sht.pictures.add(fig, name='MyPlot', update=True)
        <Picture 'MyPlot' in <Sheet [Book1]Sheet1>>
        """
        if anchor:
            if top or left:
                raise ValueError(
                    "You must either provide 'anchor' or 'top'/'left', but not both."
                )
        if update:
            if name is None:
                raise ValueError("If update is true then name must be specified")
            else:
                try:
                    pic = self[name]
                    return pic.update(
                        image, format=format, export_options=export_options
                    )
                except KeyError:
                    pass

        if name and name in self.parent.pictures:
            raise ShapeAlreadyExists(
                f"'{name}' is already present on {self.parent.name}."
            )

        filename, is_temp_file = utils.process_image(
            image,
            format="png" if not format else format,
            export_options=export_options,
        )

        if not (link_to_file or save_with_document):
            raise Exception(
                "Arguments link_to_file and save_with_document cannot both be false"
            )

        if (
            (height and width is None)
            or (width and height is None)
            or (width is None and height is None)
        ):
            # If only height or width are provided, it will be scaled after adding it
            # with the original dimensions
            im_width, im_height = -1, -1
        else:
            im_width, im_height = width, height

        picture = Picture(
            impl=self.impl.add(
                filename,
                link_to_file,
                save_with_document,
                left if left else None,
                top if top else None,
                width=im_width,
                height=im_height,
                anchor=anchor,
            )
        )

        if (height and width is None) or (width and height is None):
            # If only height or width are provided, lock aspect ratio so the picture
            # won't be distorted
            picture.lock_aspect_ratio = True
            if height:
                picture.height = height
            else:
                picture.width = width

        if scale:
            self.parent.shapes[picture.name].scale_width(
                factor=scale, relative_to_original_size=True
            )
            self.parent.shapes[picture.name].scale_height(
                factor=scale, relative_to_original_size=True
            )

        if name is not None:
            picture.name = name

        # Cleanup temp file
        if is_temp_file:
            try:
                os.unlink(filename)
            except:  # noqa: E722
                pass
        return picture


class Names:
    """
    A collection of all :meth:`name <Name>` objects in the workbook:

    >>> import xlwings as xw
    >>> book = xw.books['Book1']  # book scope and sheet scope
    >>> book.names
    [<Name 'MyName': =Sheet1!$A$3>]
    >>> book.sheets[0].names  # sheet scope only

    .. versionadded:: 0.9.0
    """

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine beingused.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def __call__(self, name_or_index):
        return Name(impl=self.impl(name_or_index))

    def contains(self, name_or_index):
        return self.impl.contains(name_or_index)

    def __len__(self):
        return len(self.impl)

    @property
    def count(self):
        """
        Returns the number of objects in the collection.
        """
        return len(self)

    def add(self, name, refers_to):
        """
        Defines a new name for a range of cells.

        Parameters
        ----------
        name : str
            Specifies the text to use as the name. Names cannot include spaces and
            cannot be formatted as cell references.

        refers_to : str
            Describes what the name refers to, in English, using A1-style notation.

        Returns
        -------
        Name


        .. versionadded:: 0.9.0
        """
        return Name(impl=self.impl.add(name, refers_to))

    def __getitem__(self, item):
        if isinstance(item, numbers.Number):
            return self(item + 1)
        else:
            return self(item)

    def __setitem__(self, key, value):
        if isinstance(value, Range):
            value.name = key
        elif key in self:
            self[key].refers_to = value
        else:
            self.add(key, value)

    def __contains__(self, item):
        if isinstance(item, numbers.Number):
            return 0 <= item < len(self)
        else:
            return self.contains(item)

    def __delitem__(self, key):
        if key in self:
            self[key].delete()
        else:
            raise KeyError(key)

    def __iter__(self):
        for i in range(len(self)):
            yield self(i + 1)

    def __repr__(self):
        r = []
        for i, n in enumerate(self):
            if i == 3:
                r.append("...")
                break
            else:
                r.append(repr(n))
        return "[" + ", ".join(r) + "]"


class Name:
    """
    The name object is a member of the :meth:`names <xlwings.main.Names>` collection:

    >>> import xlwings as xw
    >>> sht = xw.books['Book1'].sheets[0]
    >>> sht.names[0]  # or sht.names['MyName']
    <Name 'MyName': =Sheet1!$A$3>

    .. versionadded:: 0.9.0
    """

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.9.0
        """
        return self.impl.api

    def delete(self):
        """
        Deletes the name.

        .. versionadded:: 0.9.0
        """
        self.impl.delete()

    @property
    def name(self):
        """
        Returns or sets the name of the name object.

        .. versionadded:: 0.9.0
        """
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def refers_to(self):
        """
        Returns or sets the formula that the name is defined to refer to,
        in A1-style notation, beginning with an equal sign.

        .. versionadded:: 0.9.0
        """
        return self.impl.refers_to

    @refers_to.setter
    def refers_to(self, value):
        self.impl.refers_to = value

    @property
    def refers_to_range(self):
        """
        Returns the Range object referred to by a Name object.

        .. versionadded:: 0.9.0
        """
        return Range(impl=self.impl.refers_to_range)

    def __repr__(self):
        return "<Name '%s': %s>" % (self.name, self.refers_to)

    def __eq__(self, other):
        return (
            type(other) is Name
            and other.name == self.name
            and other.refers_to_range == self.refers_to_range
            and other.refers_to == self.refers_to
        )


def view(obj, sheet=None, table=True, chunksize=5000):
    """
    Opens a new workbook and displays an object on its first sheet by default. If you
    provide a sheet object, it will clear the sheet before displaying the object on the
    existing sheet.

    .. note::
      Only use this in an interactive context like e.g., a Jupyter notebook! Don't use
      this in a script as it depends on the active book.

    Parameters
    ----------
    obj : any type with built-in converter
        the object to display, e.g. numbers, strings, lists, numpy arrays, pandas
        DataFrames

    sheet : Sheet, default None
        Sheet object. If none provided, the first sheet of a new workbook is used.

    table : bool, default True
        If your object is a pandas DataFrame, by default it is formatted as an Excel
        Table

    chunksize : int, default 5000
        Chunks the loading of big arrays.

    Examples
    --------

    >>> import xlwings as xw
    >>> import pandas as pd
    >>> import numpy as np
    >>> df = pd.DataFrame(np.random.rand(10, 4), columns=['a', 'b', 'c', 'd'])
    >>> xw.view(df)

    See also: :meth:`load <xlwings.load>`

    .. versionchanged:: 0.22.0
    """
    if sheet is None:
        sheet = Book().sheets.active
    else:
        sheet.clear()

    app = sheet.book.app
    app.activate(steal_focus=True)

    with app.properties(screen_updating=False):
        if pd and isinstance(obj, pd.DataFrame):
            if table:
                sheet["A1"].options(
                    assign_empty_index_names=True, chunksize=chunksize
                ).value = obj
                sheet.tables.add(sheet["A1"].expand())
            else:
                sheet["A1"].options(
                    assign_empty_index_names=False, chunksize=chunksize
                ).value = obj
        else:
            sheet["A1"].value = obj
        sheet.autofit()


def load(index=1, header=1, chunksize=5000):
    """
    Loads the selected cell(s) of the active workbook into a pandas DataFrame. If you
    select a single cell that has adjacent cells, the range is auto-expanded (via
    current region) and turned into a pandas DataFrame. If you don't have pandas
    installed, it returns the values as nested lists.

    .. note::
      Only use this in an interactive context like e.g. a Jupyter notebook! Don't use
      this in a script as it depends on the active book.

    Parameters
    ----------
    index : bool or int, default 1
        Defines the number of columns on the left that will be turned into the
        DataFrame's index

    header : bool or int, default 1
        Defines the number of rows at the top that will be turned into the DataFrame's
        columns

    chunksize : int, default 5000
        Chunks the loading of big arrays.

    Examples
    --------
    >>> import xlwings as xw
    >>> xw.load()

    See also: :meth:`view <xlwings.view>`

    .. versionchanged:: 0.23.1
    """
    selection = books.active.selection
    if selection.shape == (1, 1):
        selection = selection.current_region
    if pd:
        values = selection.options(
            pd.DataFrame, index=index, header=header, chunksize=chunksize
        ).value
    else:
        values = selection.options(chunksize=chunksize).value
    return values


class Macro:
    def __init__(self, app, macro):
        self.app = app
        self.macro = macro

    def run(self, *args):
        args = [
            i.api
            if isinstance(i, (App, Book, Sheet, Range, Shape, Chart, Picture, Name))
            else i
            for i in args
        ]
        return self.app.impl.run(self.macro, args)

    __call__ = run


class Characters:
    """
    The characters object can be accessed as an attribute of the range or shape object.

    * ``mysheet['A1'].characters``
    * ``mysheet.shapes[0].characters``

    .. note:: On macOS, ``characters`` are currently not supported due to bugs/lack of
              support in AppleScript.

    .. versionadded:: 0.23.0
    """

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj) of the engine
        being used.

        .. versionadded:: 0.23.0
        """
        return self.impl.api

    @property
    def text(self):
        """
        Returns or sets the text property of a ``characters`` object.

        >>> sheet['A1'].value = 'Python'
        >>> sheet['A1'].characters[:3].text
        Pyt

        .. versionadded:: 0.23.0
        """
        return self.impl.text

    @property
    def font(self):
        """
        Returns or sets the text property of a ``characters`` object.

        >>> sheet['A1'].characters[1:3].font.bold = True
        >>> sheet['A1'].characters[1:3].font.bold
        True

        .. versionadded:: 0.23.0
        """
        return Font(self.impl.font)

    def __getitem__(self, item):
        if (
            isinstance(item, slice)
            and (item.start and item.stop)
            and (item.start == item.stop)
        ):
            raise ValueError(
                self.__class__.__name__ + " object does not support empty slices"
            )
        if isinstance(item, slice) and item.step is not None:
            raise ValueError(
                self.__class__.__name__
                + " object does not support slicing with non-default steps"
            )
        if isinstance(item, slice):
            return Characters(self.impl[item.start : item.stop])
        else:
            return Characters(self.impl[item])


class Font:
    """
    The font object can be accessed as an attribute of the range or shape object.

    * ``mysheet['A1'].font``
    * ``mysheet.shapes[0].font``

    .. versionadded:: 0.23.0
    """

    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.

        .. versionadded:: 0.23.0
        """
        return self.impl.api

    @property
    def bold(self):
        """
        Returns or sets the bold property (boolean).

        >>> sheet['A1'].font.bold = True
        >>> sheet['A1'].font.bold
        True

        .. versionadded:: 0.23.0
        """
        return self.impl.bold

    @bold.setter
    def bold(self, value):
        self.impl.bold = value

    @property
    def italic(self):
        """
        Returns or sets the italic property (boolean).

        >>> sheet['A1'].font.italic = True
        >>> sheet['A1'].font.italic
        True

        .. versionadded:: 0.23.0
        """
        return self.impl.italic

    @italic.setter
    def italic(self, value):
        self.impl.italic = value

    @property
    def size(self):
        """
        Returns or sets the size (float).

        >>> sheet['A1'].font.size = 13
        >>> sheet['A1'].font.size
        13

        .. versionadded:: 0.23.0
        """
        return self.impl.size

    @size.setter
    def size(self, value):
        self.impl.size = value

    @property
    def color(self):
        """
        Returns or sets the color property (tuple).

        >>> sheet['A1'].font.color = (255, 0, 0)  # or '#ff0000'
        >>> sheet['A1'].font.color
        (255, 0, 0)

        .. versionadded:: 0.23.0
        """
        return self.impl.color

    @color.setter
    def color(self, value):
        if isinstance(value, str):
            value = utils.hex_to_rgb(value)
        self.impl.color = value

    @property
    def name(self):
        """
        Returns or sets the name of the font (str).

        >>> sheet['A1'].font.name = 'Calibri'
        >>> sheet['A1'].font.name
        Calibri

        .. versionadded:: 0.23.0
        """
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value


class Books(Collection):
    """
    A collection of all :meth:`book <Book>` objects:

    >>> import xlwings as xw
    >>> xw.books  # active app
    Books([<Book [Book1]>, <Book [Book2]>])
    >>> xw.apps[10559].books  # specific app, get the PIDs via xw.apps.keys()
    Books([<Book [Book1]>, <Book [Book2]>])

    .. versionadded:: 0.9.0
    """

    _wrap = Book

    @property
    def active(self):
        """
        Returns the active Book.
        """
        return Book(impl=self.impl.active)

    def add(self):
        """
        Creates a new Book. The new Book becomes the active Book. Returns a Book object.
        """
        return Book(impl=self.impl.add())

    def open(
        self,
        fullname=None,
        update_links=None,
        read_only=None,
        format=None,
        password=None,
        write_res_password=None,
        ignore_read_only_recommended=None,
        origin=None,
        delimiter=None,
        editable=None,
        notify=None,
        converter=None,
        add_to_mru=None,
        local=None,
        corrupt_load=None,
        json=None,
    ):
        """
        Opens a Book if it is not open yet and returns it. If it is already open,
        it doesn't raise an exception but simply returns the Book object.

        Parameters
        ----------
        fullname : str or path-like object
            filename or fully qualified filename, e.g. ``r'C:\\path\\to\\file.xlsx'``
            or ``'file.xlsm'``. Without a full path, it looks for the file in the
            current working directory.

        Other Parameters
            see: :meth:`xlwings.Book()`

        Returns
        -------
        Book : Book that has been opened.

        """
        if self.impl.app.engine.type == "remote":
            return Book(impl=self.impl.open(json=json))

        fullname = utils.fspath(fullname)
        if not os.path.exists(fullname):
            raise FileNotFoundError("No such file: '%s'" % fullname)
        fullname = os.path.realpath(fullname)
        _, name = os.path.split(fullname)

        if self.impl.app.engine.type == "reader":
            return Book(impl=self.impl.open(filename=fullname))

        try:
            impl = self.impl(name)
            if not os.path.samefile(impl.fullname, fullname):
                raise ValueError(
                    "Cannot open two workbooks named '%s', even if they are saved in"
                    "different locations." % name
                )
        except KeyError:
            impl = self.impl.open(
                fullname,
                update_links,
                read_only,
                format,
                password,
                write_res_password,
                ignore_read_only_recommended,
                origin,
                delimiter,
                editable,
                notify,
                converter,
                add_to_mru,
                local,
                corrupt_load,
            )
        return Book(impl=impl)


class Sheets(Collection):
    """
    A collection of all :meth:`sheet <Sheet>` objects:

    >>> import xlwings as xw
    >>> xw.sheets  # active book
    Sheets([<Sheet [Book1]Sheet1>, <Sheet [Book1]Sheet2>])
    >>> xw.Book('Book1').sheets  # specific book
    Sheets([<Sheet [Book1]Sheet1>, <Sheet [Book1]Sheet2>])

    .. versionadded:: 0.9.0
    """

    _wrap = Sheet

    @property
    def active(self):
        """
        Returns the active Sheet.
        """
        return Sheet(impl=self.impl.active)

    def __call__(self, name_or_index):
        if isinstance(name_or_index, Sheet):
            return name_or_index
        else:
            return Sheet(impl=self.impl(name_or_index))

    def __delitem__(self, name_or_index):
        self[name_or_index].delete()

    def add(self, name=None, before=None, after=None):
        """
        Creates a new Sheet and makes it the active sheet.

        Parameters
        ----------
        name : str, default None
            Name of the new sheet. If None, will default to Excel's default name.
        before : Sheet, default None
            An object that specifies the sheet before which the new sheet is added.
        after : Sheet, default None
            An object that specifies the sheet after which the new sheet is added.

        Returns
        -------

        sheet : Sheet
            Added sheet object

        """
        if name is not None:
            if name.lower() in (s.name.lower() for s in self):
                raise ValueError("Sheet named '%s' already present in workbook" % name)
        if before is not None and not isinstance(before, Sheet):
            before = self(before)
        if after is not None and not isinstance(after, Sheet):
            after = self(after)
        impl = self.impl.add(before and before.impl, after and after.impl, name)
        return Sheet(impl=impl)


class ActiveEngineApps(Apps):
    def __init__(self):
        pass

    _name = "Apps"

    @property
    def impl(self):
        if engines.active is None:
            if not (
                sys.platform.startswith("darwin") or sys.platform.startswith("win")
            ):
                raise XlwingsError(
                    "The interactive mode of xlwings is only supported on Windows and "
                    "macOS. On Linux, you can use xlwings Server or xlwings Reader."
                )
            elif sys.platform.startswith("darwin"):
                raise XlwingsError(
                    'Make sure to have "appscript" and "psutil", '
                    "dependencies of xlwings, installed."
                )
            elif sys.platform.startswith("win"):
                raise XlwingsError(
                    'Make sure to have "pywin32", a dependency of xlwings, installed.'
                )
        return engines.active.apps.impl


class ActiveAppBooks(Books):
    def __init__(self):
        pass

    # override class name which appears in repr
    _name = "Books"

    @property
    def impl(self):
        if not apps:
            raise XlwingsError("Couldn't find any active App!")
        return apps.active.books.impl


class ActiveBookSheets(Sheets):
    def __init__(self):
        pass

    # override class name which appears in repr
    _name = "Sheets"

    @property
    def impl(self):
        return books.active.sheets.impl


apps = ActiveEngineApps()

books = ActiveAppBooks()

sheets = ActiveBookSheets()
