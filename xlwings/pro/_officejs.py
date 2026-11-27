"""Office.js access from Python in xlwings Lite: `book.officejs()`.

Proxies record property gets, method calls and property sets as plain-data ops.
`await ctx.sync()` ships them as JSON text to `js.xlwings.officejs`, where a
generic executor replays them against Office.js inside one held `Excel.run`
per session, and returns the values requested with `load()`.

While a session is open, `book.api`, `sheet.api` and `range.api` return proxies
bound to it, and the curated API doesn't talk to Excel (see
`_xlremote.check_no_officejs_session`). If the session wrote anything, the
book's snapshot is reloaded when it ends.
"""

import itertools
import json
import math
import numbers

from .. import XlwingsError

_session_ids = itertools.count(1)


class OfficeJsError(XlwingsError):
    """Raised when Office.js work in a {meth}`Book.officejs <xlwings.Book.officejs>` session fails (xlwings Lite only).

    This happens when Excel rejects what `await ctx.sync()` or the end of the session sent, when recorded code can't run (e.g. a method that doesn't exist), and when an Office.js object is used as a value before the property is loaded and synced. After a failed sync, the session can't sync again and discards its unsynced work when it ends. What was synced before stays applied, and the failed sync may have applied part of its changes.

    Attributes:
        code: Office.js error code such as `"ItemNotFound"`, if there is one.
        name: Name of the JavaScript error, such as `"RichApi.Error"`.
        op: The recorded expression that failed, such as `context.workbook.worksheets.getItem("Sheet1").getRange("A1").nope()`, or `book.sheets["Sheet1"]["$A$1"].api.nope()` for an object from `.api`. `None` for errors that Excel reports for a whole sync.
        debug_info: The `debugInfo` of an Office.js error as a dictionary, e.g. with the `errorLocation`.

    ```{warning}
    Experimental: the API may change.
    ```
    """

    def __init__(self, message, *, code=None, name=None, op=None, debug_info=None):
        super().__init__(message)
        self.code = code
        self.name = name
        self.op = op
        self.debug_info = debug_info


class OfficeJsWriteWithheld(OfficeJsError):
    """Raised when a Wingman dry run in xlwings Lite reaches an Office.js write.

    A dry run withholds workbook writes. It can't continue past an Office.js write, because the code after it may depend on objects that the write would have created, so the preview ends here. Nothing from the withheld sync is written, and the session discards its unsynced work when it ends.

    ```{warning}
    Experimental: the API may change.
    ```
    """


def _encode(value):
    """Turn an argument or set value into JSON data, proxies into handles."""
    if isinstance(value, Proxy):
        return {"$h": value._emit()}
    if type(value).__module__ == "numpy":
        kind = getattr(getattr(value, "dtype", None), "kind", None)
        if kind == "b":
            # `numpy.bool_` isn't a `bool`
            value = bool(value)
        elif kind in ("m", "M"):
            # `timedelta64` is a `numbers.Integral`: don't send it as a number
            raise _unsupported(value)
    if value is None or isinstance(value, (bool, str)):
        return value
    if isinstance(value, numbers.Integral):
        return int(value)
    if isinstance(value, numbers.Real):
        value = float(value)
        if not math.isfinite(value):
            raise ValueError(
                f"Can't send {value} to Office.js (JSON has no NaN or infinity). "
                'For an empty cell, use "" instead.'
            )
        return value
    if isinstance(value, (list, tuple)):
        return [_encode(v) for v in value]
    if isinstance(value, dict):
        if not all(isinstance(k, str) for k in value):
            raise TypeError("Dictionary keys sent to Office.js must be strings")
        return {k: _encode(v) for k, v in value.items()}
    raise _unsupported(value)


def _unsupported(value):
    cls = type(value)
    name = (
        cls.__name__
        if cls.__module__ == "builtins"
        else f"{cls.__module__}.{cls.__name__}"
    )
    return TypeError(
        f"Can't send {name} to Office.js. Use str, int, float, bool, None, lists "
        "and dicts of those, or Office.js objects. Convert NumPy arrays and "
        "DataFrame values with .tolist(), and dates to strings or numbers."
    )


class OfficeJsSession:
    """Asynchronous context manager returned by {meth}`Book.officejs <xlwings.Book.officejs>`.

    The public members are `context` (the Office.js `RequestContext`), `workbook` (shortcut for `context.workbook`) and `sync()`.
    """

    def __init__(self, book, send=None):
        # `book` is the remote Book impl
        self._book = book
        self._send = send
        self._session_id = f"s{next(_session_ids)}"
        self._request_ids = itertools.count(1)
        self._op_ids = itertools.count(1)
        self._ops = []
        self._returns = {}
        self._proxies = {0: None}
        self._labels = {}
        # `.api` proxies, so repeated accesses return the same object
        self._api_proxies = {}
        self._state = "new"
        # Set once a batch with write ops may have reached Excel
        self.wrote = False
        # Set when a blocked curated call may have changed the local snapshot
        self.snapshot_changed = False
        self.discard_error = None
        self.context = Proxy(self, None, "context", kind="root", handle=0)
        self._proxies[0] = self.context

    @property
    def workbook(self):
        return self.context.workbook

    async def __aenter__(self):
        book = self._book
        if getattr(book, "_officejs_session", None) is not None:
            raise XlwingsError("An Office.js session is already open for this book.")
        # Curated work before the session is applied first.
        if book._json["actions"] or book._dispatch is not None:
            await book.flush()
        await self._request("officejs.open")
        self._state = "open"
        book._officejs_session = self
        return self

    async def __aexit__(self, exc_type, exc, tb):
        discard = exc_type is not None or self._state == "failed"
        ops = self._ops
        self._ops, self._returns = [], {}
        self._state = "closed"
        self._book._officejs_session = None
        try:
            if not discard:
                await self._request("officejs.close", ops=ops)
            else:
                try:
                    response = await self._request("officejs.close", discard=True)
                    self.discard_error = response.get("discardError")
                except Exception:
                    # Don't mask the exception that ended the session.
                    if exc_type is None:
                        raise
        except BaseException:
            await self._reload(quiet=True)
            raise
        await self._reload(quiet=exc_type is not None)
        return False

    async def _reload(self, quiet):
        """Refresh the book's snapshot if the session may have changed Excel, or
        a blocked curated call may have changed the snapshot itself.

        Also after a failed or discarded session: earlier syncs may have
        committed, and the script may catch the exception and go on.
        """
        if not (self.wrote or self.snapshot_changed):
            return
        try:
            await self._book.reload_after_officejs()
        except Exception:
            if not quiet:
                raise

    async def sync(self):
        """Run the recorded ops in Excel and fetch the values requested with `load()`."""
        self._check_open()
        if self._state == "failed":
            raise OfficeJsError(
                "This Office.js session failed earlier and can't sync anymore.",
                code="SessionFailed",
            )
        ops, returns = self._ops, self._returns
        self._ops, self._returns = [], {}
        try:
            response = await self._request(
                "officejs.sync", ops=ops, **{"return": list(returns)}
            )
        except OfficeJsError:
            self._state = "failed"
            raise
        for key, data in response.get("results", {}).items():
            handle = int(key) if key.lstrip("-").isdigit() else key
            proxy = self._proxies.get(handle)
            if proxy is not None:
                proxy._bind(data)

    def _check_open(self):
        if self._state == "closed":
            raise XlwingsError(
                "This Office.js object belongs to a closed session. After the "
                "'async with book.officejs()' block, you can only read properties "
                "that were loaded and synced inside it."
            )
        if self._state == "new":
            raise XlwingsError(
                "Use the session with 'async with book.officejs() as ctx'."
            )

    # `.api` of xlwings objects

    def _api_proxy(self, key, make, label):
        proxy = self._api_proxies.get(key)
        if proxy is None:
            proxy = self._api_proxies[key] = make()
            # Messages name the xlwings expression rather than the sheet ID
            object.__setattr__(proxy, "_label", label)
            self._labels[proxy._handle] = label
        return proxy

    def book_api(self):
        return self.workbook

    def sheet_api(self, sheet):
        """`worksheets.getItem()` by sheet ID, so a rename doesn't break it."""
        key = sheet.api.get("id") or sheet.name
        return self._api_proxy(
            ("sheet", key),
            lambda: self.workbook.worksheets.getItem(key),
            f"book.sheets[{json.dumps(sheet.name)}].api",
        )

    def range_api(self, rng):
        address = rng.address
        if address is None:
            raise XlwingsError("This range has no address (no cells are selected).")
        sheet = rng.sheet
        key = sheet.api.get("id") or sheet.name
        return self._api_proxy(
            ("range", key, address),
            lambda: self.sheet_api(sheet).getRange(address),
            f"book.sheets[{json.dumps(sheet.name)}][{json.dumps(address)}].api",
        )

    def _add_op(self, kind, target, name, label, **fields):
        op_id = next(self._op_ids)
        self._ops.append(
            {"id": op_id, "kind": kind, "target": target, "name": name, **fields}
        )
        self._labels[op_id] = label
        return op_id

    async def _request(self, type_, **fields):
        request = {
            "type": type_,
            "requestId": next(self._request_ids),
            "executionId": None,
            "sessionId": self._session_id,
            **fields,
        }
        text = json.dumps(request, allow_nan=False)
        send = self._send
        if send is None:
            import js

            send = js.xlwings.officejs
        response = json.loads(await send(text))
        error = response.get("error") or {}
        if response.get("wrote") or error.get("wrote"):
            self.wrote = True
        if not response.get("ok"):
            raise self._error(error)
        return response

    def _error(self, error):
        op_id = error.get("opId")
        op = self._labels.get(op_id) if op_id is not None else None
        code = error.get("code")
        message = error.get("message") or "Office.js request failed"
        prefix = f"{code}: " if code else ""
        suffix = f" (at {op})" if op else ""
        cls = OfficeJsWriteWithheld if code == "withheld" else OfficeJsError
        return cls(
            f"{prefix}{message}{suffix}",
            code=code,
            name=error.get("name"),
            op=op,
            debug_info=error.get("debugInfo"),
        )


_PROXY_SLOTS = (
    "_session",
    "_parent",
    "_name",
    "_kind",
    "_handle",
    "_label",
    "_children",
    "_data",
    "_items",
    "_plain",
)


class Proxy:
    """An Office.js object or property recorded in a session.

    Attribute access, calls and sets record ops; values come from the data
    loaded with `load()` at the last `await ctx.sync()`.
    """

    __slots__ = _PROXY_SLOTS

    def __init__(self, session, parent, name, *, kind, handle=None, label=None):
        set_ = object.__setattr__
        set_(self, "_session", session)
        set_(self, "_parent", parent)
        set_(self, "_name", name)
        set_(self, "_kind", kind)
        set_(self, "_handle", handle)
        set_(self, "_label", label or name)
        set_(self, "_children", {})
        set_(self, "_data", None)
        set_(self, "_items", None)
        # Paths in `_data` of plain object values (not navigation properties)
        set_(self, "_plain", frozenset())

    # Recording

    def _emit(self):
        """The handle of this object, emitting its `get` op on first use."""
        self._session._check_open()
        if self._handle is None:
            handle = self._session._add_op(
                "get", self._parent._emit(), self._name, self._label
            )
            object.__setattr__(self, "_handle", handle)
            self._session._proxies[handle] = self
        return self._handle

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        if name == "items" and self._items is not None:
            return list(self._items)
        found, value, plain = self._loaded(name)
        if found and (plain or not isinstance(value, dict)):
            return value
        if not found:
            # Loaded data stays readable after the session; anything else
            # would need a new op.
            self._session._check_open()
        # Unloaded or navigation property: the same proxy every time
        child = self._children.get(name)
        if child is None:
            child = Proxy(
                self._session, self, name, kind="get", label=f"{self._label}.{name}"
            )
            self._children[name] = child
        return child

    def __setattr__(self, name, value):
        if name in _PROXY_SLOTS:
            object.__setattr__(self, name, value)
            return
        encoded = _encode(value)
        target = self._emit()
        self._session._add_op(
            "set", target, name, f"{self._label}.{name}", value=encoded
        )

    def __call__(self, *args, **kwargs):
        if self._kind != "get":
            raise TypeError(f"{self._label} isn't callable")
        if kwargs:
            raise TypeError(
                "Office.js methods take positional arguments only; pass options "
                "as a dict"
            )
        parent = self._parent
        encoded = [_encode(arg) for arg in args]
        target = parent._emit()
        label = f"{parent._label}.{self._name}({_short_args(encoded)})"
        handle = self._session._add_op("call", target, self._name, label, args=encoded)
        if self._name == "load":
            # Office.js returns the object itself; its data comes back after sync
            self._session._returns[parent._handle] = None
            return parent
        result = Proxy(
            self._session, parent, self._name, kind="call", handle=handle, label=label
        )
        self._session._proxies[handle] = result
        return result

    # Loaded data

    def _loaded(self, name):
        """Find `name` in the data of this object or its nearest loaded ancestor.

        Returns `(found, value, plain)`; `plain` is True for an object value
        that the executor marked as plain data rather than a navigation
        property (e.g. `ConditionalCellValueRule`).
        """
        path = [name]
        proxy = self
        while proxy is not None:
            if proxy._data is not None:
                data = proxy._data
                for key in path:
                    if isinstance(data, dict) and key in data:
                        data = data[key]
                    else:
                        return False, None, False
                return True, data, "/".join(path) in proxy._plain
            if proxy._kind != "get":
                break
            path.insert(0, proxy._name)
            proxy = proxy._parent
        return False, None, False

    def _bind(self, data):
        if isinstance(data, dict):
            object.__setattr__(self, "_plain", frozenset(data.pop("$plain", ())))
        if isinstance(data, dict) and "$items" in data:
            handles = data.pop("$items")
            items_data = data.get("items") or []
            items = []
            for index, handle in enumerate(handles):
                item = Proxy(
                    self._session,
                    self,
                    "items",
                    kind="item",
                    handle=handle,
                    label=f"{self._label}.items[{index}]",
                )
                item._bind(items_data[index] if index < len(items_data) else {})
                self._session._proxies[handle] = item
                items.append(item)
            object.__setattr__(self, "_items", items)
        object.__setattr__(self, "_data", data)

    # Proxies have no value of their own

    def _is_loaded_object(self):
        """Whether loaded data shows this property to be an object, e.g. after
        `workbook.load("worksheets")`."""
        found, value, plain = self._parent._loaded(self._name)
        return found and isinstance(value, dict) and not plain

    def _no_value(self, *args, **kwargs):
        if self._kind == "get" and self._name == "isNullObject":
            # Office.js sets it on `*OrNullObject` results at sync, no load()
            raise OfficeJsError(f"{self._label} is available after 'await ctx.sync()'.")
        parent = self._parent
        if self._kind == "get" and not self._is_loaded_object():
            if self._name == "value" and parent._kind == "call":
                # A ClientResult (e.g. `getCount()`) has no load()
                raise OfficeJsError(
                    f"{self._label} isn't available yet. The result of a method "
                    "is available after 'await ctx.sync()'; a property also needs "
                    f'{parent._label}.load("value") before that.'
                )
            raise OfficeJsError(
                f"{self._label} isn't loaded; call "
                f'{parent._label}.load("{self._name}") and '
                "'await ctx.sync()' first."
            )
        # A method's result, a collection item or a loaded navigation property
        raise OfficeJsError(
            f"{self._label} is an Office.js object, not a value. Load the "
            "properties you need and 'await ctx.sync()', then read them: .items "
            "for a collection, .value for the result of a method such as "
            "getCount() (which needs no load)."
        )

    __str__ = __bool__ = __len__ = __iter__ = __contains__ = __getitem__ = _no_value
    __int__ = __float__ = __index__ = _no_value
    __eq__ = __ne__ = __lt__ = __le__ = __gt__ = __ge__ = _no_value
    __add__ = __radd__ = __sub__ = __rsub__ = __mul__ = __rmul__ = _no_value
    __truediv__ = __rtruediv__ = _no_value
    __hash__ = object.__hash__

    def __repr__(self):
        unloaded = self._kind == "get" and not self._is_loaded_object()
        state = ", not loaded" if unloaded else ""
        return f"<Proxy {self._label}{state}>"


def _short_args(args):
    text = ", ".join(json.dumps(arg) for arg in args)
    return text if len(text) <= 60 else text[:57] + "..."
