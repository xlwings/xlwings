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

from .. import OfficeJsError, OfficeJsWriteWithheld, XlwingsError

_session_ids = itertools.count(1)


def _encode(value):
    """Turn an argument or set value into JSON data, proxies into handles."""
    if isinstance(value, Proxy):
        return {"$h": value._emit()}
    if type(value).__module__ == "numpy" and getattr(value, "ndim", None) == 0:
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

    `sync()` runs the recorded ops, `context` is the Office.js `RequestContext`, and any other Office.js member, e.g. `workbook`, `application` or `load()`, is forwarded to `context`.
    """

    def __init__(self, book, send=None):
        # `book` is the remote Book impl
        self._book = book
        self._send = send
        self._session_id = f"s{next(_session_ids)}"
        self._request_ids = itertools.count(1)
        self._op_ids = itertools.count(1)
        # Numbers each sync, so the freshest loaded data wins (see `Proxy._find`)
        self._sync_ids = itertools.count(1)
        self._ops = []
        self._returns = {}
        self._proxies = {0: None}
        self._labels = {}
        # `.api` proxies, so repeated accesses return the same object
        self._api_proxies = {}
        self._state = "new"
        # Set once a batch with write ops may have reached Excel
        self._wrote = False
        # Set when a blocked curated call may have changed the local snapshot
        self._snapshot_changed = False
        self._discard_error = None
        self.context = Proxy(self, None, "context", kind="root", handle=0)
        self._proxies[0] = self.context

    def __getattr__(self, name):
        # Only called for names the session doesn't have: forward Office.js
        # members (`ctx.workbook`, `ctx.application`, `ctx.load(...)`) to the
        # RequestContext, so Office.js code translates with `context` -> `ctx`.
        context = self.__dict__.get("context")
        if name.startswith("_") or context is None:
            raise AttributeError(name)
        return getattr(context, name)

    async def __aenter__(self):
        book = self._book
        if getattr(book, "_officejs_session", None) is not None:
            raise XlwingsError("An Office.js session is already open for this book.")
        # Claimed before the awaits, so a concurrent task can't open a second one
        book._officejs_session = self
        try:
            # Curated work before the session is applied first.
            if book._json["actions"] or book._dispatch is not None:
                await book.flush()
            await self._request("officejs.open")
        except BaseException:
            book._officejs_session = None
            raise
        self._state = "open"
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
                    self._discard_error = response.get("discardError")
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
        if not (self._wrote or self._snapshot_changed):
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
        sync_id = next(self._sync_ids)
        for key, data in response.get("results", {}).items():
            handle = int(key) if key.lstrip("-").isdigit() else key
            proxy = self._proxies.get(handle)
            if proxy is not None:
                proxy._bind(data, sync_id)

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
            self._wrote = True
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
    "_sync_id",
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
        # Item proxies of the collections in `_data`, by path ("" for this object)
        set_(self, "_items", {})
        # Paths in `_data` of plain object values (not navigation properties)
        set_(self, "_plain", frozenset())
        # The sync that returned `_data`
        set_(self, "_sync_id", 0)

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
        found, value, plain, owner, path = self._find(name)
        if found and name == "items":
            items = owner._items.get("/".join(path[:-1]))
            if items is not None:
                return list(items)
        # A collection loaded through its parent is a list in its data, but it
        # stays a proxy, with `items`
        collection = (
            found and isinstance(value, list) and "/".join(path) in owner._items
        )
        if found and not collection and (plain or not isinstance(value, dict)):
            return value
        # Unloaded or navigation property: the same proxy every time
        child = self._children.get(name)
        if child is None:
            if not found:
                # Loaded data stays readable after the session, also through
                # proxies from inside it (e.g. `ctx.workbook`); anything else
                # would need a new op.
                self._session._check_open()
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
            # Office.js returns the loaded object; its data comes back after sync.
            # `context.load(obj, props)` loads `obj`.
            loaded = parent
            if parent._kind == "root" and args and isinstance(args[0], Proxy):
                loaded = args[0]
            self._session._returns[loaded._handle] = None
            return loaded
        result = Proxy(
            self._session, parent, self._name, kind="call", handle=handle, label=label
        )
        self._session._proxies[handle] = result
        return result

    # Loaded data

    def _loaded(self, name):
        """`(found, value, plain)` of `name` in loaded data, see `_find`."""
        return self._find(name)[:3]

    def _find(self, name):
        """Find `name` in the data of this object or its ancestors.

        Returns `(found, value, plain, owner, path)`. Of several loads that
        contain `name`, the one from the latest sync wins, e.g. a later
        `rng.load("format/fill/color")` over an earlier `fill.load("color")`.
        `plain` is True for an object value that the executor marked as plain
        data rather than a navigation property (e.g. `ConditionalCellValueRule`);
        `owner` is the proxy whose data holds the value at `path`.
        """
        best = (False, None, False, None, None)
        path = [name]
        proxy = self
        while proxy is not None:
            if proxy._data is not None and (
                best[3] is None or proxy._sync_id > best[3]._sync_id
            ):
                data = proxy._data
                for key in path:
                    if isinstance(data, list) and key == "items":
                        # `items` of a collection that is a list in its parent's data
                        continue
                    if not (isinstance(data, dict) and key in data):
                        break
                    data = data[key]
                else:
                    plain = "/".join(path) in proxy._plain
                    best = (True, data, plain, proxy, list(path))
            if proxy._kind != "get":
                break
            path.insert(0, proxy._name)
            proxy = proxy._parent
        return best

    def _bind(self, data, sync_id):
        items = {}
        if isinstance(data, dict):
            object.__setattr__(self, "_plain", frozenset(data.pop("$plain", ())))
            # Item handles of this collection ("$items") and of collections
            # loaded through it ("$itemsAt", by path)
            handles_at = data.pop("$itemsAt", None) or {}
            if "$items" in data:
                handles_at = {"": data.pop("$items"), **handles_at}
            for path, handles in handles_at.items():
                keys = path.split("/") if path else []
                container = data
                for key in keys:
                    container = (
                        container.get(key) if isinstance(container, dict) else None
                    )
                # A collection loaded through its parent is a list in the
                # parent's data, e.g. {"worksheets": [{"name": "Sheet1"}]}
                items_data = (
                    container.get("items") if isinstance(container, dict) else container
                )
                items_data = items_data if isinstance(items_data, list) else []
                label = "".join(f".{key}" for key in [self._label, *keys])[1:]
                items[path] = []
                for index, handle in enumerate(handles):
                    item = Proxy(
                        self._session,
                        self,
                        "items",
                        kind="item",
                        handle=handle,
                        label=f"{label}.items[{index}]",
                    )
                    item_data = items_data[index] if index < len(items_data) else {}
                    item._bind(item_data, sync_id)
                    self._session._proxies[handle] = item
                    items[path].append(item)
        object.__setattr__(self, "_items", items)
        object.__setattr__(self, "_sync_id", sync_id)
        object.__setattr__(self, "_data", data)

    # Proxies have no value of their own

    def _is_loaded_object(self):
        """Whether loaded data shows this property to be an object, e.g. after
        `workbook.load("worksheets")`."""
        found, value, plain, owner, path = self._parent._find(self._name)
        if found and isinstance(value, list):
            return "/".join(path) in owner._items
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
