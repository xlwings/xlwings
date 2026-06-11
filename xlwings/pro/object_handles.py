"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import uuid

try:
    import numpy as np
except ImportError:
    np = None
try:
    import pandas as pd
except ImportError:
    pd = None

from .. import ObjectCacheMissError, ObjectHandle, XlwingsError
from ..constants import ObjectHandleIcons
from ..conversion import Converter

# Reserved Entity property that carries the cache key (a UUID). It travels with the Entity
# (so it survives copy/paste and `=A1`) but is hidden from the user via excludeFrom (see
# write_value). User-supplied properties must never overwrite it.
# NOTE: this name is a contract with the frontend (custom-functions-code.js in
# xlwings-server), which reads it to substitute Entity arguments with their cache key.
RESERVED_PROPERTY = "object_handle_cache_key"

# Marker string the frontend substitutes for an Entity argument that isn't one of our
# object handles (e.g., a Stocks/Geography entity passed by mistake). It's a plain string
# (like the cache key) so it passes through xlwings' value cleaning unchanged - a dict
# without a "type" key would raise a KeyError there. A real cache key is a UUID, so it can
# never collide with this sentinel.
# NOTE: this string is a contract with the frontend (custom-functions-code.js in
# xlwings-server).
NOT_A_HANDLE_MARKER = "__xlwings_not_an_object_handle__"


class LRUObjectCache:
    """The default object cache: an in-process, capacity-bounded store holding raw Python
    objects, keyed by the handle's UUID.

    Every write stores under a fresh UUID (recalculating a producing function never
    overwrites the previous entry), so without a bound the cache would grow with every
    recalculation. The LRU cap keeps that in check: reads refresh an entry's recency, and
    writes evict the least recently used entries beyond ``maxsize``. Eviction is graceful
    by design - resolving an evicted handle raises ``ObjectCacheMissError``, which the
    runtimes turn into the "Expired object" card that a recalculation regenerates.

    Backends with their own storage (e.g. Redis in xlwings Server) replace the module's
    ``cache`` attribute with an object implementing ``get``/``set``/``clear``. Keys passed
    to the store are bare UUIDs; prefixing/partitioning and serialization are backend
    concerns - this store keeps plain object references.
    """

    def __init__(self, maxsize=100):
        # A maxsize < 1 is always a misconfiguration: 0 would evict every entry on
        # write (every handle instantly "expired"), negative values would crash the
        # eviction loop. Fail at construction, where the bad config is visible.
        if maxsize < 1:
            raise ValueError(f"maxsize must be a positive integer, got {maxsize!r}")
        self.maxsize = maxsize
        self._store = {}

    def get(self, cache_id):
        """Returns the cached object and refreshes its recency. None if absent."""
        if cache_id not in self._store:
            return None
        obj = self._store.pop(cache_id)
        self._store[cache_id] = obj
        return obj

    def set(self, cache_id, obj):
        self._store.pop(cache_id, None)
        self._store[cache_id] = obj
        while len(self._store) > self.maxsize:
            del self._store[next(iter(self._store))]

    def clear(self):
        self._store.clear()

    def __len__(self):
        return len(self._store)


# The active store. Runtimes with their own backend replace this, e.g.:
# from xlwings.pro import object_handles; object_handles.cache = RedisObjectCache()
cache = LRUObjectCache()


def _derived_properties(obj):
    """Returns the automatically derived Entity properties (type, shape, columns, index)
    for the given object. Used as the card's properties only when neither the function
    level (@ret/annotated hint) nor ObjectHandle supplies its own; the caller
    (write_value) decides which set to use."""
    properties = {
        "Type": {"type": "String", "basicValue": type(obj).__name__},
    }

    # Shape
    shape_value = _get_shape(obj)
    if shape_value:
        properties["Shape"] = {"type": "String", "basicValue": shape_value}

    # Columns
    if pd and isinstance(obj, pd.DataFrame):
        cols_info = ", ".join(f"{col} [{obj[col].dtype}]" for col in obj.columns)
        properties["Columns"] = {"type": "String", "basicValue": cols_info}

    # Index
    if pd and isinstance(obj, pd.DataFrame):
        index_type = type(obj.index).__name__
        index_length = len(obj.index)
        index_info = f"{index_type}: {index_length} entries"
        if index_length:
            # Only show the range for a non-empty index (obj.index[0] would raise on empty).
            index_info += f", {obj.index[0]} to {obj.index[-1]}"
        properties["Index"] = {"type": "String", "basicValue": index_info}

    return properties


def _get_shape(obj):
    if pd and isinstance(obj, pd.DataFrame):
        return f"{obj.shape}"
    if np and isinstance(obj, np.ndarray):
        return f"{obj.shape}"
    elif isinstance(obj, (list, tuple)):
        if obj and isinstance(obj[0], (list, tuple)):
            return f"({len(obj)}, {len(obj[0])})"
        return f"({len(obj)},)"
    else:
        try:
            return f"{len(obj)} (length)"
        except Exception:
            return None


def stale_object_handle():
    """Builds the Entity shown when a consumed object handle is no longer cached. It stays
    an Entity (rather than an error) so downstream cards still render and the user sees an
    actionable message. There's no refresh button (Excel entity cards can't host one), so
    the user recalculates to regenerate the object."""
    # Recovery is a full recalculation, which regenerates the object handles. We point at
    # Excel's built-in recalc rather than a custom button: a button could only do a full
    # recalc too (Excel can't enumerate which cells hold stale handles), so it would add
    # nothing over Calculate Now / Ctrl+Alt+F9, which every platform already provides.
    hint = "recalculate (Formulas > Calculate Now, or press Ctrl+Alt+F9 on the desktop)"
    icon = ObjectHandleIcons.warning
    if isinstance(icon, ObjectHandleIcons):
        icon = icon.value
    entity = {
        "type": "Entity",
        "text": "Expired object",
        "properties": {
            "Status": {
                "type": "String",
                "basicValue": f"This object is no longer cached. Please {hint}.",
            },
        },
        "layouts": {"compact": {"icon": icon}},
    }
    # Custom function results must be a 2D array. The normal return path goes through
    # conversion.write(), which wraps the entity in [[...]]; the stale path bypasses that
    # (it's returned by the runtime after catching ObjectCacheMissError), so wrap it here.
    # Without this, Excel receives a scalar where it expects a grid and renders the cell
    # as a #VALUE! error.
    return [[entity]]


class ObjectCacheConverter(Converter):
    @staticmethod
    def read_value(value, options):
        # For custom function args of type Entity, the frontend sends the object handle's
        # cache key (a UUID) instead of the cell value.
        if value == NOT_A_HANDLE_MARKER:
            raise XlwingsError("Argument is not an xlwings object handle")
        obj = cache.get(value)
        if obj is None:
            # Object expired or evicted. Raised so the runtime can turn it into a stale
            # object handle centrally instead of poisoning the function.
            raise ObjectCacheMissError(value)
        return obj

    @staticmethod
    def write_value(obj, options):
        # text/icon/properties can be set at the function level via the @ret decorator or an
        # annotated type hint; xw.ObjectHandle can additionally override them per object.
        text = options.get("text")
        icon = options.get("icon")
        user_properties = options.get("properties") or {}
        if isinstance(obj, ObjectHandle):
            text = obj.text or text
            icon = obj.icon or icon
            user_properties = obj.properties or user_properties
            obj = obj.obj

        if obj is None:
            # Usually a function that forgot its return statement. Failing here keeps
            # the error at the producing cell; a cached None would be indistinguishable
            # from a missing entry on read (the stores return None for absent keys) and
            # would surface as a misleading "Expired object" card on every consumer.
            raise XlwingsError(
                "Cannot create an object handle for None. "
                "Did your function forget to return a value?"
            )

        if RESERVED_PROPERTY in user_properties:
            raise XlwingsError(
                f"'{RESERVED_PROPERTY}' is a reserved object handle property name"
            )

        cache_id = str(uuid.uuid4())
        cache.set(cache_id, obj)

        obj_type = type(obj).__name__
        icon = icon or ObjectHandleIcons.generic
        if isinstance(icon, ObjectHandleIcons):
            icon = icon.value

        # If the user supplies properties (via @ret/annotated type hint or ObjectHandle),
        # they're the complete set shown on the card; otherwise fall back to the
        # automatically derived ones (type, shape, ...). The reserved cache key is always
        # added below, regardless.
        properties = (
            dict(user_properties) if user_properties else _derived_properties(obj)
        )
        # The reserved cache key is written last so it can never be shadowed. `excludeFrom`
        # hides it from the user (cardView: not on the card, autoComplete: not in formula
        # suggestions, dotNotation: not readable via FIELDVALUE()) while it still persists
        # on the Entity, so it survives copy/paste and `=A1`.
        #
        # Note: `calcCompare` is intentionally NOT excluded. The UUID is the only property
        # that changes when a handle is regenerated, so it must take part in recalc
        # change-detection - otherwise Excel considers the entity unchanged and skips
        # recalculating functions that consume it (e.g. =VIEW(A1) wouldn't update).
        properties[RESERVED_PROPERTY] = {
            "type": "String",
            "basicValue": cache_id,
            "propertyMetadata": {
                "excludeFrom": {
                    "cardView": True,
                    "autoComplete": True,
                    "dotNotation": True,
                },
            },
        }

        return {
            "type": "Entity",
            "text": text or obj_type,
            "properties": properties,
            "layouts": {"compact": {"icon": icon}},
        }
