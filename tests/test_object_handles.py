"""
Tests for the object-handle converter and its default in-process LRU store
(xlwings/pro/object_handles.py). Backend-specific behavior (Redis, user partitioning,
HTTP status codes) is tested in xlwings-server.
"""

import types
import uuid

import pandas as pd
import pytest

import xlwings as xw
from xlwings.pro import object_handles as oh

Converter = oh.ObjectCacheConverter


@pytest.fixture
def anyio_backend():
    return "asyncio"


@pytest.fixture(autouse=True)
def _cache_context():
    # Register the converter as the runtimes (xlwings-server/Lite) do, so resolution via
    # conversion.read() works, and give each test a fresh default store.
    Converter.register(object, "object", "obj")
    original_cache = oh.cache
    oh.cache = oh.LRUObjectCache()
    yield
    oh.cache = original_cache


def _write(obj, options=None):
    """Writes an object handle and returns (entity, cache_key)."""
    entity = Converter.write_value(obj, options or {})
    key = entity["properties"][oh.RESERVED_PROPERTY]["basicValue"]
    return entity, key


def test_write_value_returns_entity_with_hidden_cache_key():
    df = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
    entity, key = _write(df)

    assert entity["type"] == "Entity"
    assert entity["text"] == "DataFrame"
    # The cache key is a UUID stored in the reserved property...
    uuid.UUID(key)
    # ...which is excluded from the card and formulas so it stays hidden from the user
    # while still travelling with the Entity (survives copy/paste and `=A1`).
    exclusions = entity["properties"][oh.RESERVED_PROPERTY]["propertyMetadata"][
        "excludeFrom"
    ]
    assert exclusions["cardView"] is True
    assert exclusions["dotNotation"] is True
    # calcCompare must NOT be excluded: the UUID has to take part in recalc
    # change-detection so consumers (e.g. =VIEW(A1)) recalculate when the handle changes.
    assert "calcCompare" not in exclusions
    # Derived properties are present.
    assert set(entity["properties"]) >= {"Type", "Shape", "Columns", "Index"}


def test_roundtrip_resolves_same_object():
    df = pd.DataFrame({"a": [1, 2]})
    _, key = _write(df)
    # The default store keeps raw object references: identity, not a serialized copy.
    assert Converter.read_value(key, {}) is df


def test_empty_dataframe_does_not_break_handle_creation():
    # An empty DataFrame is a valid result; deriving the Index property must not read
    # obj.index[0] (which would raise IndexError on an empty index).
    entity, key = _write(pd.DataFrame({"a": []}))
    assert entity["properties"]["Index"]["basicValue"] == "RangeIndex: 0 entries"
    assert Converter.read_value(key, {}).empty


def test_read_value_raises_on_cache_miss():
    with pytest.raises(xw.ObjectCacheMissError) as excinfo:
        Converter.read_value(str(uuid.uuid4()), {})
    assert excinfo.value.key is not None


def test_read_value_rejects_foreign_entity():
    # The frontend sends a plain marker string (not a dict) for a non-handle Entity, so it
    # passes through xlwings' value cleaning unchanged before reaching read_value.
    with pytest.raises(xw.XlwingsError, match="not an xlwings object handle"):
        Converter.read_value(oh.NOT_A_HANDLE_MARKER, {})


def test_object_handle_wrapper_customizes_presentation():
    df = pd.DataFrame({"a": [1]})
    handle = xw.ObjectHandle(
        df,
        text="1 row",
        icon=xw.ObjectHandleIcons.table,
        properties={"Region": {"type": "String", "basicValue": "EU"}},
    )
    entity, key = _write(handle)

    assert entity["text"] == "1 row"
    assert entity["layouts"]["compact"]["icon"] == xw.ObjectHandleIcons.table.value
    assert entity["properties"]["Region"]["basicValue"] == "EU"
    # Supplied properties are the complete set: the auto-derived ones are NOT shown
    # (only the supplied properties plus the always-present reserved cache key).
    assert set(entity["properties"]) == {"Region", oh.RESERVED_PROPERTY}
    # The wrapped object (not the wrapper) is what gets cached.
    assert Converter.read_value(key, {}) is df


def test_object_handle_without_properties_keeps_derived_ones():
    # When no properties are supplied, the auto-derived ones are still shown.
    handle = xw.ObjectHandle(pd.DataFrame({"a": [1]}), text="just text")
    entity, _ = _write(handle)
    assert set(entity["properties"]) >= {"Type", "Shape", "Columns", "Index"}


def test_function_level_properties_via_options():
    # text/icon/properties can also be set at the function level (via @ret or an annotated
    # type hint), which arrive in `options`. Properties there behave like the wrapper's:
    # they're the complete set, replacing the derived ones.
    entity = Converter.write_value(
        pd.DataFrame({"a": [1]}),
        {"properties": {"Region": {"type": "String", "basicValue": "EU"}}},
    )
    assert set(entity["properties"]) == {"Region", oh.RESERVED_PROPERTY}


def test_object_handle_properties_override_function_level():
    # The wrapper's per-object properties take precedence over the function-level ones.
    handle = xw.ObjectHandle(
        pd.DataFrame({"a": [1]}),
        properties={"FromWrapper": {"type": "String", "basicValue": "w"}},
    )
    entity = Converter.write_value(
        handle, {"properties": {"FromRet": {"type": "String", "basicValue": "r"}}}
    )
    assert set(entity["properties"]) == {"FromWrapper", oh.RESERVED_PROPERTY}


def test_object_handle_properties_cannot_shadow_reserved_key():
    handle = xw.ObjectHandle(
        pd.DataFrame({"a": [1]}),
        properties={oh.RESERVED_PROPERTY: {"type": "String", "basicValue": "hacked"}},
    )
    with pytest.raises(xw.XlwingsError):
        Converter.write_value(handle, {})


def test_stale_object_handle():
    # Custom function results must be a 2D array, so the stale entity is wrapped in [[...]].
    result = oh.stale_object_handle()
    assert isinstance(result, list) and isinstance(result[0], list)
    entity = result[0][0]
    # The text must not look like an Excel error literal (e.g. "#STALE!"), or Excel renders
    # the cell as a #VALUE! error instead of an object handle card.
    assert not entity["text"].startswith("#")
    # The card points at Excel's built-in recalc (no custom refresh button exists).
    status = entity["properties"]["Status"]["basicValue"]
    assert "recalculate" in status
    assert "Ctrl+Alt+F9" in status
    # The icon must be the serialized enum value (a string), not the enum object.
    assert isinstance(entity["layouts"]["compact"]["icon"], str)


# LRU store


def test_lru_evicts_oldest_beyond_maxsize():
    oh.cache = oh.LRUObjectCache(maxsize=2)
    _, key1 = _write("one")
    _, key2 = _write("two")
    _, key3 = _write("three")

    # The oldest entry fell off; resolving it degrades into the stale-card path.
    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key1, {})
    assert Converter.read_value(key2, {}) == "two"
    assert Converter.read_value(key3, {}) == "three"
    assert len(oh.cache) == 2


def test_lru_read_refreshes_recency():
    oh.cache = oh.LRUObjectCache(maxsize=2)
    _, key1 = _write("one")
    _, key2 = _write("two")
    # Reading key1 makes key2 the least recently used...
    Converter.read_value(key1, {})
    _, key3 = _write("three")

    # ...so the next write evicts key2, not the in-use key1.
    assert Converter.read_value(key1, {}) == "one"
    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key2, {})
    assert Converter.read_value(key3, {}) == "three"


def test_lru_clear():
    _, key = _write("one")
    oh.cache.clear()
    assert len(oh.cache) == 0
    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key, {})


# Type hints (ObjectHandle[T] / CachedObject[T])


def test_object_handle_type_hint_resolves_via_cache():
    # ObjectHandle[T] opts an argument into object-cache resolution while keeping T as the
    # type seen by editors/type checkers, instead of having to annotate the arg as object.
    from xlwings.server import func

    @func
    async def view(obj: xw.ObjectHandle[pd.DataFrame]):
        return obj

    arg_options = view.__xlfunc__["args"][0]["options"]
    # The arg is converted via the object cache (registered for `object`)...
    assert arg_options["convert"] is object
    # ...while the annotation still carries the real type for static tooling: it resolves
    # to Annotated[pd.DataFrame, ObjectHandle], which type checkers read as pd.DataFrame.
    annotation = view.__annotations__["obj"]
    assert annotation.__args__[0] is pd.DataFrame
    assert xw.ObjectHandle in annotation.__metadata__


def test_bare_object_handle_return_hint_is_alias_for_object():
    # `-> ObjectHandle` is an alias for `-> object`: it converts via the object cache.
    from xlwings.server import func

    @func
    async def make() -> xw.ObjectHandle:
        return pd.DataFrame({"a": [1]})

    assert make.__xlfunc__["ret"]["options"]["convert"] is object


@pytest.mark.anyio
async def test_object_handle_argument_resolves_to_wrapped_object():
    # End-to-end: a function annotated with CachedObject[pd.DataFrame] receives the cached
    # DataFrame (not the cache key) in its body.
    from xlwings.server import custom_functions_call, func

    @func
    async def consume(df: xw.CachedObject[pd.DataFrame]):
        # If resolution works, `df` is a DataFrame and `.shape` succeeds. Return it as
        # plain values (not an object handle) so we can assert on the result directly.
        return list(df.shape)

    module = types.ModuleType("_oh_test_module")
    module.consume = consume

    # Write a handle, then call the function with the handle's cache key as its argument.
    _, key = _write(pd.DataFrame({"a": [1, 2, 3]}))
    result = await custom_functions_call(
        {
            "func_name": "consume",
            "args": [[[key]]],
            "version": xw.__version__,
            "client": "Office.js",
            "runtime": "1.4",
        },
        module,
    )
    assert result == [[3, 1]]  # (3 rows, 1 col)
