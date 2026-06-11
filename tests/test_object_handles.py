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
    Converter.register(*oh.CONVERTER_KEYS)
    original_cache = oh.cache
    oh.cache = oh.LRUObjectCache()
    oh._producer_cache_ids.clear()
    yield
    oh.cache = original_cache
    oh._producer_cache_ids.clear()


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


def test_none_cannot_be_cached():
    # Usually a function that forgot its return statement. None must be rejected at
    # write time: the stores use None for "missing", so a cached None would read as a
    # permanently "Expired object" card on every consumer.
    with pytest.raises(xw.XlwingsError, match="forget to return"):
        Converter.write_value(None, {})
    # The wrapper can't smuggle one in either.
    with pytest.raises(xw.XlwingsError, match="forget to return"):
        Converter.write_value(xw.ObjectHandle(None, text="empty"), {})


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


def test_lru_rejects_nonpositive_maxsize():
    # 0 would evict every entry on write; negatives would crash the eviction loop.
    # Both are misconfigurations (e.g. via XLWINGS_OBJECT_CACHE_MAXSIZE) that must fail
    # at construction, not on the first object-handle write.
    with pytest.raises(ValueError, match="positive integer"):
        oh.LRUObjectCache(maxsize=0)
    with pytest.raises(ValueError, match="positive integer"):
        oh.LRUObjectCache(maxsize=-5)


def test_lru_clear():
    _, key = _write("one")
    oh.cache.clear()
    assert len(oh.cache) == 0
    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key, {})


# Superseded-generation cleanup (evict_superseded)


def test_recalculation_evicts_superseded_generation():
    # Every write stores under a fresh UUID, so a recalculating cell would otherwise
    # leave its previous generation in the cache until LRU eviction.
    addr = "Excel[Book1.xlsx]Sheet1!A1"
    entity1, key1 = _write("gen1")
    oh.evict_superseded(addr, [[entity1]])
    entity2, key2 = _write("gen2")
    oh.evict_superseded(addr, [[entity2]])

    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key1, {})
    assert Converter.read_value(key2, {}) == "gen2"
    assert len(oh.cache) == 1


def test_superseded_eviction_is_scoped_to_the_caller():
    # One cell's recalculation must never touch another cell's live handle.
    entity_a, key_a = _write("a")
    oh.evict_superseded("Excel[Book1.xlsx]Sheet1!A1", [[entity_a]])
    entity_b, key_b = _write("b")
    oh.evict_superseded("Excel[Book1.xlsx]Sheet1!B1", [[entity_b]])

    assert Converter.read_value(key_a, {}) == "a"
    assert Converter.read_value(key_b, {}) == "b"


def test_non_handle_result_clears_previous_generation():
    # A cell that used to produce a handle but now returns a plain value no longer
    # references its old entry, so the entry (and the tracking) must go.
    addr = "Excel[Book1.xlsx]Sheet1!A1"
    entity, key = _write("gen1")
    oh.evict_superseded(addr, [[entity]])
    oh.evict_superseded(addr, [["plain value"]])

    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key, {})
    assert addr not in oh._producer_cache_ids


def test_superseded_eviction_is_scoped_to_the_user():
    # Caller addresses are not unique across users (everybody has a "Book1.xlsx"), so
    # one user's recalculation must not evict another user's live handle.
    addr = "Excel[Book1.xlsx]Sheet1!A1"
    entity_a, key_a = _write("user a's object")
    oh.evict_superseded(addr, [[entity_a]], user_id="user_a")
    entity_b, key_b = _write("user b's object")
    oh.evict_superseded(addr, [[entity_b]], user_id="user_b")

    assert Converter.read_value(key_a, {}) == "user a's object"
    assert Converter.read_value(key_b, {}) == "user b's object"


def test_superseded_eviction_is_scoped_to_the_session():
    # With auth disabled all users share one user id, so the frontend's per-runtime
    # session id is what keeps two users' independent copies of an identically-named
    # workbook ("workbook1.xlsx"!A1) from evicting each other's live handles.
    addr = "Excel[workbook1.xlsx]Sheet1!A1"
    entity_a, key_a = _write("session a's object")
    oh.evict_superseded(addr, [[entity_a]], user_id="n/a", session_id="session_a")
    entity_b, key_b = _write("session b's object")
    oh.evict_superseded(addr, [[entity_b]], user_id="n/a", session_id="session_b")

    assert Converter.read_value(key_a, {}) == "session a's object"
    assert Converter.read_value(key_b, {}) == "session b's object"


def test_store_can_take_over_producer_tracking():
    # A store implementing evict_superseded(scope, new_ids) (e.g. Redis in
    # xlwings-server, where the map must be visible to all workers) replaces the
    # in-process map entirely.
    calls = []

    class TrackingStore(oh.LRUObjectCache):
        def evict_superseded(self, scope, new_ids):
            calls.append((scope, new_ids))

    oh.cache = TrackingStore()
    entity, key = _write("gen1")
    oh.evict_superseded("Excel[Book1.xlsx]Sheet1!A1", [[entity]], user_id="user_a")

    assert calls == [("user_a:Excel[Book1.xlsx]Sheet1!A1", {key})]
    assert not oh._producer_cache_ids


def test_evict_superseded_tolerates_store_without_delete():
    # Custom backends only need get/set/clear; without delete, superseded entries are
    # left to the store's own expiry policy, but tracking must still advance.
    class MinimalStore:
        def __init__(self):
            self._d = {}

        def get(self, cache_id):
            return self._d.get(cache_id)

        def set(self, cache_id, obj):
            self._d[cache_id] = obj

        def clear(self):
            self._d.clear()

    oh.cache = MinimalStore()
    addr = "Excel[Book1.xlsx]Sheet1!A1"
    entity1, key1 = _write("gen1")
    oh.evict_superseded(addr, [[entity1]])
    entity2, key2 = _write("gen2")
    oh.evict_superseded(addr, [[entity2]])

    assert oh._producer_cache_ids[addr] == {key2}
    # No delete method: the old entry stays (backend expiry's job), but nothing blew up.
    assert Converter.read_value(key1, {}) == "gen1"


@pytest.mark.anyio
async def test_custom_functions_call_evicts_previous_generation():
    # End-to-end: recalculating a producing cell (same caller_address) replaces its
    # cache entry instead of accumulating one orphan per recalculation.
    from xlwings.server import custom_functions_call, func

    @func
    async def make() -> object:
        return pd.DataFrame({"a": [1]})

    module = types.ModuleType("_oh_test_module_producer")
    module.make = make

    data = {
        "func_name": "make",
        "args": [],
        "version": xw.__version__,
        "client": "Office.js",
        "runtime": "1.4",
        "caller_address": "Excel[Book1.xlsx]Sheet1!A1",
        "session_id": "test-session",
    }
    result1 = await custom_functions_call(dict(data), module)
    key1 = result1[0][0]["properties"][oh.RESERVED_PROPERTY]["basicValue"]
    result2 = await custom_functions_call(dict(data), module)
    key2 = result2[0][0]["properties"][oh.RESERVED_PROPERTY]["basicValue"]

    with pytest.raises(xw.ObjectCacheMissError):
        Converter.read_value(key1, {})
    assert Converter.read_value(key2, {}) is not None
    assert len(oh.cache) == 1


@pytest.mark.anyio
async def test_non_producing_functions_skip_producer_tracking():
    # Functions that can't return handles must not pay the producer-map lookup on every
    # call (a Redis round trip per call in xlwings Server). Consequence: a cell whose
    # formula changes from a producing to a non-producing function keeps its last
    # generation until LRU eviction/expiry - the same backstop that covers deleted
    # formulas, which never trigger a call at all.
    from xlwings.server import custom_functions_call, func

    @func
    async def make() -> object:
        return pd.DataFrame({"a": [1]})

    @func
    async def plain():
        return 1

    module = types.ModuleType("_oh_test_module_plain")
    module.make = make
    module.plain = plain

    data = {
        "args": [],
        "version": xw.__version__,
        "client": "Office.js",
        "runtime": "1.4",
        "caller_address": "Excel[Book1.xlsx]Sheet1!A1",
        "session_id": "test-session",
    }
    result = await custom_functions_call({**data, "func_name": "make"}, module)
    key = result[0][0]["properties"][oh.RESERVED_PROPERTY]["basicValue"]
    await custom_functions_call({**data, "func_name": "plain"}, module)

    # The plain call neither deleted the old generation nor touched the tracking.
    assert Converter.read_value(key, {}) is not None
    assert oh._producer_cache_ids == {"test-session:Excel[Book1.xlsx]Sheet1!A1": {key}}


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


def test_object_handle_return_hints_are_aliases_for_object():
    # `-> ObjectHandle` and `-> ObjectHandle[T]` are aliases for `-> object`: they
    # convert via the object cache. custom_functions_call relies on this to decide
    # which functions take part in superseded-generation tracking.
    from xlwings.server import func

    @func
    async def make() -> xw.ObjectHandle:
        return pd.DataFrame({"a": [1]})

    @func
    async def make_typed() -> xw.ObjectHandle[pd.DataFrame]:
        return pd.DataFrame({"a": [1]})

    assert make.__xlfunc__["ret"]["options"]["convert"] is object
    assert make_typed.__xlfunc__["ret"]["options"]["convert"] is object


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
