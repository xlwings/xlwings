import sys
from types import ModuleType

import pytest

from xlwings.server import (
    custom_functions_call,
    custom_functions_meta,
    register_injectable_typehint,
)


def make_module(monkeypatch, name, source):
    module = ModuleType(name)
    monkeypatch.setitem(sys.modules, name, module)
    exec(compile(source, f"{name}.py", "exec"), module.__dict__)
    return module


def test_module_level_namespace(monkeypatch):
    module = make_module(
        monkeypatch,
        "module_level_namespace",
        """
from xlwings.server import func

__xlwings_func_namespace__ = "finance"

@func
def present_value(rate):
    return rate
""",
    )

    metadata = custom_functions_meta(module)

    assert metadata["functions"][0]["name"] == "FINANCE.PRESENT_VALUE"


def test_function_namespace_overrides_module_namespace(monkeypatch):
    module = make_module(
        monkeypatch,
        "function_namespace_override",
        """
from xlwings.server import func

__xlwings_func_namespace__ = "finance"

@func(namespace="statistics")
def mean(values):
    return values
""",
    )

    metadata = custom_functions_meta(module)

    assert metadata["functions"][0]["name"] == "STATISTICS.MEAN"


def test_namespace_comes_from_defining_module(monkeypatch):
    defining_module = make_module(
        monkeypatch,
        "namespaced_functions",
        """
from xlwings.server import func

__xlwings_func_namespace__ = "finance"

@func
def present_value(rate):
    return rate
""",
    )
    aggregate_module = ModuleType("custom_functions")
    aggregate_module.present_value = defining_module.present_value

    metadata = custom_functions_meta(aggregate_module)

    assert metadata["functions"][0]["name"] == "FINANCE.PRESENT_VALUE"


# --- Injectable typehints (e.g. CurrentUser), see GH 391 ---


class FakeCurrentUser:
    """Stands in for xlwings Server's CurrentUser."""

    def __init__(self, name="alice"):
        self.name = name
        self.roles = []

    async def has_required_roles(self, required_roles):
        return True


register_injectable_typehint(FakeCurrentUser)


@pytest.fixture
def user_types(monkeypatch):
    module = ModuleType("injectable_types")
    module.FakeCurrentUser = FakeCurrentUser
    monkeypatch.setitem(sys.modules, "injectable_types", module)
    return module


def make_call_data(func_name, args):
    return {
        "func_name": func_name,
        "args": args,
        "version": "dev",
        "client": "Office.js",
        "runtime": "1.4",
    }


async def call(module, func_name, args):
    user = FakeCurrentUser()
    return await custom_functions_call(
        make_call_data(func_name, args),
        module,
        current_user=user,
        typehint_to_value={FakeCurrentUser: user},
    )


@pytest.mark.parametrize(
    "signature",
    [
        "current_user: FakeCurrentUser, required_arg: str, *args: list[str]",
        "required_arg: str, *args: list[str], current_user: FakeCurrentUser = None",
        "required_arg: str, *args: list[str], current_user: FakeCurrentUser",
    ],
    ids=["injected-first", "keyword-only-with-default", "keyword-only"],
)
def test_injected_arg_with_varargs(monkeypatch, user_types, signature):
    """An injected param must work in any position, also combined with *args."""
    module = make_module(
        monkeypatch,
        f"injected_varargs_{abs(hash(signature))}",
        f"""
from xlwings.server import func
from injectable_types import FakeCurrentUser

@func
async def test_func({signature}):
    return f"{{current_user.name}}|{{required_arg}}|{{list(args)}}"
""",
    )

    # The injected param is not part of the Excel-facing signature
    metadata = custom_functions_meta(module)
    params = [param["name"] for param in metadata["functions"][0]["parameters"]]
    assert params == ["required_arg", "args"]

    import anyio

    result = anyio.run(call, module, "test_func", [["req"], ["a", "b"]])
    assert result == [["alice|req|['a', 'b']"]]


def test_injected_arg_in_the_middle(monkeypatch, user_types):
    import anyio

    module = make_module(
        monkeypatch,
        "injected_middle",
        """
from xlwings.server import func
from injectable_types import FakeCurrentUser

@func
async def test_func(a: str, current_user: FakeCurrentUser, b: str):
    return f"{current_user.name}|{a}|{b}"
""",
    )

    metadata = custom_functions_meta(module)
    params = [param["name"] for param in metadata["functions"][0]["parameters"]]
    assert params == ["a", "b"]

    result = anyio.run(call, module, "test_func", [["x"], ["y"]])
    assert result == [["alice|x|y"]]


def test_non_injected_keyword_only_arg_still_rejected(monkeypatch):
    """Only injected params are exempt from the keyword-argument restriction."""
    from xlwings import XlwingsError

    with pytest.raises(XlwingsError, match="keyword arguments"):
        make_module(
            monkeypatch,
            "kwonly_rejected",
            """
from xlwings.server import func

@func
def test_func(a: str, *args: list[str], b: str = None):
    return a
""",
        )


def test_book_is_not_a_custom_function_injectable(monkeypatch):
    """xw.Book is injected into custom *scripts*, not custom functions.

    Registering it globally would hide the param from the Excel-facing signature and
    silently shift the remaining args (GH 391 follow-up).
    """
    import xlwings as xw
    from xlwings.server import is_injectable_typehint

    assert not is_injectable_typehint(xw.Book)
    assert not is_injectable_typehint(xw.BookAsync)

    module = make_module(
        monkeypatch,
        "book_param_custom_function",
        """
from xlwings.server import func
import xlwings as xw

@func
def uses_book(book: xw.Book, name: str):
    return f"{book}|{name}"
""",
    )

    metadata = custom_functions_meta(module)
    params = [param["name"] for param in metadata["functions"][0]["parameters"]]
    assert params == ["book", "name"]


# WithScript: request a custom script to run after the custom function returns


def test_with_script_rejects_bad_args():
    import xlwings as xw

    with pytest.raises(xw.XlwingsError, match="must be a list"):
        xw.WithScript("value", "myscript", args="notalist")

    with pytest.raises(xw.XlwingsError, match="JSON-serializable"):
        xw.WithScript("value", "myscript", args=[object()])

    with pytest.raises(xw.XlwingsError, match="nest"):
        xw.WithScript(xw.WithScript("value", "inner"), "outer")

    # NaN/Infinity are accepted by json.dumps but aren't valid JSON: they must be
    # rejected here rather than blowing up when the HTTP response is serialized.
    for value in (float("nan"), float("inf"), float("-inf")):
        with pytest.raises(xw.XlwingsError, match="JSON-serializable"):
            xw.WithScript("value", "myscript", args=[value])


def test_with_script_normalizes_include_exclude():
    """The client splits these on "," so lists have to be normalized server-side."""
    import xlwings as xw

    assert xw.WithScript("v", "s", include=["Sheet1", "Sheet2"]).include == (
        "Sheet1,Sheet2"
    )
    assert xw.WithScript("v", "s", exclude=("A", "B")).exclude == "A,B"
    assert xw.WithScript("v", "s", include="Sheet1,Sheet2").include == "Sheet1,Sheet2"
    assert xw.WithScript("v", "s").include == ""

    with pytest.raises(xw.XlwingsError, match="sheet names as strings"):
        xw.WithScript("v", "s", include=[1, 2])
    with pytest.raises(xw.XlwingsError, match="must be a string or a list"):
        xw.WithScript("v", "s", exclude=42)


def test_with_script_rejects_include_and_exclude_together():
    """getBookData() throws on this, but only once the follow-up runs - long after the
    cell value was delivered successfully."""
    import xlwings as xw

    with pytest.raises(xw.XlwingsError, match="not both"):
        xw.WithScript("v", "s", include="A", exclude="B")
    with pytest.raises(xw.XlwingsError, match="not both"):
        xw.WithScript("v", "s", include=["A"], exclude=["B"])


def test_with_script_rejects_non_boolean_lazy():
    """A non-bool like "false" is truthy in JavaScript, so it would silently do the
    opposite of what was written."""
    import xlwings as xw

    for value in ("false", "", 0, 1, None):
        with pytest.raises(xw.XlwingsError, match="must be True or False"):
            xw.WithScript("v", "s", lazy=value)

    assert xw.WithScript("v", "s", lazy=True).lazy is True
    assert xw.WithScript("v", "s", lazy=False).lazy is False


def test_with_script_accepts_callable_or_string():
    import xlwings as xw

    def myscript(book):
        pass

    assert xw.WithScript("v", myscript).script_name == "myscript"
    assert xw.WithScript("v", "myscript").script_name == "myscript"


def test_with_script_is_not_a_sequence():
    """Ensure2DStage and _check_not_jagged special-case list/tuple: a WithScript that
    subclassed either would be spread across cells instead of treated as one value."""
    import xlwings as xw

    assert not isinstance(xw.WithScript("v", "s"), (list, tuple))


def test_with_script_payload():
    import xlwings as xw

    wrapped = xw.WithScript("v", "myscript", args=[1, "a"], include="Sheet1", lazy=True)
    assert wrapped.payload == {
        "script_name": "myscript",
        "args": [1, "a"],
        "include": "Sheet1",
        "exclude": "",
        "lazy": True,
    }


def test_with_script_returns_envelope(monkeypatch):
    import anyio

    import xlwings as xw

    module = make_module(
        monkeypatch,
        "with_script_envelope",
        """
from xlwings.server import func
import xlwings as xw

@func
def plain(name):
    return f"Hello {name}!"

@func
def wrapped(name):
    return xw.WithScript(f"Hello {name}!", "hello_args", args=[name], exclude="MySheet")
""",
    )

    # An unwrapped return must stay a bare value: regression guard for every existing
    # caller, which relies on custom_functions_call returning the converted value.
    assert anyio.run(call, module, "plain", ["x"]) == [["Hello x!"]]

    result = anyio.run(call, module, "wrapped", ["x"])
    assert isinstance(result, xw.CustomFunctionResult)
    assert result.value == [["Hello x!"]]
    assert result.script == {
        "script_name": "hello_args",
        "args": ["x"],
        "include": "",
        "exclude": "MySheet",
        "lazy": False,
    }


@pytest.mark.parametrize(
    "returns, options",
    [
        ("[[1, 2, 3], [4, 5, 6]]", "transpose=True"),
        ("[[1, 2, 3], [4, 5, 6]]", ""),
        ("'hello'", ""),
        ("42", ""),
        ("None", ""),
        ("dt.datetime(2026, 8, 14)", ""),
        ("dt.datetime(2026, 8, 14)", "transpose=True"),
        ("dt.datetime(2026, 8, 14)", 'date_format="dd/mm/yyyy"'),
        ("pd.DataFrame({'a': [1, 2], 'b': [3, 4]})", ""),
        ("pd.DataFrame({'a': [1, 2], 'b': [3, 4]})", "transpose=True"),
        ("pd.DataFrame({'a': [dt.datetime(2026, 8, 14)]})", 'date_format="dd/mm/yyyy"'),
        ("pd.Series([1, 2], name='s')", ""),
    ],
)
def test_with_script_conversion_matches_unwrapped(monkeypatch, returns, options):
    """The wrapped value must convert exactly as the bare value would.

    Unwrapping inside a Converter instead would run the generic write stages twice, so
    e.g. two transposes cancel out and silently return the wrong shape.
    """
    decorator = f"@ret({options})\n" if options else ""
    module = make_module(
        monkeypatch,
        f"with_script_conv_{abs(hash((returns, options)))}",
        f"""
import datetime as dt
import pandas as pd
from xlwings.server import func, ret
import xlwings as xw

@func
{decorator}def bare():
    return {returns}

@func
{decorator}def wrapped():
    return xw.WithScript({returns}, "s")
""",
    )

    import anyio

    bare = anyio.run(call, module, "bare", [])
    wrapped = anyio.run(call, module, "wrapped", [])
    assert wrapped.value == bare


def test_with_script_composes_with_object_handles(monkeypatch):
    """An explicit return converter must receive the INNER value, not the wrapper.

    ret_info drives converter selection, so unwrapping after conversion would hand the
    object cache the WithScript itself - Excel gets a valid-looking Entity card labeled
    "WithScript" and the script silently never runs.
    """
    import anyio

    import xlwings as xw
    from xlwings.pro.object_handles import CONVERTER_KEYS, ObjectCacheConverter

    # The object-handle converter is registered by the runtime (xlwings Server does this
    # at import), not by core itself.
    ObjectCacheConverter.register(*CONVERTER_KEYS)

    module = make_module(
        monkeypatch,
        "with_script_object_handle",
        """
from xlwings.server import func
import xlwings as xw

class Model:
    pass

@func
def make_model() -> object:
    return xw.WithScript(Model(), "refresh")
""",
    )

    result = anyio.run(call, module, "make_model", [])
    assert isinstance(result, xw.CustomFunctionResult)
    assert result.script["script_name"] == "refresh"
    entity = result.value[0][0]
    assert entity["type"] == "Entity"
    assert entity["text"] == "Model"


def test_with_script_rejected_in_streaming_functions(monkeypatch):
    """Streaming results travel over socket.io and never pass through the HTTP response
    that carries script requests, so this must fail rather than silently no-op."""
    import anyio

    results = []

    module = make_module(
        monkeypatch,
        "with_script_streaming",
        """
from xlwings.server import func
import xlwings as xw

@func
async def stream():
    yield xw.WithScript(1, "myscript")
""",
    )

    async def run_stream():
        import asyncio

        await custom_functions_call(
            {**make_call_data("stream", []), "task_key": "stream_"},
            module,
            streaming_callback=results.append,
        )
        task = next(t for t in asyncio.all_tasks() if t.get_name() == "xlwings-stream_")
        await task

    anyio.run(run_stream)

    assert len(results) == 1
    assert "WithScript is not supported in streaming functions" in results[0][0][0]
