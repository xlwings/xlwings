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
