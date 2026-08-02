import sys
from types import ModuleType

from xlwings.server import custom_functions_meta


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
