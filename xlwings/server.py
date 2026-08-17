# ruff: noqa: F401
from . import CustomFunctionResult, WithScript
from .pro.caller import Caller, caller_from_address, parse_caller_address
from .pro.object_handles import (
    CONVERTER_KEYS as OBJECT_CONVERTER_KEYS,
    ObjectCacheConverter,
    stale_object_handle,
)
from .pro.udfs_officejs import (
    custom_functions_call,
    custom_functions_code,
    custom_functions_meta,
    custom_scripts_call,
    custom_scripts_meta,
    get_custom_function_namespace,
    is_injectable_typehint,
    register_injectable_typehint,
    script,
    sio_cancel_task,
    sio_connect,
    sio_custom_function_call,
    sio_disconnect,
    xlarg as arg,
    xlfunc as func,
    xlret as ret,
)

# xlwings ships py.typed, so type checkers hold this module to PEP 484's re-export rules:
# a plain `from X import Y` doesn't count as a public re-export, and importing these names
# from here reports as reportPrivateImportUsage ("Import from xlwings.pro... instead").
# Listing them in __all__ marks them public, so that the documented import path is also
# the one type checkers accept.
__all__ = (
    "CustomFunctionResult",
    "WithScript",
    "Caller",
    "caller_from_address",
    "parse_caller_address",
    "OBJECT_CONVERTER_KEYS",
    "ObjectCacheConverter",
    "stale_object_handle",
    "custom_functions_call",
    "custom_functions_code",
    "custom_functions_meta",
    "custom_scripts_call",
    "custom_scripts_meta",
    "get_custom_function_namespace",
    "is_injectable_typehint",
    "register_injectable_typehint",
    "script",
    "sio_cancel_task",
    "sio_connect",
    "sio_custom_function_call",
    "sio_disconnect",
    "arg",
    "func",
    "ret",
)
