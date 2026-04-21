"""
_com_helper.py — pywin32 late-bound COM utilities.

In SolidWorks Large Design Review (LDR / ViewOnly) mode, some COM methods are
registered as PROPERTYGET in the dispatch table instead of as callable methods.
When that happens, accessing obj.MethodName evaluates the property and returns
the value immediately.  Calling obj.MethodName() then tries to invoke that value
(e.g. a tuple), which raises:
    TypeError: 'tuple' object is not callable

sw_call() handles both modes transparently.
"""

from __future__ import annotations
from typing import Any


def sw_call(obj: Any, method: str, *args: Any) -> Any:
    """
    Safely call a SolidWorks COM method that may be a property in some SW modes.

    Usage:
        sheet_names = sw_call(swDraw, "GetSheetNames")
        sw_call(swDraw, "ActivateSheet", name)
        swSheet = sw_call(swDraw, "GetCurrentSheet")
    """
    attr = getattr(obj, method, None)
    if attr is None:
        raise AttributeError(f"COM method {method} not available")
    if callable(attr):
        try:
            return attr(*args)
        except TypeError:
            # attr may already be the value (property, not method) — return as-is
            return attr
        except Exception as e:
            raise RuntimeError(f"COM method {method} failed: {type(e).__name__}: {e}")
    # attr is already the value (PROPERTYGET pre-evaluated)
    return attr


def to_list(val: Any) -> list:
    """
    Normalise a COM SAFEARRAY / tuple / None / scalar to a plain Python list.

    SolidWorks COM methods that have ByRef output parameters return a tuple
    (retval, byref1, ...) via late-bound dispatch.  This helper finds the
    first iterable-but-not-string element and returns it as a list.
    If the whole value is already list-like, it returns list(val).
    """
    if val is None:
        return []
    if isinstance(val, str):
        return [val]
    # Tuple that looks like (count, array) — take the array element
    if isinstance(val, tuple):
        for item in val:
            if hasattr(item, "__iter__") and not isinstance(item, str):
                return list(item)
        # All elements are scalars — return them
        return list(val)
    if hasattr(val, "__iter__"):
        return list(val)
    return [val]
