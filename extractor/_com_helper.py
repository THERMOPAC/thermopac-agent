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

try:
    import win32com.client
    import pythoncom
except Exception:
    win32com = None
    pythoncom = None


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


def _query_dispatch_interface(obj: Any, interface_names: tuple[str, ...]) -> Any:
    if obj is None or win32com is None or pythoncom is None:
        return None
    ole = getattr(obj, "_oleobj_", None)
    if ole is None:
        return None
    try:
        ti = ole.GetTypeInfo()
        typelib, _ = ti.GetContainingTypeLib()
        count = typelib.GetTypeInfoCount()
    except Exception:
        return None
    wanted = {name.lower() for name in interface_names}
    for idx in range(count):
        try:
            type_info = typelib.GetTypeInfo(idx)
            name = str(type_info.GetDocumentation(-1)[0] or "")
            if name.lower() not in wanted:
                continue
            guid = str(type_info.GetTypeAttr()[0])
            qi = ole.QueryInterface(pythoncom.MakeIID(guid), pythoncom.IID_IDispatch)
            if qi is not None:
                return win32com.client.Dispatch(qi)
        except Exception:
            continue
    return None


def cast_to_drawing_doc(obj: Any) -> Any:
    if obj is None or win32com is None:
        return obj
    for cast_name in ("IDrawingDoc", "DrawingDoc"):
        try:
            casted = win32com.client.CastTo(obj, cast_name)
            if casted is not None:
                return casted
        except Exception:
            pass
    queried = _query_dispatch_interface(obj, ("IDrawingDoc", "DrawingDoc"))
    if queried is not None:
        return queried
    return obj


def get_active_doc(swApp: Any) -> Any:
    for method in ("ActiveDoc", "IActiveDoc2"):
        try:
            value = getattr(swApp, method, None)
            if value is None:
                continue
            doc = value() if callable(value) else value
            if doc is not None:
                return doc
        except Exception:
            pass
    return None


def refetch_active_drawing_doc(swApp: Any, fallback: Any = None) -> Any:
    active = get_active_doc(swApp)
    if active is not None:
        return cast_to_drawing_doc(active)
    return cast_to_drawing_doc(fallback)


def activate_sheet_and_get_current_sheet(swApp: Any, swDraw: Any, sheet_name: str, logger: Any = None) -> tuple[Any, Any]:
    sw_call(swDraw, "ActivateSheet", sheet_name)
    active_draw = refetch_active_drawing_doc(swApp, swDraw)
    for candidate in (active_draw, swDraw):
        try:
            swSheet = sw_call(candidate, "GetCurrentSheet")
            if swSheet is not None:
                return candidate, swSheet
        except Exception as e:
            if logger is not None:
                logger.debug(f"[COM] GetCurrentSheet after ActivateSheet('{sheet_name}') failed on {type(candidate).__name__}: {e}")
    return active_draw, None


def iter_drawing_views(swDraw: Any, sheet_names: list | None = None) -> list[tuple[str, Any]]:
    try:
        raw = sw_call(swDraw, "GetViews")
    except Exception:
        return []
    groups = to_list(raw)
    result: list[tuple[str, Any]] = []
    current_sheet = ""
    for idx, group in enumerate(groups):
        if isinstance(group, str):
            current_sheet = group
            continue
        if isinstance(group, (list, tuple)):
            values = list(group)
            sheet = sheet_names[idx] if sheet_names and idx < len(sheet_names) else current_sheet
            if values and isinstance(values[0], str):
                sheet = values[0]
                values = values[1:]
            for view in values:
                if view is not None and not isinstance(view, str):
                    result.append((str(sheet or ""), view))
        elif group is not None:
            sheet = sheet_names[idx] if sheet_names and idx < len(sheet_names) else current_sheet
            result.append((str(sheet or ""), group))
    return result


def com_type_summary(obj: Any) -> dict:
    summary = {
        "python_type": type(obj).__name__ if obj is not None else "None",
        "python_module": type(obj).__module__ if obj is not None else "",
        "repr": repr(obj)[:240] if obj is not None else "None",
        "typeinfo_name": "",
        "typeinfo_doc": "",
        "typeattr_guid": "",
    }
    try:
        ole = getattr(obj, "_oleobj_", None)
        if ole is not None:
            ti = ole.GetTypeInfo()
            if ti is not None:
                try:
                    doc = ti.GetDocumentation(-1)
                    summary["typeinfo_name"] = str(doc[0] or "")
                    summary["typeinfo_doc"] = str(doc[1] or "")
                except Exception:
                    pass
                try:
                    attr = ti.GetTypeAttr()
                    summary["typeattr_guid"] = str(attr[0])
                except Exception:
                    pass
    except Exception:
        pass
    return summary


def probe_method(obj: Any, method: str, *args: Any) -> dict:
    info = {
        "method": method,
        "has_attr": False,
        "callable": False,
        "call_ok": False,
        "result_type": "",
        "result_preview": "",
        "error": "",
    }
    try:
        attr = getattr(obj, method)
        info["has_attr"] = True
        info["callable"] = callable(attr)
        value = attr(*args) if callable(attr) else attr
        info["call_ok"] = True
        info["result_type"] = type(value).__name__
        info["result_preview"] = repr(value)[:180]
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
    return info
