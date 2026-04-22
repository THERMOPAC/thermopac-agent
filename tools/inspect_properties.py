"""
inspect_properties.py — Dump ALL custom properties from CustomPropertyManager("").

Run with the drawing open in SolidWorks:

    python tools/inspect_properties.py

Shows:
- Every property name returned by GetNames
- Raw value (val, index 0) from Get5(useCached=True) and Get5(useCached=False)
- Resolved value (resolvedVal, index 1) from Get5(useCached=True) and Get5(useCached=False)
- Which of the 21 target properties are present vs missing
"""

from __future__ import annotations
import sys

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed.  Run:  pip install pywin32")
    sys.exit(1)

_TARGET_PROPERTIES = [
    "HYDRO_TEST_POSITION",
    "SHELL_IDP",  "SHELL_MOT",
    "TUBE_IDP",   "TUBE_MOT",
    "JACKET_IDP", "JACKET_MOT",
    "Drawing_Number", "Tag_No", "Equipment_Type", "Equipment_Configuration",
    "Design_Code", "Material_Code", "Inspection_By",
    "DrawnBy", "DrawnDate", "CheckedBy", "CheckedDate",
    "EngineeringApproval", "EngAppDate",
    "Revision",
]

def _com_call(obj, method, *args):
    raw = getattr(obj, method)
    return raw(*args) if callable(raw) else raw

def _get_ret(mgr, api, name, use_cached):
    try:
        ret = getattr(mgr, api)(name, use_cached)
        if isinstance(ret, str):
            return ret.strip(), ret.strip(), "str"
        if isinstance(ret, (list, tuple)):
            val      = str(ret[0]).strip() if len(ret) > 0 else ""
            resolved = str(ret[1]).strip() if len(ret) > 1 else ""
            return val, resolved, f"tuple[{len(ret)}]"
        return "", "", f"{type(ret).__name__}"
    except Exception as e:
        return "", "", f"ERR:{e}"


def run():
    pythoncom.CoInitialize()

    # Connect to running SolidWorks
    for progid in ("SldWorks.Application.27", "SldWorks.Application"):
        try:
            raw   = win32com.client.GetActiveObject(progid)
            swApp = win32com.client.Dispatch(raw)
            print(f"[SW] Connected via {progid}")
            break
        except Exception:
            continue
    else:
        print("ERROR: SolidWorks not running")
        sys.exit(1)

    swModel = swApp.ActiveDoc
    if swModel is None:
        print("ERROR: No active document")
        sys.exit(1)

    path = ""
    try:
        path = swModel.GetPathName()
    except Exception:
        pass
    print(f"[SW] Active document: {path or '(unsaved)'}")
    print()

    # Drawing-level CustomPropertyManager("")
    try:
        mgr = swModel.Extension.CustomPropertyManager("")
    except Exception as e:
        print(f"ERROR: Cannot get CustomPropertyManager: {e}")
        sys.exit(1)

    names = _com_call(mgr, "GetNames")
    if not names:
        print("CustomPropertyManager(\"\") returned NO properties.")
        print("The drawing has no drawing-level custom properties.")
        sys.exit(0)

    print(f"CustomPropertyManager(\"\") — {len(names)} properties found")
    print()

    # Header
    W = 30
    print(f"{'Property':<{W}}  {'useCached=True':<35}  {'useCached=False':<35}  {'type'}")
    print(f"{'':─<{W}}  {'':─<35}  {'':─<35}  {'':─<10}")
    print(f"{'':>{W}}  {'val[0] / resolvedVal[1]':<35}  {'val[0] / resolvedVal[1]':<35}")
    print()

    rows = {}
    for name in names:
        v_t,  r_t,  type_t  = _get_ret(mgr, "Get5", name, True)
        v_f,  r_f,  type_f  = _get_ret(mgr, "Get5", name, False)
        rows[name] = (v_t, r_t, type_t, v_f, r_f, type_f)
        cached_str = f"{v_t!r} / {r_t!r}"
        live_str   = f"{v_f!r} / {r_f!r}"
        print(f"{name:<{W}}  {cached_str:<35}  {live_str:<35}  {type_t}")

    print()
    print("── Target property status ───────────────────────────────────────────────")
    present = set(names)
    target_set = set(_TARGET_PROPERTIES)
    found_names = []
    missing_names = []
    for prop in _TARGET_PROPERTIES:
        if prop in present:
            v_t, r_t, _, v_f, r_f, _ = rows[prop]
            best = r_t or v_t or r_f or v_f
            status = f"FOUND  value={best!r}" if best else "PRESENT (value empty)"
            found_names.append(prop)
        else:
            status = "MISSING — not in CustomPropertyManager(\"\")"
            missing_names.append(prop)
        flag = "✓" if prop in present else "✗"
        print(f"  {flag}  {prop:<30}  {status}")

    print()
    extra = [n for n in names if n not in target_set]
    if extra:
        print(f"Extra properties (not in target list): {extra}")

    print()
    print(f"Summary: {len(found_names)}/21 target properties present in drawing-level CPM")
    if missing_names:
        print(f"Missing from drawing-level CPM: {missing_names}")
        print()
        print("Action required:")
        print("  In SolidWorks: File > Properties > Custom tab")
        print("  Add the missing properties with the exact names listed above.")
        print("  Property values must be stored at DRAWING level (not configuration).")

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    run()
