"""
solidworks_extractor.py — Opens a dedicated SolidWorks instance and runs all
10 extraction modules sequentially.

Safety contract (from baseline v3):
  - v1.0.45 test mode attaches to a running SolidWorks session when present so
    pre-opened referenced models can be detected and measured
  - If no running session exists, COM starts SolidWorks normally
  - OpenDoc6 with ReadOnly | Silent flags
  - CloseDoc + ExitApp always in finally block
  - Never calls Save / SaveAs
  - Works on temp copy only
  - Checks cancel_event between modules
"""

from __future__ import annotations
import os
import threading
import time

try:
    import win32com.client
    import win32com.client.gencache
    import pythoncom
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

def _connect_sw_application(progid: str, logger):
    """
    Connect to SolidWorks with early binding when available, otherwise fall back
    to late binding. Some Windows/SolidWorks installations successfully run
    makepy but still fail EnsureDispatch with "cannot automate the makepy
    process", so this must not block extraction.
    """
    for pid in (progid, "SldWorks.Application"):
        try:
            active = win32com.client.GetActiveObject(pid)
            sw_app = win32com.client.Dispatch(active)
            logger.info(f"[COM] Binding mode: late (attached to running session via {pid})")
            version = "unknown"
            for attr in ("RevisionNumber", "Version"):
                try:
                    value = getattr(sw_app, attr)
                    version = value() if callable(value) else value
                    if version:
                        break
                except Exception:
                    pass
            logger.info(f"[COM] SolidWorks version detected: {version}")
            return sw_app, "late-attached", True
        except Exception:
            pass
    try:
        sw_app = win32com.client.gencache.EnsureDispatch(progid)
        binding_mode = "early"
        logger.info("[COM] Binding mode: early (EnsureDispatch)")
    except Exception as e:
        logger.warning(f"[COM] EnsureDispatch failed: {type(e).__name__}: {e}")
        sw_app = win32com.client.Dispatch(progid)
        binding_mode = "late"
        logger.info("[COM] Binding mode: late (Dispatch)")

    version = "unknown"
    for attr in ("RevisionNumber", "Version"):
        try:
            value = getattr(sw_app, attr)
            version = value() if callable(value) else value
            if version:
                break
        except Exception:
            pass
    logger.info(f"[COM] SolidWorks version detected: {version}")
    return sw_app, binding_mode, False

try:
    import win32gui
    import win32con
    WIN32GUI_AVAILABLE = True
except ImportError:
    WIN32GUI_AVAILABLE = False


from extractor.verify_custom_properties import verify_custom_properties
from extractor._com_helper import (
    cast_to_drawing_doc,
    com_type_summary,
    get_active_doc,
    probe_method,
    refetch_active_drawing_doc,
    sw_call,
)

# SolidWorks constants  (swOpenDocOptions_e)
SW_DOC_DRAWING           = 3
SW_OPEN_SILENT           = 1     # swOpenDocOptions_Silent      — suppresses ALL missing-ref dialogs
SW_OPEN_READ_ONLY        = 2     # swOpenDocOptions_ReadOnly
SW_OPEN_VIEW_ONLY        = 4     # swOpenDocOptions_ViewOnly    — Large Design Review, no 3-D load
SW_OPEN_RAPID_DRAFT      = 8     # swOpenDocOptions_RapidDraft / Detailing Mode for drawings
SW_OPEN_LOAD_MODEL       = 16    # swOpenDocOptions_LoadModel — detached drawing model-load fallback only
SW_OPEN_OVERRIDE_DEFAULT = 64    # swOpenDocOptions_OverrideDefaultLoadLightweight
SW_OPEN_LOAD_LIGHTWEIGHT = 128   # swOpenDocOptions_LoadLightweight
# NOTE: 64 = swOpenDocOptions_OverrideDefaultLoadedData (was wrongly used as SW_OPEN_SILENT before v1.0.4)

_DOC_TYPES = {
    1: "Part",
    2: "Assembly",
    3: "Drawing",
}

_OPEN_OPTION_NAMES = [
    (SW_OPEN_SILENT, "Silent"),
    (SW_OPEN_READ_ONLY, "ReadOnly"),
    (SW_OPEN_VIEW_ONLY, "ViewOnly/LargeDesignReview"),
    (SW_OPEN_RAPID_DRAFT, "RapidDraft/DetailingMode"),
    (SW_OPEN_LOAD_MODEL, "LoadModel"),
    (SW_OPEN_OVERRIDE_DEFAULT, "OverrideDefaultLoadedData"),
    (SW_OPEN_LOAD_LIGHTWEIGHT, "LoadLightweight"),
]

# swOpenDocError_e decode map (for diagnostics)
_SW_OPEN_ERRORS = {
    1:       "GenericError",
    2:       "FileNotFound",
    4:       "LockedFile",
    8:       "UserDeclined",
    128:     "AlreadyOpen",
    512:     "FileReadOnly",
    1024:    "ConversionRequired",
    4096:    "NeedToActivateDoc",
    65536:   "ExternalRefsNotLoaded",   # ext refs unavailable; LDR opens OK, full-mode fails
    131072:  "GenRenderMat",
    262144:  "IdMismatch",
    524288:  "AddToCurrentDoc",
    1048576: "SWXOnly",
    2097152: "HeavyWeightComponents",  # referenced 3-D files missing from temp folder
}

def _decode_sw_error(code: int) -> str:
    if code == 0:
        return "0 (none)"
    parts = [name for bit, name in _SW_OPEN_ERRORS.items() if code & bit]
    return f"{code} ({', '.join(parts) if parts else 'unknown'})"


def _decode_open_options(options: int) -> str:
    if options == 0:
        return "0 (none)"
    names = [name for bit, name in _OPEN_OPTION_NAMES if options & bit]
    return f"{options} ({' | '.join(names) if names else 'unknown'})"


def _reference_load_mode(options: int) -> str:
    if options & SW_OPEN_VIEW_ONLY:
        return "deferred/view-only; referenced 3-D model load avoided"
    if options & SW_OPEN_RAPID_DRAFT:
        return "detailing/rapid-draft; drawing data loaded while referenced model load is minimized"
    if options & SW_OPEN_LOAD_MODEL:
        return "detached drawing model load; higher risk of resolved/heavyweight references"
    if (options & SW_OPEN_OVERRIDE_DEFAULT) and (options & SW_OPEN_LOAD_LIGHTWEIGHT):
        return "override default + load lightweight; avoids forcing heavyweight references"
    if options & SW_OPEN_OVERRIDE_DEFAULT:
        return "default loaded data overridden; avoids forcing heavyweight references where possible"
    if options & SW_OPEN_SILENT:
        return "silent full drawing; references may resolve per SW defaults/search paths"
    return "interactive full drawing; references may resolve per SW defaults/search paths"


def _safe_doc_type(model) -> tuple[int | None, str]:
    if model is None:
        return None, "None"
    try:
        raw = model.GetType
        doc_type = raw() if callable(raw) else int(raw)
        return doc_type, _DOC_TYPES.get(doc_type, f"Unknown({doc_type})")
    except Exception as e:
        return None, f"unavailable ({type(e).__name__}: {e})"


def _log_open_attempt(logger, api: str, label: str, doc_type: int, options: int | None, mode: str, result, err: int, warn: int):
    opened_type, opened_type_name = _safe_doc_type(result)
    requested_type = _DOC_TYPES.get(doc_type, f"Unknown({doc_type})")
    opt_text = "DocSpec(properties logged above)" if options is None else _decode_open_options(options)
    ref_mode = mode if mode else (_reference_load_mode(options) if options is not None else "docSpec full drawing; references may resolve per SW defaults/search paths")
    logger.info(
        f"[Extractor] {api} {label}: requested_doc_type={doc_type} ({requested_type}) "
        f"open_options={opt_text} solidworks_open_mode='{ref_mode}' "
        f"model={'OK' if result else 'None'} opened_doc_type={opened_type_name} "
        f"errors={_decode_sw_error(err)} warnings={_decode_sw_error(warn)}"
    )


def _log_reference_diagnostics(swApp, temp_path: str, logger, stage: str) -> dict:
    logger.info(f"[Extractor] Reference diagnostics ({stage}) for {os.path.basename(temp_path)}")
    out = {
        "stage": stage,
        "search_folders": {},
        "dependency_call": "",
        "dependency_entries": 0,
        "path_entries": [],
        "missing_paths": [],
    }
    try:
        for folder_type in [1, 2, 4, 7]:
            try:
                paths = swApp.GetSearchFolders(folder_type)
                out["search_folders"][str(folder_type)] = paths if paths else ""
                logger.info(f"[Extractor] SearchFolders type={folder_type}: {paths if paths else '(empty)'}")
            except Exception as e:
                out["search_folders"][str(folder_type)] = f"unavailable ({type(e).__name__}: {e})"
                logger.info(f"[Extractor] SearchFolders type={folder_type}: unavailable ({type(e).__name__}: {e})")
    except Exception:
        pass
    dependency_calls = [
        ("GetDocumentDependencies2(False, True, False)", lambda: swApp.GetDocumentDependencies2(temp_path, False, True, False)),
        ("GetDocumentDependencies2(True, True, False)", lambda: swApp.GetDocumentDependencies2(temp_path, True, True, False)),
        ("GetDocumentDependencies(temp)", lambda: swApp.GetDocumentDependencies(temp_path)),
    ]
    for label, fn in dependency_calls:
        try:
            deps = fn()
            if deps is None:
                logger.info(f"[Extractor] Reference diagnostics {label}: None")
                continue
            dep_list = list(deps) if isinstance(deps, (list, tuple)) else [deps]
            missing = []
            paths = []
            for value in dep_list:
                text = str(value)
                if text.lower().endswith((".sldprt", ".sldasm", ".slddrw")) or "\\" in text or "/" in text:
                    paths.append(text)
                    if text and not os.path.exists(text):
                        missing.append(text)
            out["dependency_call"] = label
            out["dependency_entries"] = len(dep_list)
            out["path_entries"] = paths[:200]
            out["missing_paths"] = missing[:200]
            logger.info(f"[Extractor] Reference diagnostics {label}: entries={len(dep_list)} path_entries={len(paths)} missing_paths={len(missing)}")
            for item in missing[:30]:
                logger.warning(f"[Extractor] Missing/problematic referenced model: {item}")
            if len(missing) > 30:
                logger.warning(f"[Extractor] Missing/problematic referenced models truncated: {len(missing) - 30} more")
            return out
        except Exception as e:
            logger.info(f"[Extractor] Reference diagnostics {label} failed: {type(e).__name__}: {e}")
    return out


def _norm_path(path: str) -> str:
    try:
        return os.path.normcase(os.path.abspath(str(path or "").strip().strip('"')))
    except Exception:
        return str(path or "").strip().lower()


def _safe_model_path(model) -> str:
    if model is None:
        return ""
    for name in ("GetPathName", "PathName"):
        try:
            value = getattr(model, name, None)
            if value is None:
                continue
            path = value() if callable(value) else value
            if path:
                return str(path)
        except Exception:
            pass
    return ""


def _safe_model_title(model) -> str:
    if model is None:
        return ""
    for name in ("GetTitle", "Title"):
        try:
            value = getattr(model, name, None)
            if value is None:
                continue
            title = value() if callable(value) else value
            if title:
                return str(title)
        except Exception:
            pass
    return ""


def _iter_open_documents(swApp, logger) -> list[dict]:
    docs: list[dict] = []
    seen = set()
    try:
        doc = swApp.GetFirstDocument()
    except Exception as e:
        logger.info(f"[Extractor] Open-document inventory unavailable: {type(e).__name__}: {e}")
        return docs
    while doc is not None:
        try:
            path = _safe_model_path(doc)
            title = _safe_model_title(doc)
            doc_type, doc_type_name = _safe_doc_type(doc)
            key = _norm_path(path) if path else f"title:{title.lower()}"
            if key and key not in seen:
                docs.append({
                    "path": path,
                    "title": title,
                    "doc_type": doc_type,
                    "doc_type_name": doc_type_name,
                })
                seen.add(key)
        except Exception as e:
            logger.info(f"[Extractor] Open-document inventory item skipped: {type(e).__name__}: {e}")
        try:
            next_doc = getattr(doc, "GetNext", None)
            doc = next_doc() if callable(next_doc) else next_doc
        except Exception:
            break
    return docs


def _detect_preopened_dependencies(swApp, reference_diagnostics: dict, logger) -> dict:
    dep_paths = []
    seen_deps = set()
    for path_value in (reference_diagnostics or {}).get("path_entries", []) or []:
        text = str(path_value or "")
        if not text:
            continue
        lower = text.lower()
        if not lower.endswith((".sldprt", ".sldasm")):
            continue
        key = _norm_path(text)
        if key not in seen_deps:
            dep_paths.append(text)
            seen_deps.add(key)

    open_docs = _iter_open_documents(swApp, logger)
    by_path = {_norm_path(d.get("path", "")): d for d in open_docs if d.get("path")}
    by_title = {}
    for doc in open_docs:
        title = (doc.get("title") or os.path.basename(doc.get("path", "")) or "").lower()
        if title:
            by_title[title] = doc

    already_open = []
    closed = []
    for dep_path in dep_paths:
        dep_key = _norm_path(dep_path)
        dep_base = os.path.basename(dep_path).lower()
        matched = by_path.get(dep_key) or by_title.get(dep_base)
        if matched is None:
            try:
                opened = swApp.GetOpenDocumentByName(dep_path)
                if opened is None and dep_base:
                    opened = swApp.GetOpenDocumentByName(os.path.basename(dep_path))
                if opened is not None:
                    matched = {
                        "path": _safe_model_path(opened),
                        "title": _safe_model_title(opened),
                        "doc_type": _safe_doc_type(opened)[0],
                        "doc_type_name": _safe_doc_type(opened)[1],
                    }
            except Exception:
                pass
        if matched is not None:
            already_open.append({
                "dependency_path": dep_path,
                "open_doc_path": matched.get("path", ""),
                "open_doc_title": matched.get("title", ""),
                "open_doc_type": matched.get("doc_type_name", ""),
            })
        else:
            closed.append(dep_path)

    diagnostics = {
        "total_dependencies": len(dep_paths),
        "already_open_count": len(already_open),
        "closed_count": len(closed),
        "already_open": already_open,
        "closed_sample": closed[:25],
        "open_documents_total": len(open_docs),
        "open_documents_sample": open_docs[:25],
    }
    logger.info(
        f"[Extractor] Pre-open dependency inventory: total_dependencies={diagnostics['total_dependencies']} "
        f"already_open={diagnostics['already_open_count']} closed={diagnostics['closed_count']} "
        f"open_documents_in_session={diagnostics['open_documents_total']}"
    )
    for item in already_open[:30]:
        logger.info(
            f"[Extractor] Dependency already open: dependency='{item['dependency_path']}' "
            f"open_doc='{item['open_doc_path'] or item['open_doc_title']}' type={item['open_doc_type']}"
        )
    if len(already_open) > 30:
        logger.info(f"[Extractor] Dependency already-open list truncated: {len(already_open) - 30} more")
    return diagnostics


# ── Target properties for Layer 1 extraction ──────────────────────────────────
_TARGET_PROPERTIES = [
    "HYDRO_TEST_POSITION",
    "SHELL_IDP", "SHELL_MOT",
    "TUBE_IDP",  "TUBE_MOT",
    "JACKET_IDP", "JACKET_MOT",
    "Drawing_Number", "Tag_No", "Equipment_Type", "Equipment_Configuration",
    "Design_Code", "Material_Code", "Inspection_By",
    "DrawnBy", "DrawnDate", "CheckedBy", "CheckedDate",
    "EngineeringApproval", "EngAppDate",
    "Revision",
]


def _com_call(obj, method: str, *args):
    """
    Call a COM method or read a COM property uniformly.
    In late-bound mode some APIs (GetNames, GetSheetNames) are exposed as
    properties (callable=False) rather than methods — accessing them with ()
    raises 'tuple object is not callable'.  This helper handles both cases.
    """
    raw = getattr(obj, method)
    if callable(raw):
        return raw(*args)
    # Property access — args are ignored (properties have no call-time args)
    return raw


def _read_cpm(mgr, source_label: str, logger, probe_names=None) -> dict[str, str]:
    """
    Read all properties from a CustomPropertyManager; return name→value dict.

    useCached strategy
    ------------------
    SolidWorks custom property values can be literal strings or expressions like
    "$PRPWLD:Design_Code" that resolve by looking up a linked model property.
    When the drawing is opened without its referenced part (ExternalRefsNotLoaded),
    Get5(..., useCached=False) forces live re-evaluation → the expression cannot
    be resolved → returns "".

    Correct approach: try useCached=True first (returns the value that was
    cached/resolved at last save — the same value that displays in the title
    block and in the Property Tab).  Only fall back to useCached=False if True
    returns nothing, in case the property is a freshly-added literal with no
    prior cached state.

    probe_names fallback
    --------------------
    In late-bind COM, ICustomPropertyManager.GetNames() can return None/empty
    even when properties exist (COM SAFEARRAY not unwrapped). When this happens
    and probe_names is supplied, we directly call Get4/Get5 for each known name
    to bypass the enumeration failure.
    """
    def _pick_value(ret):
        """
        Extract the best non-empty string from a COM return value.
        In SolidWorks Get4/Get5/Get6 the tuple order is:
          (raw_expr, resolved_val, wasResolved[, linkToProp])
        We prefer resolved_val (idx 1) over raw_expr (idx 0), but skip
        $PRPWLD / $PRP expressions unless nothing else is available.
        """
        if ret is None:
            return ""
        if isinstance(ret, str):
            v = ret.strip()
            return "" if v.startswith("$") else v
        if isinstance(ret, (list, tuple)) and len(ret) > 0:
            # Try indices in preference order: resolved first (1), then raw (0)
            best = ""
            for idx in (1, 0):
                if len(ret) > idx:
                    v = str(ret[idx]).strip()
                    if v and not v.startswith("$"):
                        return v
                    if v and not best:
                        best = v     # keep raw expression as last resort
            return best
        return ""

    def _extract_one(mgr, name):
        """
        Read one SolidWorks custom property value.

        Diagnostic finding (Job 83): Get4/Get5/Get6 require ByRef (VT_BYREF) VARIANT
        objects for their out-params — plain Python strings cause "Type mismatch".
        Calling with only (name, useCached) gives "Parameter not optional".

        Strategy A (InvokeTypes): Use _oleobj_.InvokeTypes() which lets us declare
        each param's VT type + IN/OUT direction.  win32com then handles VT_BYREF
        marshaling correctly and returns the out-param values in the result tuple.
        This is the only reliable path for ByRef params in late-bind COM.

        Strategy B (GetAll3): Bulk array read — arrays survive SAFEARRAY late-bind
        better than scalar ByRef strings.

        Strategy C (no extra args, fallback): original approach; works if COM
        happens to return HRESULT-as-tuple or value-as-string.

        API param flags: PARAMFLAG_FIN=1, PARAMFLAG_FOUT=2
        Get6 → (name/in, useCached/in, val/out, resolvedVal/out, wasResolved/out, linkToProp/out)
        Get5 → (name/in, useCached/in, val/out, resolvedVal/out, wasResolved/out)
        Get4 → (name/in, useCached/in, val/out, resolvedVal/out)
        resolvedVal is the 2nd out-param; in the InvokeTypes result tuple it is at index 2
        (index 0 = HRESULT/retcode, index 1 = val, index 2 = resolvedVal).
        """
        import pythoncom
        FIN, FOUT = 1, 2
        VBS, VBL = pythoncom.VT_BSTR, pythoncom.VT_BOOL
        VI4 = pythoncom.VT_I4

        # Spec: (ret_type, arg_types_tuple, resolved_val_idx_in_result)
        _invoke_specs = {
            "Get6": ((VI4, 0),
                     ((VBS, FIN), (VBL, FIN), (VBS, FOUT), (VBS, FOUT), (VBL, FOUT), (VBL, FOUT)),
                     2),
            "Get5": ((VI4, 0),
                     ((VBS, FIN), (VBL, FIN), (VBS, FOUT), (VBS, FOUT), (VBL, FOUT)),
                     2),
            "Get4": ((VI4, 0),
                     ((VBS, FIN), (VBL, FIN), (VBS, FOUT), (VBS, FOUT)),
                     2),
        }

        # ── Strategy A: win32com VARIANT(VT_BYREF) ───────────────────────────
        # Pass explicit VT_BSTR|VT_BYREF VARIANT objects as the out-param slots.
        # win32com CDispatch sees them as pre-typed VARIANTs and passes them
        # through IDispatch without re-wrapping, so COM can write back through
        # the ByRef pointer.  After the call, variant.value holds the result.
        try:
            from win32com.client import VARIANT as _VARIANT
            _VT_BS_REF  = pythoncom.VT_BSTR | pythoncom.VT_BYREF
            _VT_BL_REF  = pythoncom.VT_BOOL | pythoncom.VT_BYREF
            for api_name, extra_ref_args in [
                ("Get6", [_VARIANT(_VT_BL_REF, False), _VARIANT(_VT_BL_REF, False)]),
                ("Get5", [_VARIANT(_VT_BL_REF, False)]),
                ("Get4", []),
            ]:
                for use_cached in (False, True):
                    try:
                        v_val  = _VARIANT(_VT_BS_REF, "")
                        v_rval = _VARIANT(_VT_BS_REF, "")
                        getattr(mgr, api_name)(name, use_cached, v_val, v_rval, *extra_ref_args)
                        raw  = (v_val.value  or "").strip()
                        res  = (v_rval.value or "").strip()
                        v = res if (res and not res.startswith("$")) else (
                            raw if (raw and not raw.startswith("$")) else "")
                        if v:
                            return v, f"VARIANT.{api_name}(uc={use_cached})"
                    except Exception:
                        continue
        except ImportError:
            pass

        # ── Strategy B: InvokeTypes (explicit VT typing + PARAMFLAG_FOUT) ─────
        # Note: GetIDsOfNames returns a plain int on SW2019 (not a tuple).
        for api_name, (ret_type, arg_types, rval_idx) in _invoke_specs.items():
            for use_cached in (False, True):
                try:
                    ids = mgr._oleobj_.GetIDsOfNames(0, api_name)
                    dispid = ids[0] if isinstance(ids, (list, tuple)) else int(ids)
                    result = mgr._oleobj_.InvokeTypes(
                        dispid, 0, 1,  # DISPATCH_METHOD
                        ret_type, arg_types,
                        name, use_cached  # only IN params go here
                    )
                    if isinstance(result, (list, tuple)) and len(result) > rval_idx:
                        rv = str(result[rval_idx]).strip()   # resolvedVal
                        ev = str(result[1]).strip() if len(result) > 1 else ""  # val/expr
                        v = rv if (rv and not rv.startswith("$")) else (
                            ev if (ev and not ev.startswith("$")) else "")
                        if v:
                            return v, f"InvokeTypes.{api_name}(uc={use_cached})"
                except Exception:
                    continue

        # ── Strategy C: bare 2-arg call (original) ───────────────────────────
        for api in ("Get6", "Get5", "Get4", "Get2"):
            for use_cached in (False, True):
                try:
                    ret = getattr(mgr, api)(name, use_cached)
                    v = _pick_value(ret)
                    if v:
                        return v, f"{api}(nobyref,uc={use_cached})"
                except Exception:
                    continue

        return "", ""

    try:
        names = _com_call(mgr, "GetNames")
        if not names:
            # Check Count to distinguish truly-empty vs SAFEARRAY unwrap failure
            count = None
            try:
                count = _com_call(mgr, "Count")
            except Exception:
                pass
            logger.info(
                f"[CP] {source_label}: CustomPropertyManager returned no names "
                f"(Count={count})"
            )
            # ── Fallback A: GetAll3 bulk read ────────────────────────────────
            # GetAll3 returns (names_array, types_array, values_array,
            # resolvedValues_array).  Array out-params survive late-bind better
            # than scalar ByRef strings.
            try:
                bulk = mgr.GetAll3()
                if isinstance(bulk, (list, tuple)) and len(bulk) >= 4:
                    b_names  = bulk[0] or ()
                    b_vals   = bulk[2] or ()
                    b_rvals  = bulk[3] or ()
                    if b_names:
                        logger.info(
                            f"[CP] {source_label}: GetAll3 found {len(b_names)} names"
                        )
                        result: dict[str, str] = {}
                        for i, n in enumerate(b_names):
                            rv = (b_rvals[i] if i < len(b_rvals) else "").strip()
                            ev = (b_vals[i]  if i < len(b_vals)  else "").strip()
                            val = rv if (rv and not rv.startswith("$")) else (
                                  ev if (ev and not ev.startswith("$")) else "")
                            result[str(n)] = val
                        found = sum(1 for v in result.values() if v)
                        logger.info(
                            f"[CP] {source_label}: GetAll3 → {len(result)} props "
                            f"({found} with values)"
                        )
                        return result
            except Exception as ge:
                logger.debug(f"[CP] {source_label}: GetAll3 fallback: {ge}")

            if not probe_names:
                return {}
            # ── Fallback B: direct probe of known target names ────────────────
            # Call Get6/Get4/Get2 with ByRef placeholders directly for each
            # target property name — bypasses SAFEARRAY enumeration entirely.
            logger.info(
                f"[CP] {source_label}: direct-probing {len(probe_names)} known property names"
            )
            result: dict[str, str] = {}
            for name in probe_names:
                value, winning_call = _extract_one(mgr, name)
                if value:
                    result[name] = value
                    logger.info(
                        f"[CP] {source_label}  {name!r} = {value!r}  via {winning_call}"
                    )
            found = sum(1 for v in result.values() if v)
            logger.info(
                f"[CP] {source_label}: direct-probe found {found}/{len(probe_names)} "
                f"properties with values"
            )
            return result

        result: dict[str, str] = {}
        for name in names:
            try:
                value, winning_call = _extract_one(mgr, name)
                result[name] = value
                logger.debug(
                    f"[CP] {source_label}  {name!r} = {value!r}"
                    + (f"  via {winning_call}" if winning_call else "  (empty)")
                )
            except Exception as ex:
                logger.debug(f"[CP] {source_label}: cannot read '{name}': {ex}")

        # Log summary with actual values for target properties
        found_count = sum(1 for v in result.values() if v)
        logger.info(
            f"[CP] {source_label}: {len(result)} properties detected "
            f"({found_count} with values): {list(result.keys())}"
        )
        for name, val in result.items():
            if val:
                logger.info(f"[CP] {source_label}  {name} = {val!r}")
            else:
                logger.debug(f"[CP] {source_label}  {name} = (empty)")
        return result
    except Exception as e:
        logger.warning(f"[CP] {source_label}: GetNames failed: {e}")
        return {}


def _extract_custom_properties(swApp, swModel, logger,
                               preopen_diags: dict | None = None,
                               temp_path: str = "") -> dict:
    """
    Extract target custom properties from three sources (priority order):
      1. Drawing-level  — CustomPropertyManager("")
      2. Sheet-level    — CustomPropertyManager(active_sheet_name)
      3. Model-level    — referenced part/assembly CustomPropertyManager("")

    Returns:
      {
        "bySource":   {"drawing": {...}, "sheet": {...}, "model": {...}},
        "resolved":   {"prop": {"value": str, "source": str}, ...},
        "allDetected": {"drawing": [...], "sheet": [...], "model": [...]},
        "totalFound": int,
      }
    """
    by_source: dict[str, dict] = {"drawing": {}, "sheet": {}, "model": {}}
    all_detected: dict[str, list] = {"drawing": [], "sheet": [], "model": []}

    # 1. Drawing-level
    try:
        mgr = swModel.Extension.CustomPropertyManager("")
        by_source["drawing"] = _read_cpm(mgr, "drawing", logger)
        all_detected["drawing"] = list(by_source["drawing"].keys())
        # Confirmed working strategy: VARIANT(VT_BYREF) via Strategy A.
        # GetIDsOfNames returns int (not tuple) on SW2019 — logged for reference only.
        import logging as _logging
        if logger.isEnabledFor(_logging.DEBUG):
            import pythoncom as _pc
            diag_props = [p for p in _TARGET_PROPERTIES if p in by_source["drawing"]][:1]
            for dp in diag_props:
                try:
                    from win32com.client import VARIANT as _V
                    _VBR = _pc.VT_BSTR | _pc.VT_BYREF
                    v_v, v_r = _V(_VBR, ""), _V(_VBR, "")
                    getattr(mgr, "Get6")(dp, False, v_v, v_r,
                        _V(_pc.VT_BOOL | _pc.VT_BYREF, False),
                        _V(_pc.VT_BOOL | _pc.VT_BYREF, False))
                    logger.debug(f"[CPRaw] VARIANT.Get6({dp!r}) → val={v_v.value!r} rval={v_r.value!r}")
                except Exception as e:
                    logger.debug(f"[CPRaw] VARIANT.Get6({dp!r}) → EXCEPTION: {e}")
    except Exception as e:
        logger.warning(f"[CP] Drawing-level CustomPropertyManager failed: {e}")

    # 2. Sheet-level
    try:
        sheet_name: str | None = None
        try:
            sheets = _com_call(swModel, "GetSheetNames")
            if sheets:
                sheet_name = str(sheets[0])
        except Exception:
            pass
        if sheet_name:
            mgr = swModel.Extension.CustomPropertyManager(sheet_name)
            by_source["sheet"] = _read_cpm(mgr, f"sheet({sheet_name})", logger)
            all_detected["sheet"] = list(by_source["sheet"].keys())
        else:
            logger.info("[CP] Sheet-level: no sheet name available")
    except Exception as e:
        logger.warning(f"[CP] Sheet-level failed: {e}")

    # 3. Model-level (first referenced part/assembly found in dependencies)
    #
    # SolidWorks COM ActivateDoc3 quirk:
    #   ActivateDoc3 quirk in late-bind mode:
    #     ActivateDoc3 uses the document *stem* as its key.  When a drawing
    #     (Part1.SLDDRW) and a part (Part1.SLDPRT) share the same stem, the
    #     DRAWING is returned because it is the currently active document.
    #     We must verify the returned doc type is Part(1) or Assembly(2),
    #     never Drawing(3).
    #
    # Primary strategy — navigate via drawing VIEWS:
    #     swModel.GetViews() (confirmed working in all job logs) returns
    #     sheet-view tuples.  view.ReferencedDocument directly returns the
    #     referenced Part/Assembly document — no path ambiguity, no stem clash.
    #
    # Fallback strategies:
    #   A — GetDocumentDependencies2 → ActivateDoc3(full, base, stem) with type check
    #   B — GetFirstDocument/GetNext chain with type check
    #   C — Try activating the dep path directly via GetDocumentDependencies2
    #       result, verifying type == 1 or 2
    try:
        model_doc = None

        def _disp(obj):
            """
            Wrap a raw COM pointer (PyIDispatch) in a win32com CDispatch so that
            named attribute/property access works.  Passes CDispatch through as-is.
            Returns None if obj is None or wrapping fails.
            """
            if obj is None:
                return None
            try:
                return win32com.client.Dispatch(obj)
            except Exception:
                return None

        def _is_part_or_asm(doc) -> bool:
            """Return True only if doc is a Part(1) or Assembly(2), not Drawing(3).

            In late-bind COM, GetType may be a non-callable property that returns
            an int directly.  doc.GetType() would then try to call that int →
            'int is not callable'.  _com_call handles both property and method
            forms by checking callable() first.
            """
            try:
                d = _disp(doc)
                if d is None:
                    return False
                t = _com_call(d, "GetType")
                logger.debug(f"[CP] _is_part_or_asm: GetType()={t!r}")
                return t in (1, 2)
            except Exception as e:
                logger.debug(f"[CP] _is_part_or_asm GetType failed: {e}")
                return False

        def _safe_path(doc) -> str:
            try:
                d = _disp(doc)
                if d is None:
                    return "(no doc)"
                return _com_call(d, "GetPathName") or "(no path)"
            except Exception:
                return "(unknown)"

        def _try_activate_part(name: str) -> object | None:
            """Activate by name; return dispatched doc only if it is Part/Assembly."""
            try:
                err_v = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                raw = swApp.ActivateDoc3(name, False, 0, err_v)
                doc = _disp(raw)
                if doc is not None and _is_part_or_asm(doc):
                    return doc
                if doc is not None:
                    logger.debug(f"[CP] _try_activate_part({name!r}): got type={doc.GetType()} — not Part/Asm, skipping")
                return None
            except Exception as e:
                logger.debug(f"[CP] _try_activate_part({name!r}): {e}")
                return None

        # ─── Pass 1 ──────────────────────────────────────────────────────────
        # Traverse drawing views → view.ReferencedDocument
        # GetViews returns ((<PyIDispatch>,),) — wrap each view with _disp().
        try:
            views_result = _com_call(swModel, "GetViews")
            logger.info(f"[CP] Pass1 GetViews raw type={type(views_result).__name__} len={len(views_result) if views_result else 0}")
            if views_result and isinstance(views_result, (list, tuple)):
                for si, sheet_entry in enumerate(views_result):
                    if sheet_entry is None:
                        continue
                    view_list = sheet_entry if isinstance(sheet_entry, (list, tuple)) else [sheet_entry]
                    logger.info(f"[CP] Pass1 sheet[{si}]: {len(view_list)} view(s)")
                    for vi, raw_view in enumerate(view_list):
                        if raw_view is None:
                            continue
                        view = _disp(raw_view)
                        if view is None:
                            logger.warning(f"[CP] Pass1 sheet[{si}] view[{vi}]: _disp() returned None")
                            continue
                        try:
                            ref_raw  = view.ReferencedDocument
                            ref_doc  = _disp(ref_raw)
                            logger.info(f"[CP] Pass1 view[{vi}].ReferencedDocument → {type(ref_raw).__name__}")
                            if ref_doc is not None and _is_part_or_asm(ref_doc):
                                model_doc = ref_doc
                                logger.info(f"[CP] Model-level Pass1 ✓ view.ReferencedDocument → {_safe_path(model_doc)!r}")
                                break
                            elif ref_doc is not None:
                                try:
                                    t = ref_doc.GetType()
                                except Exception:
                                    t = "?"
                                logger.warning(f"[CP] Pass1 view[{vi}].ReferencedDocument type={t} — not Part/Asm")
                            else:
                                logger.warning(f"[CP] Pass1 view[{vi}].ReferencedDocument is None")
                        except Exception as ve:
                            logger.warning(f"[CP] Pass1 view[{vi}].ReferencedDocument: {ve}")
                    if model_doc is not None:
                        break
        except Exception as vp_e:
            logger.warning(f"[CP] Pass1 GetViews traversal failed: {vp_e}")

        # ─── Pass 2 ──────────────────────────────────────────────────────────
        # GetDocumentDependencies2(temp_path) → GetOpenDocumentByName per dep.
        # Uses temp_path directly to avoid swModel.GetPathName() which fails in
        # late-bind (it's a property, not a method — calling it() raises
        # 'str object is not callable').  Also uses GetOpenDocumentByName
        # instead of ActivateDoc3 because the latter returns None even for
        # documents that are already loaded.
        if model_doc is None and temp_path:
            try:
                deps = swApp.GetDocumentDependencies2(temp_path, False, True, False) or []
                logger.info(f"[CP] Pass2 GetDocumentDependencies2({temp_path!r}): {len(deps)} entries")
                for dep in deps:
                    dep_str = str(dep or "")
                    if not dep_str.lower().endswith((".sldprt", ".sldasm")):
                        continue
                    basename = os.path.basename(dep_str)
                    logger.info(f"[CP] Pass2 dep={dep_str!r} → GetOpenDocumentByName(full, base)")
                    for try_name in (dep_str, basename):
                        try:
                            raw = swApp.GetOpenDocumentByName(try_name)
                            doc = _disp(raw)
                            if doc is not None and _is_part_or_asm(doc):
                                model_doc = doc
                                logger.info(f"[CP] Model-level Pass2 ✓ GetOpenDocumentByName({try_name!r}) → Part/Asm")
                                break
                            elif raw is not None:
                                logger.warning(f"[CP] Pass2 GetOpenDocumentByName({try_name!r}) → {type(raw).__name__} (not Part/Asm)")
                            else:
                                logger.info(f"[CP] Pass2 GetOpenDocumentByName({try_name!r}) → None")
                        except Exception as e:
                            logger.warning(f"[CP] Pass2 GetOpenDocumentByName({try_name!r}): {e}")
                    if model_doc is not None:
                        break
                if model_doc is None:
                    logger.warning("[CP] Pass2: no Part/Asm found via GetOpenDocumentByName")
            except Exception as dep_e:
                logger.warning(f"[CP] Pass2 GetDocumentDependencies2: {dep_e}")

        # ─── Pass 3 ──────────────────────────────────────────────────────────
        # Walk GetFirstDocument → GetNext, pick first Part/Asm.
        # Documents are PyIDispatch in late-bind; must _disp() each.
        if model_doc is None:
            try:
                raw_first = _com_call(swApp, "GetFirstDocument")
                doc = _disp(raw_first)
                logger.info(f"[CP] Pass3 GetFirstDocument → {type(raw_first).__name__}")
                seen: set[int] = set()
                iters = 0
                while doc is not None and iters < 20:
                    iters += 1
                    doc_id = id(doc)
                    if doc_id in seen:
                        break
                    seen.add(doc_id)
                    try:
                        t = doc.GetType()
                        p = _safe_path(doc)
                        logger.info(f"[CP] Pass3 doc[{iters}] type={t} path={p!r}")
                        if t in (1, 2):
                            model_doc = doc
                            logger.info(f"[CP] Model-level Pass3 ✓ open-doc scan → {p!r}")
                            break
                    except Exception as te:
                        logger.warning(f"[CP] Pass3 doc[{iters}].GetType: {te}")
                    try:
                        raw_next = _com_call(doc, "GetNext")
                        doc = _disp(raw_next)
                    except Exception:
                        break
                if model_doc is None:
                    logger.warning(f"[CP] Pass3: scanned {iters} doc(s), no Part/Asm found")
            except Exception as scan_e:
                logger.warning(f"[CP] Pass3 open-doc scan: {scan_e}")

        # ─── Pass 4 ──────────────────────────────────────────────────────────
        # GetOpenDocumentByName with paths from the pre-open dependency scan.
        # This API is confirmed working BEFORE the drawing opens; try it again
        # here in case it still resolves the already-open part after opening
        # the drawing.  This avoids all ReferencedDocument / GetFirstDocument
        # COM issues entirely.
        if model_doc is None:
            open_items = (preopen_diags or {}).get("already_open", []) or []
            logger.info(f"[CP] Pass4: checking {len(open_items)} pre-open dependency path(s) via GetOpenDocumentByName")
            for item in open_items:
                dep_path = item.get("dependency_path", "")
                if not dep_path:
                    continue
                if not dep_path.lower().endswith((".sldprt", ".sldasm")):
                    continue
                for try_name in (dep_path, os.path.basename(dep_path)):
                    try:
                        raw = swApp.GetOpenDocumentByName(try_name)
                        doc = _disp(raw)
                        if doc is not None and _is_part_or_asm(doc):
                            model_doc = doc
                            logger.info(f"[CP] Model-level Pass4 ✓ GetOpenDocumentByName({try_name!r}) → Part/Asm")
                            break
                        elif raw is not None:
                            logger.warning(f"[CP] Pass4 GetOpenDocumentByName({try_name!r}) → {type(raw).__name__} (not Part/Asm)")
                        else:
                            logger.info(f"[CP] Pass4 GetOpenDocumentByName({try_name!r}) → None")
                    except Exception as e:
                        logger.warning(f"[CP] Pass4 GetOpenDocumentByName({try_name!r}): {e}")
                if model_doc is not None:
                    break
            if model_doc is None and open_items:
                logger.warning("[CP] Pass4: GetOpenDocumentByName returned None for all known part paths")

        if model_doc is not None:
            # Read document-level CPM ("") first
            try:
                mgr = model_doc.Extension.CustomPropertyManager("")
                doc_props = _read_cpm(mgr, "model/doc", logger)
            except Exception as e:
                logger.warning(f"[CP] model CustomPropertyManager(''): {e}")
                doc_props = {}

            # Read every configuration-level CPM — $PRPWLD resolves from the
            # active configuration, so the real property values live in
            # CustomPropertyManager("Default") or similar, not in ("").
            config_props: dict[str, str] = {}
            try:
                cfg_names = _com_call(model_doc, "GetConfigurationNames") or ()
                if isinstance(cfg_names, str):
                    cfg_names = (cfg_names,)
                logger.info(f"[CP] model configurations: {list(cfg_names)}")
                for cfg in cfg_names:
                    try:
                        cmgr = model_doc.Extension.CustomPropertyManager(str(cfg))
                        cfg_result = _read_cpm(
                            cmgr, f"model/cfg({cfg})", logger,
                            probe_names=_TARGET_PROPERTIES
                        )
                        # Merge: first non-empty value per property wins
                        for k, v in cfg_result.items():
                            if v and not config_props.get(k):
                                config_props[k] = v
                    except Exception as ce:
                        logger.warning(f"[CP] model CPM({cfg!r}): {ce}")
            except Exception as ge:
                logger.warning(f"[CP] model GetConfigurationNames: {ge}")

            # Merge: config wins over doc-level for same key (doc-level is usually
            # the expression link with an empty resolved value)
            merged: dict[str, str] = {**doc_props}
            for k, v in config_props.items():
                if v:  # config value wins if non-empty
                    merged[k] = v
                elif k not in merged:
                    merged[k] = v

            by_source["model"] = merged
            all_detected["model"] = list(merged.keys())
        else:
            logger.info("[CP] Model-level: no referenced model document accessible (all passes failed)")
    except Exception as e:
        logger.warning(f"[CP] Model-level failed: {e}")

    # Resolve priority: drawing > sheet > model
    resolved: dict[str, dict] = {}
    for prop in _TARGET_PROPERTIES:
        found = False
        for src in ("drawing", "sheet", "model"):
            val = by_source[src].get(prop)
            if val is not None and val != "":
                resolved[prop] = {"value": val, "source": src}
                found = True
                break
        if not found:
            resolved[prop] = {"value": "", "source": "none"}

    total_found = sum(1 for v in resolved.values() if v["value"])
    logger.info(f"[CP] Resolution complete: {total_found}/{len(_TARGET_PROPERTIES)} target properties found")
    for prop, info in resolved.items():
        if info["value"]:
            logger.info(f"[CP]   {prop} = {info['value']!r}  (source={info['source']})")
        else:
            logger.info(f"[CP]   {prop} = MISSING")

    return {
        "bySource":    by_source,
        "resolved":    resolved,
        "allDetected": all_detected,
        "totalFound":  total_found,
    }


def _set_doc_spec_attr(docSpec, names: tuple[str, ...], value, logger) -> str:
    for name in names:
        try:
            setattr(docSpec, name, value)
            try:
                actual = getattr(docSpec, name)
            except Exception:
                actual = value
            logger.info(f"[Extractor] OpenDoc7 docSpec {name}={actual}")
            return name
        except Exception:
            continue
    logger.info(f"[Extractor] OpenDoc7 docSpec skipped attrs={names} unsupported")
    return ""


def _activate_and_refetch(swApp, swModel, temp_path: str, logger, label: str):
    candidates = []
    activated_candidates = []
    doc_names = []
    if temp_path:
        doc_names = [temp_path, os.path.basename(temp_path), os.path.splitext(os.path.basename(temp_path))[0]]
    for doc_name in doc_names:
        try:
            errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            activated = swApp.ActivateDoc3(doc_name, False, 0, errors)
            logger.info(f"[COMDBG] ActivateDoc3 {label} name='{doc_name}' returned {type(activated).__name__ if activated else 'None'} errors={getattr(errors, 'value', '')}")
            if activated is not None:
                candidates.append(activated)
                activated_candidates.append((doc_name, activated))
        except Exception as e:
            logger.info(f"[COMDBG] ActivateDoc3 {label} name='{doc_name}' failed: {type(e).__name__}: {e}")
        try:
            opened = swApp.GetOpenDocumentByName(doc_name)
            logger.info(f"[COMDBG] GetOpenDocumentByName {label} name='{doc_name}' returned {type(opened).__name__ if opened else 'None'}")
            if opened is not None:
                candidates.append(opened)
        except Exception as e:
            logger.info(f"[COMDBG] GetOpenDocumentByName {label} name='{doc_name}' failed: {type(e).__name__}: {e}")
    active = get_active_doc(swApp)
    if active is not None:
        logger.info(f"[COMDBG] ActiveDoc {label} returned {type(active).__name__}")
        candidates.append(active)
    candidates.append(swModel)
    for candidate in candidates:
        if candidate is None:
            continue
        logger.info(f"[COMDBG] {label} candidate before drawing cast: {com_type_summary(candidate)}")
        draw = cast_to_drawing_doc(candidate)
        logger.info(f"[COMDBG] {label} candidate after drawing cast: {com_type_summary(draw)}")
        first_view = probe_method(draw, "GetFirstView")
        current_sheet = probe_method(draw, "GetCurrentSheet")
        if first_view["call_ok"] or current_sheet["call_ok"]:
            return candidate, draw
    for doc_name, candidate in activated_candidates:
        if candidate is None:
            continue
        draw = cast_to_drawing_doc(candidate)
        logger.info(
            f"[COMDBG] {label}: accepting ActivateDoc3 document for '{doc_name}' "
            "without drawing API probe success; SolidWorks may expose drawing APIs only after extraction starts"
        )
        return candidate, draw
    return swModel, cast_to_drawing_doc(swModel)


def _log_com_debug(swApp, swModel, swDraw, logger, label: str):
    for obj_label, obj in [("swModel", swModel), ("swDraw", swDraw)]:
        summary = com_type_summary(obj)
        logger.info(
            f"[COMDBG] {label} {obj_label}: "
            f"python={summary['python_module']}.{summary['python_type']} "
            f"typeinfo='{summary['typeinfo_name']}' guid='{summary['typeattr_guid']}' repr='{summary['repr']}'"
        )
        for method in ("GetFirstView", "GetCurrentSheet", "GetViews", "GetFirstAnnotation", "GetSheetNames", "GetDimensionNames"):
            probe = probe_method(obj, method)
            logger.info(
                f"[COMDBG] {label} {obj_label}.{method}: "
                f"has={probe['has_attr']} callable={probe['callable']} ok={probe['call_ok']} "
                f"result={probe['result_type']} preview='{probe['result_preview']}' error='{probe['error']}'"
            )
    try:
        sheet_names = sw_call(swDraw, "GetSheetNames")
        sheet_list = list(sheet_names) if hasattr(sheet_names, "__iter__") and not isinstance(sheet_names, str) else [sheet_names]
        if sheet_list:
            try:
                sw_call(swDraw, "ActivateSheet", sheet_list[0])
                _, swDraw2 = _activate_and_refetch(swApp, swModel, "", logger, f"{label}/after-sheet-activation")
                sheet = sw_call(swDraw2, "GetCurrentSheet")
                logger.info(f"[COMDBG] {label} current sheet object after activation: {com_type_summary(sheet)}")
                probe = probe_method(sheet, "GetViews")
                logger.info(
                    f"[COMDBG] {label} currentSheet.GetViews: has={probe['has_attr']} callable={probe['callable']} "
                    f"ok={probe['call_ok']} result={probe['result_type']} preview='{probe['result_preview']}' error='{probe['error']}'"
                )
            except Exception as e:
                logger.info(f"[COMDBG] {label} sheet activation probe failed: {type(e).__name__}: {e}")
    except Exception as e:
        logger.info(f"[COMDBG] {label} sheet probe skipped: {type(e).__name__}: {e}")




def _get_user_sw_search_paths(progid: str, logger) -> dict:
    """
    Read-only: query the user's already-running SolidWorks session for its
    configured file-search folders, so the agent's dedicated instance can
    inherit the same paths and find referenced parts/assemblies.

    swSearchFolderTypes_e: 1=Parts, 2=Assemblies, 4=Drawings, 7=ReferencedDocuments
    Safe — no changes to user session; pure read.
    """
    inherited: dict = {}
    # Try versioned ProgID first, then base ProgID
    for pid in [progid, "SldWorks.Application"]:
        try:
            existing_sw = win32com.client.GetActiveObject(pid)
            for folder_type in [1, 2, 4, 7]:
                try:
                    paths = existing_sw.GetSearchFolders(folder_type)
                    if paths:
                        inherited[folder_type] = paths
                except Exception:
                    pass
            if inherited:
                logger.info(
                    f"[Extractor] Inherited {len(inherited)} search-path type(s) "
                    f"from running SW session ({pid})"
                )
            break
        except Exception:
            pass  # no SW session running — normal if user has SW closed
    return inherited


def _auto_dismiss_sw_dialogs(stop_event: threading.Event, logger) -> None:
    """
    Background thread — polls for SolidWorks 'referenced file not found' dialogs
    and clicks the dismissal button (No / Cancel / Skip / Ignore / Don't Search).

    This runs alongside an OpenDoc6 call that has NO Silent flag, so SW can show
    its dialogs; we auto-dismiss them and the drawing opens in full mode with
    degraded (empty) views but full table/annotation API access.
    """
    if not WIN32GUI_AVAILABLE:
        return

    # Buttons that dismiss the "can't find file" prompt without searching
    _DISMISS = frozenset([
        "no", "cancel", "skip", "ignore",
        "don't search", "do not search", "suppress",
    ])

    def _click_dismiss(child_hwnd, _parent):
        try:
            txt = win32gui.GetWindowText(child_hwnd).strip().lower()
            if txt in _DISMISS:
                win32gui.PostMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
        except Exception:
            pass

    def _check_window(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        cls  = win32gui.GetClassName(hwnd)
        if cls not in ('#32770', 'MsoCommandBar', 'SWFormsClass',
                       'ThunderRT6Form', 'SldWorks'):
            return
        title = win32gui.GetWindowText(hwnd)
        # Match any SW dialog that may be blocking on missing files
        kws = ('solidworks', 'not found', 'missing', 'reference',
               'cannot', 'locate', 'resolve', 'file')
        if any(k in title.lower() for k in kws) or title == '':
            try:
                win32gui.EnumChildWindows(hwnd, _click_dismiss, None)
            except Exception:
                pass

    while not stop_event.is_set():
        try:
            win32gui.EnumWindows(_check_window, None)
        except Exception:
            pass
        time.sleep(0.08)


def run_extraction(temp_path: str, config, cancel_event: threading.Event,
                   logger) -> dict:
    """
    Main entry point called by job_runner in a worker thread.
    Returns the full extraction result dict (without agent metadata — runner stamps that).
    Raises on SolidWorks launch/open failures. Extraction modules are best-effort.
    """
    if not PYWIN32_AVAILABLE:
        raise RuntimeError(
            "pywin32 not available — this agent must run on Windows with pywin32 installed."
        )

    import hashlib
    file_size = os.path.getsize(temp_path)
    sha256    = _sha256(temp_path)
    filename  = os.path.basename(temp_path)

    logger.info(f"[Extractor] Starting: file={filename} size={file_size:,} bytes")
    logger.info(f"[Extractor] Using ProgID: {config.sw_progid}")

    result = {
        "schema_version":             "1.1",
        "agent":                      {},   # stamped by runner
        "file": {
            "original_filename": filename,
            "file_size_bytes":   file_size,
            "sha256":            sha256,
        },
        "customPropertyVerification": {},
        "extraction_warnings":        [],
        "extraction_errors":          {},
    }

    swApp  = None
    swModel = None

    try:
        # ── COM initialisation (must be called in each thread) ─────────────────
        pythoncom.CoInitialize()

        # ── Connect to SolidWorks ──────────────────────────────────────────────
        _check_cancel(cancel_event, "before SW launch")
        logger.info(f"[Extractor] Launching SolidWorks ({config.sw_progid})…")
        t_launch = time.monotonic()
        # ── Inherit search paths from user's running SW session (read-only) ───
        inherited_paths = _get_user_sw_search_paths(config.sw_progid, logger)

        logger.info("[Extractor] Connecting to SolidWorks COM…")
        swApp, binding_mode, attached_existing_session = _connect_sw_application(config.sw_progid, logger)
        # Force full COM initialisation so IDrawingDoc interface is fully loaded.
        # Visible=True + UserControl=True + a brief delay ensures SolidWorks has
        # finished its own COM registration before we attempt OpenDoc.
        swApp.Visible = True
        try:
            swApp.UserControl = True
        except Exception:
            pass
        swApp.UserControlBackground = True
        logger.info("[Extractor] Waiting 2.5 s for SolidWorks COM to fully initialise…")
        time.sleep(2.5)
        logger.info(f"[Extractor] SolidWorks ready ({time.monotonic() - t_launch:.1f}s)")

        # ── Open document (read-only, silent) ─────────────────────────────────
        _check_cancel(cancel_event, "before OpenDoc6")
        logger.info(f"[Extractor] Opening: {temp_path}")

        # Pre-configure SW to suppress missing-reference prompts at app level
        for pref_id, pref_val in [
            (57,  0),   # swFileMissingReferenceBehavior — 0=ignore, 1=use last paths
            (176, 0),   # swReferencedDocumentsMissingBehavior — 0=don't open
            (177, 0),   # related missing-ref preference
        ]:
            try:
                swApp.SetUserPreferenceIntegerValue(pref_id, pref_val)
            except Exception:
                pass

        # Apply inherited search paths from user's SW session
        for folder_type, paths in inherited_paths.items():
            try:
                swApp.SetSearchFolders(folder_type, paths)
            except Exception:
                pass

        # Also apply any manually configured model_search_path from config.ini
        if getattr(config, "sw_model_search_path", ""):
            raw = config.sw_model_search_path
            folder_list = "\r\n".join(
                p.strip() for p in raw.replace(";", "\n").splitlines() if p.strip()
            )
            for folder_type in [1, 2, 7]:
                try:
                    swApp.SetSearchFolders(folder_type, folder_list)
                except Exception:
                    pass
            logger.info(f"[Extractor] SW model search path applied: {raw}")

        swModel  = None
        pass_num = 0
        err_val  = 0
        warn_val = 0

        reference_diagnostics = _log_reference_diagnostics(swApp, temp_path, logger, "before-open")
        preopen_diagnostics = _detect_preopened_dependencies(swApp, reference_diagnostics, logger)

        open_diagnostics = {
            "file_size_bytes": file_size,
            "open_pass": "",
            "open_mode": "",
            "solidworks_session": "attached_existing" if attached_existing_session else "created_or_reused_by_com",
            "reference_diagnostics": reference_diagnostics,
            "preopened_dependencies": preopen_diagnostics,
            "open_mode_influence": {
                "preopened_dependencies_present": preopen_diagnostics["already_open_count"] > 0,
                "candidate_full_resolved_open_first": True,
                "open_sequence_changed_before_open": False,
                "reason": "Full resolved open is always attempted first; this test records whether already-open references make that pass succeed before fallbacks.",
            },
        }

        def _close_zombie(label: str):
            """CloseDoc after a failed open attempt to clear SW's internal open-state registry.
            Without this, SW returns AlreadyOpen (65536) for every subsequent attempt on the
            same path, even when OpenDoc returned None (zombie registration)."""
            try:
                swApp.CloseDoc(temp_path)
                logger.info(f"[Extractor] CloseDoc cleanup after failed {label}")
            except Exception:
                pass  # expected if nothing was actually registered

        def _try_refetch(label: str):
            nonlocal swModel
            if swModel is not None:
                return True
            candidate, draw = _activate_and_refetch(swApp, swModel, temp_path, logger, label)
            doc_type, _ = _safe_doc_type(candidate)
            if candidate is not None and doc_type == SW_DOC_DRAWING:
                logger.info(f"[Extractor] {label}: Open returned None but active document refetch found drawing")
                swModel = candidate
                return True
            try:
                first_view = probe_method(draw, "GetFirstView")
                current_sheet = probe_method(draw, "GetCurrentSheet")
                if candidate is not None and (first_view["call_ok"] or current_sheet["call_ok"]):
                    logger.info(f"[Extractor] {label}: Open returned None but active document refetch found drawing API")
                    swModel = candidate
                    return True
            except Exception:
                pass
            return False

        def _open_doc7(label: str, attr_sets: list[tuple[tuple[str, ...], object]], mode: str, read_only: bool = True):
            nonlocal swModel, err_val, warn_val, pass_num
            try:
                docSpec = swApp.GetOpenDocSpec(temp_path)
                docSpec.FileName     = temp_path
                docSpec.DocumentType = SW_DOC_DRAWING
                docSpec.ReadOnly     = read_only
                docSpec.Silent       = True
                logger.info("[Extractor] OpenDoc7 docSpec FileName set")
                logger.info("[Extractor] OpenDoc7 docSpec DocumentType=3 (Drawing)")
                logger.info(f"[Extractor] OpenDoc7 docSpec ReadOnly={read_only}")
                logger.info("[Extractor] OpenDoc7 docSpec Silent=True")
                for names, value in attr_sets:
                    _set_doc_spec_attr(docSpec, names, value, logger)
                swModel = swApp.OpenDoc7(docSpec)
                err_val  = docSpec.Error
                warn_val = docSpec.Warning
                _log_open_attempt(logger, "OpenDoc7", label, SW_DOC_DRAWING, None, mode, swModel, err_val, warn_val)
                pass_num = label
                open_diagnostics["open_mode"] = mode
            except Exception as e:
                logger.info(f"[Extractor] OpenDoc7 {label} not available/failed: {type(e).__name__}: {e}")
                swModel = None
            if _try_refetch(f"OpenDoc7 {label} active-refetch"):
                return True
            _close_zombie(f"OpenDoc7 {label}")
            return False

        open_doc7_passes = [
            ("pass 0 full-silent", [], "docSpec full drawing; references may resolve per SW defaults/search paths", True),
            ("pass L lightweight-default", [
                (("UseLightWeightDefault", "UseLightweightDefault", "LightWeight", "Lightweight"), True),
                (("LoadModel", "LoadExternalReferences", "LoadExternalRefs"), False),
                (("OpenDocOptions", "Options"), SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_OVERRIDE_DEFAULT | SW_OPEN_LOAD_LIGHTWEIGHT),
            ], "lightweight/deferred where supported; referenced model load not forced", True),
            ("pass R detailing-readonly", [
                (("DetailingMode", "RapidDraft"), True),
                (("ViewOnly", "LargeDesignReview"), False),
                (("LoadModel", "LoadExternalReferences", "LoadExternalRefs"), False),
                (("OpenDocOptions", "Options"), SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_RAPID_DRAFT),
            ], "detailing/rapid-draft read-only; drawing annotations/dimensions should remain available without resolving models", True),
            ("pass R2 detailing-editable-temp", [
                (("DetailingMode", "RapidDraft"), True),
                (("ViewOnly", "LargeDesignReview"), False),
                (("LoadModel", "LoadExternalReferences", "LoadExternalRefs"), False),
                (("OpenDocOptions", "Options"), SW_OPEN_SILENT | SW_OPEN_RAPID_DRAFT),
            ], "detailing/rapid-draft editable temp copy; deeper drawing APIs without saving or loading heavyweight models", False),
            ("pass V view-only", [
                (("ViewOnly", "LargeDesignReview"), True),
                (("DetailingMode", "RapidDraft"), False),
                (("LoadModel", "LoadExternalReferences", "LoadExternalRefs"), False),
                (("OpenDocOptions", "Options"), SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_VIEW_ONLY),
            ], "view-only/large-design-review where supported; referenced 3-D model load avoided", True),
        ]
        for label, attr_sets, mode, read_only in open_doc7_passes:
            if _open_doc7(label, attr_sets, mode, read_only):
                break

        if swModel is None:
            open_doc6_passes = [
                ("pass 1 full-silent", SW_OPEN_READ_ONLY | SW_OPEN_SILENT),
                ("pass 2 override-load-lightweight", SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_OVERRIDE_DEFAULT | SW_OPEN_LOAD_LIGHTWEIGHT),
                ("pass 3 detailing-rapid-draft", SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_RAPID_DRAFT),
                ("pass 4 detailing-rapid-draft-editable-temp", SW_OPEN_SILENT | SW_OPEN_RAPID_DRAFT),
                ("pass 5 view-only", SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_VIEW_ONLY),
                ("pass 6 detached-load-model-last-resort", SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_LOAD_MODEL),
            ]
            for label, options in open_doc6_passes:
                errors   = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                swModel  = swApp.OpenDoc6(
                    temp_path, SW_DOC_DRAWING, options, "", errors, warnings)
                err_val  = errors.value
                warn_val = warnings.value
                pass_num = label
                open_diagnostics["open_mode"] = _reference_load_mode(options)
                _log_open_attempt(logger, "OpenDoc6", label, SW_DOC_DRAWING, options, "", swModel, err_val, warn_val)
                if swModel is not None or _try_refetch(f"OpenDoc6 {label} active-refetch"):
                    break
                _close_zombie(label)
                if label == "pass 1 full-silent" and WIN32GUI_AVAILABLE:
                    logger.info("[Extractor] OpenDoc6 pass D dialog-dismiss: doc_type=3 (Drawing) open_options=2 (ReadOnly) solidworks_open_mode='interactive full drawing; missing-reference dialogs auto-dismissed'")
                    stop_dismiss = threading.Event()
                    dismisser    = threading.Thread(
                        target=_auto_dismiss_sw_dialogs,
                        args=(stop_dismiss, logger),
                        daemon=True,
                    )
                    dismisser.start()
                    try:
                        d_errors   = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                        d_warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                        d_model    = swApp.OpenDoc6(
                            temp_path, SW_DOC_DRAWING, SW_OPEN_READ_ONLY, "", d_errors, d_warnings)
                        d_err  = d_errors.value
                        d_warn = d_warnings.value
                        _log_open_attempt(logger, "OpenDoc6", "pass D dialog-dismiss", SW_DOC_DRAWING, SW_OPEN_READ_ONLY, "", d_model, d_err, d_warn)
                        if d_model is not None:
                            swModel  = d_model
                            err_val  = d_err
                            warn_val = d_warn
                            pass_num = "pass D dialog-dismiss"
                        if swModel is not None or _try_refetch("OpenDoc6 pass D active-refetch"):
                            break
                        _close_zombie("pass D")
                    except Exception as ex_d:
                        logger.info(f"[Extractor] OpenDoc6 pass D exception: {ex_d}")
                        _close_zombie("pass D (exception)")
                    finally:
                        stop_dismiss.set()
                        dismisser.join(timeout=2)

        if swModel is None:
            open_diagnostics["reference_diagnostics_after_failure"] = _log_reference_diagnostics(swApp, temp_path, logger, "after-open-failure")
            raise RuntimeError(
                f"All open passes failed for {filename}. "
                f"Last errors={_decode_sw_error(err_val)} warnings={_decode_sw_error(warn_val)}"
            )

        logger.info(f"[Extractor] Document open OK (pass={pass_num})")
        open_diagnostics["open_pass"] = str(pass_num)
        open_diagnostics["open_mode_influence"]["full_resolved_open_succeeded"] = str(pass_num) == "pass 0 full-silent"
        open_diagnostics["open_mode_influence"]["observed_richer_open_with_preopened_dependencies"] = (
            preopen_diagnostics["already_open_count"] > 0 and str(pass_num) == "pass 0 full-silent"
        )
        logger.info(
            f"[Extractor] Pre-open influence after drawing open: preopened_dependencies={preopen_diagnostics['already_open_count']} "
            f"open_pass='{pass_num}' full_resolved_open_succeeded={str(pass_num) == 'pass 0 full-silent'}"
        )
        result["open_diagnostics"] = open_diagnostics

        # ── Verify document type ───────────────────────────────────────────────
        # swDocumentTypes_e: 1=Part, 2=Assembly, 3=Drawing
        # In late-binding mode GetType may be a property (int) not a method —
        # try both access patterns so we never crash here.
        try:
            raw = swModel.GetType
            doc_type = raw() if callable(raw) else int(raw)
            type_name = {1: "Part", 2: "Assembly", 3: "Drawing"}.get(doc_type, f"Unknown({doc_type})")
            logger.info(f"[Extractor] Model type: {doc_type} ({type_name})")
            if doc_type != 3:
                raise RuntimeError(f"Expected swDocDRAWING (3) but got type {doc_type}. Wrong file type.")
            logger.info("[Extractor] Drawing detected")
        except RuntimeError:
            raise
        except Exception as e:
            logger.warning(f"[Extractor] GetType() skipped: {e} — assuming Drawing and continuing")

        swModel, swDraw = _activate_and_refetch(swApp, swModel, temp_path, logger, "after-open")
        _log_com_debug(swApp, swModel, swDraw, logger, "after-open")

        # ── Layer 1: Custom property extraction + verification ─────────────────
        _check_cancel(cancel_event, "before custom property extraction")
        try:
            cp_extraction = _extract_custom_properties(swApp, swModel, logger, preopen_diagnostics, temp_path)
            cp_result = verify_custom_properties(cp_extraction, logger)
            result.update(cp_result)

            # ── Populate `customProperties` — flat field list ─────────────
            resolved_map = cp_extraction.get("resolved", {})
            total_target = len(_TARGET_PROPERTIES)
            total_found  = cp_extraction.get("totalFound", 0)
            cp_fields = []
            for prop in _TARGET_PROPERTIES:
                info   = resolved_map.get(prop, {"value": "", "source": "none"})
                val    = info.get("value", "")
                source = info.get("source", "none")
                found  = bool(val) and source != "none"
                cp_fields.append({
                    "property":      prop,
                    "value":         val,
                    "resolvedValue": val,
                    "source":        source,
                    "found":         found,
                })
            result["customProperties"] = {
                "fields":           cp_fields,
                "foundCount":       total_found,
                "missingCount":     total_target - total_found,
                "totalTargetCount": total_target,
            }

            cp_status = result.get("customPropertyVerification", {}).get("status", "unknown")
            eq_cfg    = result.get("customPropertyVerification", {}).get("equipmentConfig", "")
            logger.info(
                f"[Extractor] customPropertyVerification: status={cp_status} equipmentConfig={eq_cfg!r}"
            )
        except Exception as e:
            err_msg = f"{type(e).__name__}: {e}"
            logger.error(f"[Extractor] customPropertyVerification SOFT FAIL: {err_msg}")
            result["extraction_errors"]["customPropertyVerification"] = err_msg
            result["extraction_warnings"].append(f"customPropertyVerification failed: {err_msg}")

        logger.info("[Extractor] Layer 1 extraction complete")
        return result

    finally:
        # Always close document and quit the dedicated SW instance
        if swModel is not None:
            try:
                swApp.CloseDoc(temp_path)
                logger.info("[Extractor] Document closed")
            except Exception as e:
                logger.warning(f"[Extractor] CloseDoc error: {e}")
        if swApp is not None and not locals().get("attached_existing_session", False):
            try:
                swApp.ExitApp()
                logger.info("[Extractor] SolidWorks instance exited")
            except Exception as e:
                logger.warning(f"[Extractor] ExitApp error: {e}")
        elif swApp is not None:
            logger.info("[Extractor] Attached SolidWorks session left running")
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _check_cancel(cancel_event: threading.Event, stage: str) -> None:
    if cancel_event.is_set():
        raise InterruptedError(f"Job cancelled at stage: {stage}")


def _sha256(path: str) -> str:
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()
