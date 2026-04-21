"""
solidworks_extractor.py — Opens a dedicated SolidWorks instance and runs all
10 extraction modules sequentially.

Safety contract (from baseline v3):
  - v1.0.43 test mode attaches to a running SolidWorks session when present so
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


from extractor.extract_properties    import ExtractProperties
from extractor.extract_sheets        import ExtractSheets
from extractor.extract_views         import ExtractViews
from extractor.extract_dimensions    import ExtractDimensions
from extractor.extract_annotations   import ExtractAnnotations
from extractor.extract_tables        import ExtractTables
from extractor.extract_references    import ExtractReferences
from extractor.extract_health        import ExtractHealth
from extractor.extract_nozzles       import ExtractNozzles
from extractor.extract_design_data   import ExtractDesignDataTable
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


def _summarize_extraction_result(result: dict) -> dict:
    views = result.get("views") or []
    dimensions = result.get("dimensions") or {}
    annotations = result.get("annotations") or {}
    tables = result.get("tables") or {}
    design_data = result.get("design_data") or {}
    model_refs = [v.get("model_reference") for v in views if isinstance(v, dict) and v.get("model_reference")]
    summary = {
        "views_total": len(views),
        "model_reference_populated_count": len(model_refs),
        "model_reference_unique_count": len(set(model_refs)),
        "dimensions_total_count": int(dimensions.get("total_count") or 0),
        "annotations_total_seen": int(annotations.get("total_annotations_seen") or 0),
        "bom_found": bool(tables.get("bom_found")),
        "bom_rows": int(tables.get("bom_rows") or 0),
        "design_data_status": design_data.get("status", ""),
        "design_data_source": design_data.get("source", ""),
    }
    return summary


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
        "schema_version": "1.0",
        "agent":          {},        # stamped by runner
        "file": {
            "original_filename": filename,
            "file_size_bytes":   file_size,
            "sha256":            sha256,
        },
        "properties":         {},
        "sheets":             [],
        "views":              [],
        "dimensions":         {},
        "annotations":        {},
        "tables":             {},
        "references":         {},
        "health":             {},
        "nozzles":            {},
        "design_data":        {},
        "design_data_table":  {},
        "extraction_warnings": [],
        "extraction_errors": {
            "properties":        None,
            "sheets":            None,
            "views":             None,
            "dimensions":        None,
            "annotations":       None,
            "tables":            None,
            "references":        None,
            "health":            None,
            "nozzles":           None,
            "design_data_table": None,
        },
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
        swDraw = refetch_active_drawing_doc(swApp, swDraw)

        try:
            first_view = swDraw.GetFirstView()
            logger.info(
                "[Extractor] GetFirstView: OK"
                + (" (view returned)" if first_view else " (None — may activate sheet first)")
            )
        except Exception as e:
            err_msg = f"{type(e).__name__}: {e}"
            logger.warning(f"[Extractor] API call failed: IDrawingDoc.GetFirstView: {err_msg}; continuing best-effort")
            result["extraction_errors"]["IDrawingDoc.GetFirstView"] = err_msg
            result["extraction_warnings"].append(f"IDrawingDoc.GetFirstView unavailable: {err_msg}")

        # ── Run modules ───────────────────────────────────────────────────────
        modules = [
            ("properties",        lambda: ExtractProperties(swApp, swModel, logger)),
            ("sheets",            lambda: ExtractSheets(swApp, swModel, swDraw, logger)),
            ("views",             lambda: ExtractViews(swApp, swModel, swDraw, logger)),
            ("dimensions",        lambda: ExtractDimensions(swApp, swModel, swDraw, logger)),
            ("annotations",       lambda: ExtractAnnotations(swApp, swModel, swDraw, logger)),
            ("tables",            lambda: ExtractTables(swApp, swModel, swDraw, logger)),
            ("references",        lambda: ExtractReferences(swApp, swModel, swDraw, logger)),
            ("health",            lambda: ExtractHealth(swApp, swModel, swDraw, logger)),
            ("nozzles",           lambda: ExtractNozzles(swApp, swModel, swDraw, logger)),
            ("design_data_table", lambda: ExtractDesignDataTable(
                swApp, swModel, swDraw, logger)),
        ]

        for key, fn in modules:
            _check_cancel(cancel_event, f"before {key}")
            t0 = time.monotonic()
            try:
                result[key] = fn()
                logger.debug(f"[Extractor] {key} OK ({time.monotonic() - t0:.2f}s)")
            except Exception as e:
                err_msg = f"{type(e).__name__}: {e}"
                logger.error(f"[Extractor] {key} SOFT FAIL: {err_msg}")
                result["extraction_errors"][key] = err_msg
                result["extraction_warnings"].append(f"{key} extraction failed: {err_msg}")

        ddt = result.get("design_data_table") or {}
        result["design_data"] = {
            "status": ddt.get("status", "missing" if not ddt.get("found") else "found"),
            "source": ddt.get("source", "missing"),
        }
        for warning in ddt.get("warnings", []) or []:
            if warning not in result["extraction_warnings"]:
                result["extraction_warnings"].append(warning)

        result["extraction_summary"] = _summarize_extraction_result(result)
        result["open_diagnostics"]["post_open_extraction_summary"] = result["extraction_summary"]
        logger.info(
            "[Extractor] Post-open extraction summary: "
            f"model_refs={result['extraction_summary']['model_reference_populated_count']}/{result['extraction_summary']['views_total']} "
            f"dims={result['extraction_summary']['dimensions_total_count']} "
            f"annotations={result['extraction_summary']['annotations_total_seen']} "
            f"bom_found={result['extraction_summary']['bom_found']} "
            f"design_data={result['extraction_summary']['design_data_status']}:{result['extraction_summary']['design_data_source']}"
        )

        logger.info("[Extractor] All modules complete")
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
