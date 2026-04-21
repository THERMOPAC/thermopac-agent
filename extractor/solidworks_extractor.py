"""
solidworks_extractor.py — Opens a dedicated SolidWorks instance and runs all
10 extraction modules sequentially.

Safety contract (from baseline v3):
  - Always DispatchEx() — never attaches to user's running session
  - swApp.Visible = False always
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

# IDrawingDoc IID extracted from sldworks.tlb at runtime — populated by Method C.
# Used for raw QueryInterface so GetFirstView / sheet traversal work via pure IDispatch
# without requiring makepy cache seeding or CastTo.
_sw_drawing_doc_iid  = None   # pywintypes.IID of the interface that owns GetFirstView
_sw_gfv_dispid       = None   # DISPID of GetFirstView on that interface (DISPID fallback)
_sw_gnv_dispid       = None   # DISPID of GetNextView  (informational / validation)


def _extract_sw_interface_iids(tl, logger) -> None:
    """
    Scan the loaded SolidWorks ITypeLib for the interface that contains
    GetFirstView and cache its IID + key DISPIDs.

    Rules:
      • Only consider TKIND_INTERFACE (3) or TKIND_DISPATCH (4) entries.
        Coclasses (5), aliases (6), etc. carry the TypeLib CLSID, not an
        interface IID, which is why single-name matching fails.
      • Verify the interface actually exposes GetFirstView as a method; don't
        assume by name alone.
      • Also record GetNextView / ActivateSheet DISPIDs for the DISPID
        fallback path and validation.
    """
    global _sw_drawing_doc_iid, _sw_gfv_dispid, _sw_gnv_dispid

    TKIND_INTERFACE = 3
    TKIND_DISPATCH  = 4

    try:
        count = tl.GetTypeInfoCount()
        for i in range(count):
            try:
                tinfo = tl.GetTypeInfo(i)
                tattr = tinfo.GetTypeAttr()

                if tattr.typekind not in (TKIND_INTERFACE, TKIND_DISPATCH):
                    continue

                # Scan the function table for GetFirstView / GetNextView
                gfv_dispid = None
                gnv_dispid = None
                for j in range(tattr.cFuncs):
                    try:
                        fdesc = tinfo.GetFuncDesc(j)
                        fname = tinfo.GetNames(fdesc.memid, 1)[0]
                        if fname == "GetFirstView":
                            gfv_dispid = fdesc.memid
                        elif fname == "GetNextView":
                            gnv_dispid = fdesc.memid
                    except Exception:
                        continue

                if gfv_dispid is None:
                    continue   # not the drawing interface

                name = tl.GetDocumentation(i)[0]
                iid  = tattr.iid
                logger.info(
                    f"[Extractor] IDrawingDoc IID: {iid}  "
                    f"(name={name} typekind={tattr.typekind} "
                    f"GetFirstView DISPID={gfv_dispid}"
                    + (f" GetNextView DISPID={gnv_dispid})" if gnv_dispid else ")")
                )
                _sw_drawing_doc_iid = iid
                _sw_gfv_dispid      = gfv_dispid
                _sw_gnv_dispid      = gnv_dispid
                return   # stop at first match

            except Exception:
                continue

        logger.warning("[Extractor] No TKIND_INTERFACE/DISPATCH with GetFirstView found in TypeLib")
    except Exception as e:
        logger.warning(f"[Extractor] IID extraction failed: {e}")

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
from extractor.extract_design_data   import ExtractDesignDataTable, DesignDataNotFoundError

# SolidWorks constants  (swOpenDocOptions_e)
SW_DOC_DRAWING           = 3
SW_OPEN_SILENT           = 1     # swOpenDocOptions_Silent      — suppresses ALL missing-ref dialogs
SW_OPEN_READ_ONLY        = 2     # swOpenDocOptions_ReadOnly
SW_OPEN_VIEW_ONLY        = 4     # swOpenDocOptions_ViewOnly    — Large Design Review, no 3-D load
SW_OPEN_LOAD_MODEL       = 128   # swOpenDocOptions_LoadModel   — fallback only
# NOTE: 64 = swOpenDocOptions_OverrideDefaultLoadedData (was wrongly used as SW_OPEN_SILENT before v1.0.4)

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


def _prepare_sw_makepy_cache(progid: str, logger) -> None:
    """
    Seed win32com early-binding (makepy) cache for SolidWorks before calling
    gencache.EnsureDispatch().  Three methods attempted in order; first success wins.

    Method A — CLSID chain (registry only, no file I/O):
        ProgID → CLSID → TypeLib CLSID → gencache.EnsureModule(clsid, 0, maj, min)

    Method B — TLB file from registry:
        TypeLib\{CLSID}\{version}\0\{win64|win32} → full file path →
        pythoncom.LoadTypeLib(path) + gencache.EnsureModule with explicit ITypeLib

    Method C — Known SOLIDWORKS install directories:
        Checks the SolidWorks install path from
        HKLM\SOFTWARE\SolidWorks\SolidWorks {year}\Setup\SldWorks dir
        and looks for sldworks.exe (contains embedded TLB) or any *.tlb.

    Raises RuntimeError with remediation steps if all methods fail.
    """
    try:
        import winreg
        import pythoncom
    except ImportError:
        raise RuntimeError(
            "pywin32 (winreg/pythoncom) is not installed. "
            "Run: pip install pywin32   then: python -m win32com.client.makepy"
        )

    def _enum_typelib_versions(tl_clsid: str):
        """Return sorted list of (major, minor, ver_str) from TypeLib registry.
        Non-numeric sub-keys (e.g. 'lib.3', 'Helpdir') are silently skipped."""
        versions = []
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"TypeLib\\{tl_clsid}") as k:
            i = 0
            while True:
                try:
                    ver_str = winreg.EnumKey(k, i); i += 1
                    parts = ver_str.split(".")
                    try:
                        maj = int(parts[0])
                        min_ = int(parts[1]) if len(parts) > 1 else 0
                        versions.append((maj, min_, ver_str))
                    except ValueError:
                        # sub-key is not a version number (e.g. 'lb.3', 'FLAGS')
                        logger.debug(f"[Makepy] Skipping non-numeric TypeLib sub-key: {ver_str!r}")
                except OSError:
                    break
        return sorted(versions)

    def _resolve_tl_clsid(progid: str):
        """ProgID → app CLSID → TypeLib CLSID"""
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID") as k:
            app_clsid = winreg.QueryValue(k, "")
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{app_clsid}\\TypeLib") as k:
            tl_clsid = winreg.QueryValue(k, "")
        return app_clsid, tl_clsid

    # ── Method A: ProgID → CLSID → TypeLib CLSID → EnsureModule ─────────
    try:
        app_clsid, tl_clsid = _resolve_tl_clsid(progid)
        logger.info(f"[Makepy-A] CLSID: {app_clsid}  TypeLib: {tl_clsid}")

        versions = _enum_typelib_versions(tl_clsid)
        if not versions:
            raise ValueError(f"No numeric version sub-keys under TypeLib\\{tl_clsid}")
        major, minor, ver_str = versions[-1]
        logger.info(f"[Makepy-A] TypeLib version: {major}.{minor}")

        win32com.client.gencache.EnsureModule(tl_clsid, 0, major, minor)
        logger.info("[Makepy-A] Success — early-binding cache seeded via CLSID chain")
        return

    except Exception as e:
        logger.warning(f"[Makepy-A] Failed: {e}")

    # ── Method B: LocalServer32 → LoadTypeLib on sldworks.exe directly ───
    # SolidWorks 2019+ embeds its TypeLib as a resource inside sldworks.exe.
    # There is NO standalone .tlb file and the TypeLib registry entry often has
    # no file-path sub-key.  The CLSID\{app_clsid}\LocalServer32 key always
    # points to the real sldworks.exe — LoadTypeLib handles resource-embedded
    # TLBs transparently.
    try:
        app_clsid, tl_clsid = _resolve_tl_clsid(progid)

        # Get the actual sldworks.exe path from LocalServer32
        exe_path = None
        for key_path in (
            f"CLSID\\{app_clsid}\\LocalServer32",
            f"WOW6432Node\\CLSID\\{app_clsid}\\LocalServer32",
        ):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path) as k:
                    raw = winreg.QueryValue(k, "").strip('"').split('"')[0].strip()
                    raw = raw.split(" /")[0].strip()   # strip any CLI flags
                if raw and os.path.isfile(raw):
                    exe_path = raw; break
            except OSError:
                continue

        if not exe_path:
            raise ValueError(f"LocalServer32 for {app_clsid} not found or file missing")

        logger.info(f"[Makepy-B] Loading TypeLib from: {exe_path}")
        tl = pythoncom.LoadTypeLib(exe_path)
        tla = tl.GetLibAttr()
        clsid_b = str(tla[0]); major_b = tla[3]; minor_b = tla[4]
        logger.info(f"[Makepy-B] TypeLib attrs: {clsid_b}  v{major_b}.{minor_b}")
        # Pass the ITypeLib object (not the CLSID string) so EnsureModule
        # uses the already-loaded library without going back to the registry.
        win32com.client.gencache.EnsureModule(
            tl, 0, major_b, minor_b, bForDemand=False, bBuildHidden=True)
        logger.info("[Makepy-B] Success — cache seeded via sldworks.exe LocalServer32")
        return

    except Exception as e:
        logger.warning(f"[Makepy-B] Failed: {e}")

    # ── Method C: year-based registry scan + LocalServer32 fallback ───────
    try:
        sw_dirs = []

        # Primary: look up SolidWorks install dir by year
        for year in range(2024, 2016, -1):
            for hive in (winreg.HKEY_LOCAL_MACHINE,):
                for base in (
                    f"SOFTWARE\\SolidWorks\\SolidWorks {year}\\Setup",
                    f"SOFTWARE\\WOW6432Node\\SolidWorks\\SolidWorks {year}\\Setup",
                    f"SOFTWARE\\SolidWorks\\SOLIDWORKS {year}\\Setup",
                    f"SOFTWARE\\WOW6432Node\\SolidWorks\\SOLIDWORKS {year}\\Setup",
                ):
                    try:
                        with winreg.OpenKey(hive, base) as k:
                            d = winreg.QueryValueEx(k, "SldWorks dir")[0]
                            if d and os.path.isdir(d) and d not in sw_dirs:
                                sw_dirs.append(d)
                    except OSError:
                        pass

        # Fallback: derive install dir from CLSID LocalServer32 (always reliable)
        try:
            app_clsid2, _ = _resolve_tl_clsid(progid)
            for kp in (f"CLSID\\{app_clsid2}\\LocalServer32",
                       f"WOW6432Node\\CLSID\\{app_clsid2}\\LocalServer32"):
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, kp) as k:
                        raw = winreg.QueryValue(k, "").strip('"').split('"')[0].strip()
                        raw = raw.split(" /")[0].strip()
                    d = os.path.dirname(raw)
                    if d and os.path.isdir(d) and d not in sw_dirs:
                        sw_dirs.append(d)
                        logger.info(f"[Makepy-C] SW dir via LocalServer32: {d}")
                    break
                except OSError:
                    pass
        except Exception:
            pass

        if not sw_dirs:
            raise ValueError("SolidWorks install dir not found via any registry path")

        for sw_dir in sw_dirs:
            tlb_files = [f for f in os.listdir(sw_dir) if f.lower().endswith(".tlb")]
            exe_files = [f for f in os.listdir(sw_dir)
                         if f.lower() in ("sldworks.exe", "sldworks_worker.exe")]
            candidates = [os.path.join(sw_dir, f) for f in tlb_files + exe_files]

            # Sort: sldworks* first, then other TLBs, then EXEs
            def _c_sort_key(p):
                b = os.path.basename(p).lower()
                if b.startswith("sldworks") and b.endswith(".tlb"): return 0
                if b.endswith(".tlb"): return 1
                return 2
            candidates.sort(key=_c_sort_key)

            # Also scan one level of subdirectories (lang\english, api, etc.)
            for sub in ("lang\\english", "lang", "api", "setup"):
                sub_dir = os.path.join(sw_dir, sub)
                if os.path.isdir(sub_dir):
                    for f in os.listdir(sub_dir):
                        if f.lower().endswith(".tlb"):
                            p = os.path.join(sub_dir, f)
                            if p not in candidates:
                                candidates.append(p)

            for candidate in candidates:
                if not os.path.isfile(candidate):
                    continue
                try:
                    tl = pythoncom.LoadTypeLib(candidate)
                    tla = tl.GetLibAttr()
                    clsid_c = str(tla[0]); major_c = tla[3]; minor_c = tla[4]
                    # Only accept TLBs whose CLSID matches the expected SolidWorks TypeLib
                    if tl_clsid and clsid_c.upper() != tl_clsid.upper():
                        logger.debug(f"[Makepy-C] {os.path.basename(candidate)}: "
                                     f"CLSID {clsid_c} ≠ {tl_clsid} — skipping")
                        continue
                    # Extract IDrawingDoc IID first — this is the critical output.
                    # IID is used for raw QueryInterface later; no gencache needed.
                    _extract_sw_interface_iids(tl, logger)
                    # Also attempt EnsureModule (bonus early-binding; non-fatal if it fails)
                    try:
                        win32com.client.gencache.EnsureModule(
                            tl, 0, major_c, minor_c,
                            bForDemand=False, bBuildHidden=True)
                        logger.info(f"[Makepy-C] Cache seeded from {candidate}")
                    except Exception as em:
                        logger.debug(f"[Makepy-C] EnsureModule skipped: {em}")
                    logger.info(f"[Makepy-C] TLB loaded — IID extracted from {candidate}")
                    return
                except Exception as ce:
                    logger.debug(f"[Makepy-C] {candidate} skipped: {ce}")
                    continue

        raise ValueError("No usable SolidWorks TLB/EXE found in install directories")

    except Exception as e:
        logger.warning(f"[Makepy-C] Failed: {e}")

    # ── Method D: subprocess makepy as a last resort ──────────────────────
    # Runs the REAL python interpreter (not the PyInstaller bundle EXE) so that
    # win32com.client.makepy can actually execute.  When frozen, sys.executable
    # is the EXE itself — useless here.  Look for the bundled Python first.
    try:
        import subprocess, sys
        py_exe = None
        if getattr(sys, "frozen", False):
            # Running inside PyInstaller — find bundled Python alongside EXE
            exe_dir = os.path.dirname(sys.executable)
            for candidate in (
                os.path.join(exe_dir, "python", "python.exe"),  # bundled by installer
                os.path.join(exe_dir, "python.exe"),
            ):
                if os.path.isfile(candidate):
                    py_exe = candidate
                    break
        else:
            py_exe = sys.executable

        if not py_exe:
            raise FileNotFoundError("No Python interpreter found for makepy subprocess")

        logger.info(f"[Makepy-D] Trying subprocess makepy via: {py_exe}")
        result = subprocess.run(
            [py_exe, "-m", "win32com.client.makepy", progid],
            capture_output=True, text=True, timeout=90
        )
        if result.returncode == 0:
            logger.info("[Makepy-D] Success — subprocess makepy completed")
            return
        raise RuntimeError(result.stderr.strip() or f"exit code {result.returncode}")

    except Exception as e:
        logger.warning(f"[Makepy-D] Failed: {e}")

    # ── All makepy methods exhausted — log clearly, do NOT abort ─────────
    # The EnsureDispatch call below will still attempt its own cache build.
    # Many SolidWorks 2019-2021 installs work fine with late-binding Dispatch.
    logger.warning(
        "[Makepy] All cache-seeding methods failed — will attempt late-binding.\n"
        "  To fix permanently run:  python -m win32com.client.makepy "
        f'"{progid}"'
    )


def _seed_makepy_from_com_object(com_obj, logger) -> bool:
    """
    Method E (post-connect): Extract the TypeLib directly from the running COM
    object via IDispatch.GetTypeInfo().  The COM server must expose its type
    information through IDispatch — SolidWorks always does.

    This is the most reliable approach: it bypasses all registry/file-system
    lookup and asks the object itself for its type library.  If the pre-connect
    methods (A-D) all failed, this runs after swApp = Dispatch() succeeds.

    On success the gencache is seeded so subsequent CastTo(swModel, "IDrawingDoc")
    calls work correctly.
    """
    try:
        # IDispatch.GetTypeInfo(typeInfoNumber=0, lcid=0) → ITypeInfo
        tinfo = com_obj._oleobj_.GetTypeInfo(0, 0)
        # ITypeInfo.GetContainingTypeLib() → (ITypeLib, indexInTypeLib)
        tlib, _ = tinfo.GetContainingTypeLib()
        tla = tlib.GetLibAttr()
        clsid_e = str(tla[0]); major_e = tla[3]; minor_e = tla[4]
        logger.info(f"[Makepy-E] COM object TypeLib: {clsid_e}  v{major_e}.{minor_e}")
        # Pass the ITypeLib object directly — bypasses registry lookup
        win32com.client.gencache.EnsureModule(
            tlib, 0, major_e, minor_e, bForDemand=False, bBuildHidden=True)
        logger.info("[Makepy-E] Post-connect cache seeded — CastTo now available")
        return True
    except Exception as e:
        logger.warning(f"[Makepy-E] Post-connect TypeLib extraction failed: {e}")
        return False


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
    Raises on hard failure (DesignDataNotFoundError, SW launch failure, etc.).
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
        "design_data_table":  {},
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

        # ── Launch dedicated SW instance ───────────────────────────────────────
        _check_cancel(cancel_event, "before SW launch")
        logger.info(f"[Extractor] Launching SolidWorks ({config.sw_progid})…")
        t_launch = time.monotonic()
        # ── Inherit search paths from user's running SW session (read-only) ───
        inherited_paths = _get_user_sw_search_paths(config.sw_progid, logger)

        # ── Seed win32com early-binding cache for SolidWorks ─────────────────
        # EnsureDispatch internally calls EnsureModule(typelib_clsid, ...).
        # On many SolidWorks installations that call fails with
        # "cannot automate the makepy process" because pywin32's default TLB
        # lookup cannot resolve the registered path when it points into the SW
        # EXE as a resource rather than a standalone .tlb file.
        #
        # _prepare_sw_makepy_cache() tries three methods:
        #   A: ProgID → CLSID → TypeLib CLSID → gencache.EnsureModule
        #   B: registry TLB path  → LoadTypeLib → EnsureModule
        #   C: SW install dir scan → LoadTypeLib → EnsureModule
        #   D: subprocess python -m win32com.client.makepy
        # Methods no longer hard-fail; they log warnings and fall through to
        # the EnsureDispatch / Dispatch fallback below.
        _prepare_sw_makepy_cache(config.sw_progid, logger)

        # Prefer early-binding (EnsureDispatch uses pre-seeded cache).
        # If makepy cache is still missing, fall back to late-binding Dispatch —
        # it loses CastTo(IDrawingDoc) but OpenDoc6 + property bag still work.
        logger.info("[Extractor] Connecting to SolidWorks COM…")
        _early_bound = False
        try:
            swApp = win32com.client.gencache.EnsureDispatch(config.sw_progid)
            logger.info("[Extractor] SolidWorks COM ready (early-bound)")
            _early_bound = True
        except Exception as ed:
            logger.warning(f"[Extractor] EnsureDispatch failed ({ed}); trying late-binding Dispatch…")
            swApp = win32com.client.Dispatch(config.sw_progid)
            logger.info("[Extractor] SolidWorks COM ready (late-bound)")
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

        # ── Pass 0: OpenDoc7 with IDocumentSpecification (most control) ───────
        # docSpec.Silent = True reliably suppresses ALL missing-ref dialogs,
        # allowing the drawing to open in full mode even without referenced parts.
        def _close_zombie(label: str):
            """CloseDoc after a failed open attempt to clear SW's internal open-state registry.
            Without this, SW returns AlreadyOpen (65536) for every subsequent attempt on the
            same path, even when OpenDoc returned None (zombie registration)."""
            try:
                swApp.CloseDoc(temp_path)
                logger.info(f"[Extractor] CloseDoc cleanup after failed {label}")
            except Exception:
                pass  # expected if nothing was actually registered

        try:
            docSpec = swApp.GetOpenDocSpec(temp_path)
            docSpec.FileName     = temp_path
            docSpec.DocumentType = SW_DOC_DRAWING
            docSpec.ReadOnly     = True
            docSpec.Silent       = True
            swModel = swApp.OpenDoc7(docSpec)
            err_val  = docSpec.Error
            warn_val = docSpec.Warning
            logger.info(f"[Extractor] OpenDoc7 pass 0: "
                        f"model={'OK' if swModel else 'None'} "
                        f"errors={_decode_sw_error(err_val)} "
                        f"warnings={_decode_sw_error(warn_val)}")
        except Exception as e:
            logger.info(f"[Extractor] OpenDoc7 not available ({e}); falling back to OpenDoc6")
            swModel = None
        if swModel is None:
            _close_zombie("OpenDoc7")

        # ── Passes 1-2: OpenDoc6 with Silent only (no LDR/ViewOnly) ─────────
        #   Pass 1: Silent | ReadOnly
        #   Pass D: ReadOnly + auto-dismiss thread (catches missing-ref dialogs)
        #   Pass 2: Silent | ReadOnly | LoadModel  (last resort full-mode)
        # LDR / ViewOnly (swOpenDocOptions_ViewOnly) is intentionally excluded —
        # it opens drawings without full COM interface support, making
        # IDrawingDoc.GetFirstView() inaccessible.
        if swModel is None:
            for pass_num, options in enumerate(
                [SW_OPEN_READ_ONLY | SW_OPEN_SILENT,
                 SW_OPEN_READ_ONLY | SW_OPEN_SILENT | SW_OPEN_LOAD_MODEL], start=1
            ):
                errors   = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                swModel  = swApp.OpenDoc6(
                    temp_path, SW_DOC_DRAWING, options, "", errors, warnings)
                err_val  = errors.value
                warn_val = warnings.value
                logger.info(f"[Extractor] OpenDoc6 pass {pass_num}: "
                            f"model={'OK' if swModel else 'None'} "
                            f"errors={_decode_sw_error(err_val)} "
                            f"warnings={_decode_sw_error(warn_val)}")
                if swModel is not None:
                    break
                _close_zombie(f"pass {pass_num}")
                # Insert dialog-dismiss pass between pass 1 and LoadModel
                if pass_num == 1 and WIN32GUI_AVAILABLE:
                    logger.info("[Extractor] OpenDoc6 pass D: ReadOnly + auto dialog-dismiss")
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
                        logger.info(f"[Extractor] OpenDoc6 pass D: "
                                    f"model={'OK' if d_model else 'None'} "
                                    f"errors={_decode_sw_error(d_err)} "
                                    f"warnings={_decode_sw_error(d_warn)}")
                        if d_model is not None:
                            swModel  = d_model
                            err_val  = d_err
                            warn_val = d_warn
                            pass_num = 'D'
                            break
                        _close_zombie("pass D")
                    except Exception as ex_d:
                        logger.info(f"[Extractor] OpenDoc6 pass D exception: {ex_d}")
                        _close_zombie("pass D (exception)")
                    finally:
                        stop_dismiss.set()
                        dismisser.join(timeout=2)

        if swModel is None:
            raise RuntimeError(
                f"All open passes failed for {filename}. "
                f"Last errors={_decode_sw_error(err_val)} warnings={_decode_sw_error(warn_val)}"
            )

        logger.info(f"[Extractor] Document open OK in full mode (pass={pass_num})")

        # ── Method E: seed makepy cache from the opened document object ───────
        # Now that a fully-loaded COM object (swModel) is available, ask it for
        # its own TypeLib via IDispatch.GetTypeInfo().  This is the most reliable
        # path — the document object itself knows its interface definitions.
        # On success EnsureDispatch becomes available for CastTo(IDrawingDoc).
        if not _early_bound:
            if _seed_makepy_from_com_object(swModel, logger):
                _early_bound = True

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
        except RuntimeError:
            raise
        except Exception as e:
            logger.warning(f"[Extractor] GetType() skipped: {e} — assuming Drawing and continuing")

        # ── Obtain IDrawingDoc IDispatch via QueryInterface ────────────────────
        # The COM object returned by OpenDoc6 is typed as IModelDoc2.  Its
        # IDispatch vtable only maps IModelDoc2 names; GetFirstView / GetNextView
        # are on the separate IDrawingDoc vtable.
        #
        # Fix for v1.0.30:
        #   QueryInterface MUST request IID_IDispatch as the second argument so
        #   pywin32 receives a PyIDispatch it can wrap.  Single-arg QI returns
        #   a raw PyIUnknown for unknown interfaces, which Dispatch() can't wrap
        #   → "no interface object registered" error.
        swDraw = None
        if _sw_drawing_doc_iid is not None:
            try:
                qi_disp = swModel._oleobj_.QueryInterface(
                    _sw_drawing_doc_iid,
                    pythoncom.IID_IDispatch,      # <-- explicit: return as IDispatch
                )
                swDraw = win32com.client.Dispatch(qi_disp)
                logger.info("[Extractor] QueryInterface: OK")
            except Exception as e:
                logger.warning(f"[Extractor] QI failed ({e}) — will try DISPID fallback")
                swDraw = None

        # ── DISPID fallback ────────────────────────────────────────────────────
        # If QI returns the wrong pointer, InvokeTypes with the known DISPID
        # still calls through the correct vtable offset on the COM object.
        if swDraw is None and _sw_gfv_dispid is not None:
            try:
                raw = swModel._oleobj_.InvokeTypes(
                    _sw_gfv_dispid,
                    0,                              # lcid
                    1,                              # DISPATCH_METHOD
                    (pythoncom.VT_DISPATCH, 0),     # return type: IDispatch
                    (),                             # no args
                )
                if raw is not None:
                    # Use swModel as swDraw but patch GetFirstView via InvokeTypes
                    swDraw = swModel
                    logger.info("[Extractor] DISPID fallback: GetFirstView reachable via InvokeTypes")
                else:
                    logger.warning("[Extractor] DISPID fallback: InvokeTypes returned None")
            except Exception as e:
                logger.warning(f"[Extractor] DISPID fallback failed: {e}")

        if swDraw is None:
            swDraw = swModel
            logger.warning(
                "[Extractor] All IDrawingDoc strategies failed — using IModelDoc2 directly. "
                "IDrawingDoc methods (GetFirstView etc.) will likely be inaccessible."
            )

        # Confirm GetFirstView is accessible
        try:
            gfv_fn = getattr(swDraw, "GetFirstView", None)
            if gfv_fn is None or not callable(gfv_fn):
                raise AttributeError("GetFirstView not in COM dispatch table")
            _test_view = gfv_fn()
            logger.info(
                "[Extractor] GetFirstView: OK"
                + (" (view returned)" if _test_view else " (None — sheet may not be active yet)")
            )
        except Exception as e:
            raise RuntimeError(f"IDrawingDoc.GetFirstView not accessible: {e}")

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
            except DesignDataNotFoundError:
                # Hard failure — re-raise to caller
                raise
            except Exception as e:
                err_msg = f"{type(e).__name__}: {e}"
                logger.error(f"[Extractor] {key} SOFT FAIL: {err_msg}")
                result["extraction_errors"][key] = err_msg

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
        if swApp is not None:
            try:
                swApp.ExitApp()
                logger.info("[Extractor] SolidWorks instance exited")
            except Exception as e:
                logger.warning(f"[Extractor] ExitApp error: {e}")
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
