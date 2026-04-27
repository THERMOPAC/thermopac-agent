"""
sw_instance.py — Shared SolidWorks COM instance helpers.

Used by both the extraction agent (solidworks_extractor.py) and the
structuring agent (solidworks_structurer.py).

Functions:
  _get_sldworks_pids()             — snapshot running SLDWORKS.EXE PIDs
  _kill_orphan_sw_process(pid)     — force-kill a specific PID via taskkill
  _launch_sw_dedicated_instance()  — create isolated SW process via DispatchEx
  _log_sw_version()                — log SW COM version string

Safety contract:
  - GetActiveObject() is NEVER used — the agent must never attach to an
    engineer's running SolidWorks session.
  - All callers are responsible for ExitApp() + orphan-kill in their
    own finally blocks.
"""

from __future__ import annotations

try:
    import win32com.client
    import win32com.client.gencache
    import pythoncom
    PYWIN32_AVAILABLE = True
except ImportError:
    win32com    = None
    pythoncom   = None
    PYWIN32_AVAILABLE = False


def _get_sldworks_pids() -> set:
    """
    Return the set of SLDWORKS.EXE process IDs currently running.
    Used to identify which PID belongs to the agent's dedicated instance
    so it can be force-killed if ExitApp() fails.
    """
    import subprocess
    try:
        result = subprocess.run(
            ['tasklist', '/fi', 'imagename eq SLDWORKS.EXE', '/fo', 'csv', '/nh'],
            capture_output=True, text=True, timeout=10,
        )
        pids = set()
        for line in result.stdout.strip().splitlines():
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2:
                try:
                    pids.add(int(parts[1]))
                except ValueError:
                    pass
        return pids
    except Exception:
        return set()


def _kill_orphan_sw_process(pid: int, logger) -> None:
    """
    Force-kill a specific SLDWORKS.EXE process by PID.
    Called only when ExitApp() fails or raises, to ensure no orphan remains.
    """
    import subprocess
    try:
        result = subprocess.run(
            ['taskkill', '/F', '/PID', str(pid)],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0:
            logger.info(f"[COM] Orphan guard: killed SLDWORKS.EXE PID {pid} via taskkill")
        else:
            logger.warning(f"[COM] Orphan guard: taskkill /F /PID {pid} — {result.stdout.strip()}")
    except Exception as e:
        logger.warning(f"[COM] Orphan guard: taskkill failed for PID {pid}: {e}")


def _log_sw_version(sw_app, logger) -> None:
    """Log the detected SolidWorks version from the COM object."""
    version = "unknown"
    for attr in ("RevisionNumber", "Version"):
        try:
            value = getattr(sw_app, attr)
            version = value() if callable(value) else value
            if version:
                break
        except Exception:
            pass
    logger.info(f"[COM] SolidWorks version: {version}")


def _launch_sw_dedicated_instance(progid: str, logger):
    """
    Launch a DEDICATED, ISOLATED SolidWorks instance.

    Decision: ALWAYS use DispatchEx() as the primary connection method.
      - DispatchEx() creates a NEW COM server process every time.
      - GetActiveObject() is intentionally NOT used — the agent must never
        attach to or interfere with an engineer's running SolidWorks session.

    Binding priority:
      1. DispatchEx(progid)                  — dedicated new process, preferred
      2. DispatchEx("SldWorks.Application")  — generic fallback
      3. gencache.EnsureDispatch(progid)     — early binding last resort

    Returns: (sw_app, binding_mode: str)
    Raises RuntimeError if all three methods fail.
    """
    logger.info(f"[COM] Launching dedicated SolidWorks instance: {progid}")
    logger.info("[COM] Mode: dedicated — GetActiveObject() path is disabled")

    # Method 1: DispatchEx with versioned ProgID (preferred)
    try:
        sw_app = win32com.client.DispatchEx(progid)
        binding_mode = "dedicated-DispatchEx"
        logger.info(f"[COM] Instance created via DispatchEx({progid})")
        _log_sw_version(sw_app, logger)
        return sw_app, binding_mode
    except Exception as e:
        logger.warning(f"[COM] DispatchEx({progid}) failed: {type(e).__name__}: {e}")

    # Method 2: DispatchEx with generic ProgID
    try:
        sw_app = win32com.client.DispatchEx("SldWorks.Application")
        binding_mode = "dedicated-DispatchEx-generic"
        logger.info("[COM] Instance created via DispatchEx(SldWorks.Application)")
        _log_sw_version(sw_app, logger)
        return sw_app, binding_mode
    except Exception as e:
        logger.warning(f"[COM] DispatchEx(SldWorks.Application) failed: {type(e).__name__}: {e}")

    # Method 3: EnsureDispatch (early binding, may share process — last resort)
    try:
        sw_app = win32com.client.gencache.EnsureDispatch(progid)
        binding_mode = "dedicated-EnsureDispatch-fallback"
        logger.warning("[COM] Instance via EnsureDispatch — may share existing process (fallback)")
        _log_sw_version(sw_app, logger)
        return sw_app, binding_mode
    except Exception as e:
        logger.error(f"[COM] EnsureDispatch fallback also failed: {type(e).__name__}: {e}")

    raise RuntimeError(
        f"All three DispatchEx methods failed for '{progid}'. "
        "Ensure SolidWorks is installed and the ProgID is registered in HKCR."
    )
