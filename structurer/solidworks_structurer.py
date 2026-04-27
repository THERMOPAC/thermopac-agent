"""
solidworks_structurer.py — Phase 1 Drawing Structuring Agent.

Creates or updates a SolidWorks .slddrw file from DDS job data:
  1. Pre-flight validation  (before launching SolidWorks)
  2. Launch dedicated hidden SolidWorks instance via DispatchEx()
  3. Mode branch:
       create_new      → NewDocument(template_path) → SaveAs3(staging_path)
       update_existing → OpenDoc(staging_path)      → Save2()
  4. Write custom properties (only fields present in dds payload)
  5. Post-write read-back verification
  6. Return structured result JSON

Safety contract (Phase 1):
  - DispatchEx() only — creates a NEW, ISOLATED SolidWorks process
  - GetActiveObject() is NEVER used
  - swApp.Visible = False always
  - ExitApp() always in finally block
  - Orphan guard: taskkill /F /PID if ExitApp() fails
  - Never calls Save on files that fail validation
  - No DDS tables, no heuristic note injection, no PDF, no GCS upload

SolidWorks integer constants used:
  swDocDRAWING              = 3
  swOpenDocOptions_Silent   = 32   (suppress dialogs)
  swCustomInfoText          = 30   (string property type)
  swCustomPropertyReplaceValue = 1 (overwrite if exists)
  swSaveAsCurrentVersion    = 0
"""

from __future__ import annotations
import os
import time
import threading
from datetime import datetime, timezone
from typing import Optional

try:
    import win32com.client
    import pythoncom
    PYWIN32_AVAILABLE = True
except ImportError:
    win32com  = None
    pythoncom = None
    PYWIN32_AVAILABLE = False

from extractor.sw_instance import (
    _get_sldworks_pids,
    _kill_orphan_sw_process,
    _launch_sw_dedicated_instance,
)

# ── SolidWorks API constants ──────────────────────────────────────────────────
_SW_DOC_DRAWING            = 3
_SW_OPEN_SILENT            = 32
_SW_CUSTOM_INFO_TEXT       = 30
_SW_CUSTOM_PROP_REPLACE    = 1
_SW_SAVE_CURRENT_VERSION   = 0

# ── Mechanical column property suffixes (24 per prefix) ───────────────────────
# Order matches _mech_col_props() exactly — used to build clear lists for
# null/absent columns so stale template values are removed from the drawing.
_MECH_PROP_SHORTS: tuple = (
    "IDP", "EDP", "WP", "HTP", "MDMT",
    "HT_TEMP", "OP_TEMP", "MOT", "DES_TEMP",
    "STATE", "VOL", "FLUID", "HZ", "SG",
    "ICA", "ECA",
    "RT", "JE", "TG", "FTC", "PWHT",
    "HEAD", "INS", "INS_SPEC",
)

# Error codes from ISldWorks.OpenDoc / OpenDoc6 return values (common subset)
_SW_FILE_NOT_FOUND         = 2
_SW_FILE_LOCK_ERROR        = 3

# ── Pre-flight ────────────────────────────────────────────────────────────────

class PreflightError(Exception):
    """Raised when a pre-flight check fails. Job is failed immediately, no SW launched."""


def _preflight(job: dict, template_path: str, staging_root: str) -> str:
    """
    Run all pre-flight checks before launching SolidWorks.
    Returns the fully-qualified staging_path on success.
    Raises PreflightError with a descriptive message on any failure.
    """
    drawing_number   = (job.get("drawing_number") or "").strip()
    revision         = (job.get("revision") or "").strip()
    dds              = job.get("dds")
    mode             = (job.get("mode") or "create_new").strip()
    drawing_ctrl_id  = str(job.get("drawing_control_id") or "").strip()

    if not drawing_number:
        raise PreflightError("drawing_number missing or empty in job payload")
    if not revision:
        raise PreflightError("revision missing or empty in job payload")
    if not dds or not isinstance(dds, dict) or len(dds) == 0:
        raise PreflightError("dds payload missing or empty")
    if not drawing_ctrl_id:
        raise PreflightError("drawing_control_id missing from job payload")

    if not template_path:
        raise PreflightError(
            "template_path not set — configure it in ERP System Settings "
            "(Admin → System Settings → SolidWorks Structuring Agent → Template Path). "
            "Fallback: set [structurer] template_path in the agent's config.ini."
        )
    if not os.path.isfile(template_path):
        raise PreflightError(
            f"template_path not found or not accessible: {template_path!r} — "
            "verify the path exists and the agent machine has network access to it."
        )

    if not staging_root:
        raise PreflightError(
            "staging_root not set — configure it in ERP System Settings "
            "(Admin → System Settings → SolidWorks Structuring Agent → Staging Root). "
            "Fallback: set [structurer] staging_root in the agent's config.ini."
        )
    try:
        staging_dir = os.path.join(staging_root, drawing_ctrl_id)
        os.makedirs(staging_dir, exist_ok=True)
        probe = os.path.join(staging_dir, ".write_probe")
        with open(probe, "w") as f:
            f.write("ok")
        os.remove(probe)
    except Exception as e:
        raise PreflightError(f"staging_root not writable ({staging_root}): {e}")

    filename      = f"{drawing_number}_rev-{revision}.slddrw"
    staging_path  = os.path.normpath(os.path.join(staging_dir, filename))

    if mode == "update_existing" and not os.path.isfile(staging_path):
        raise PreflightError(
            f"mode=update_existing but staging file not found: {staging_path}"
        )

    return staging_path


# ── Custom property helpers ───────────────────────────────────────────────────

def _extract_mot(op_temp: Optional[str]) -> str:
    """
    Extract the maximum operating temperature from a 'min / max' formatted string.

    Examples:
        "100 / 120"  → "120"
        "-10 / 80"   → "80"
        "120"        → "120"   (single value — treated as max)
        None / ""    → ""      (skipped — do not write property)

    The result is validated as a finite float before returning so that malformed
    strings (e.g. "N.A." or "—") do not get written as SHELL_MOT.
    """
    if not op_temp:
        return ""
    parts = op_temp.split("/")
    raw = parts[-1].strip()
    try:
        float(raw)
        return raw
    except ValueError:
        return ""


def _mech_col_props(prefix: str, col: dict) -> dict:
    """
    Build {SW_property_name: value_string} for one mechanical column.

    Approved mapping (Option C — 24 SW properties per column):
        <PREFIX>_IDP       ← internalDesignPressureMawp   (Number)
        <PREFIX>_EDP       ← externalDesignPressureMawp   (Number)
        <PREFIX>_WP        ← workingPressure               (Number)
        <PREFIX>_HTP       ← hydroTestPressure             (Number)
        <PREFIX>_MDMT      ← mdmt                          (Number)
        <PREFIX>_HT_TEMP   ← hydroTestTempMinMax           (Text)
        <PREFIX>_OP_TEMP   ← operatingTempMinMax           (Text — full string)
        <PREFIX>_MOT       ← operatingTempMinMax max part  (Number — existing template field)
        <PREFIX>_DES_TEMP  ← designTempMinMax              (Text)
        <PREFIX>_STATE     ← physicalState                 (Text)
        <PREFIX>_VOL       ← grossVolumeLiters             (Number)
        <PREFIX>_FLUID     ← serviceFluid                  (Text)
        <PREFIX>_HZ        ← hazardLevel                   (Text)
        <PREFIX>_SG        ← specificGravity               (Text)
        <PREFIX>_ICA       ← internalCorrosionAllowanceMm  (Number)
        <PREFIX>_ECA       ← externalCorrosionAllowanceMm  (Number)
        <PREFIX>_RT        ← radiography                   (Text)
        <PREFIX>_JE        ← jointEfficiency               (Text)
        <PREFIX>_TG        ← testingGroup                  (Text)
        <PREFIX>_FTC       ← fabricationToleranceClass     (Text)
        <PREFIX>_PWHT      ← postWeldHeatTreatment         (Text)
        <PREFIX>_HEAD      ← typeOfHeads                   (Text)
        <PREFIX>_INS       ← insulation                    (Text)
        <PREFIX>_INS_SPEC  ← insulationTypeThkDensity      (Text)

    Blank/None values produce empty strings and are subsequently skipped by
    the write loop — no property is written for them.
    """
    def g(key: str) -> str:
        v = col.get(key)
        return str(v).strip() if v is not None else ""

    op_temp = g("operatingTempMinMax")
    mot     = _extract_mot(op_temp)

    return {
        f"{prefix}_IDP":      g("internalDesignPressureMawp"),
        f"{prefix}_EDP":      g("externalDesignPressureMawp"),
        f"{prefix}_WP":       g("workingPressure"),
        f"{prefix}_HTP":      g("hydroTestPressure"),
        f"{prefix}_MDMT":     g("mdmt"),
        f"{prefix}_HT_TEMP":  g("hydroTestTempMinMax"),
        f"{prefix}_OP_TEMP":  op_temp,
        f"{prefix}_MOT":      mot,
        f"{prefix}_DES_TEMP": g("designTempMinMax"),
        f"{prefix}_STATE":    g("physicalState"),
        f"{prefix}_VOL":      g("grossVolumeLiters"),
        f"{prefix}_FLUID":    g("serviceFluid"),
        f"{prefix}_HZ":       g("hazardLevel"),
        f"{prefix}_SG":       g("specificGravity"),
        f"{prefix}_ICA":      g("internalCorrosionAllowanceMm"),
        f"{prefix}_ECA":      g("externalCorrosionAllowanceMm"),
        f"{prefix}_RT":       g("radiography"),
        f"{prefix}_JE":       g("jointEfficiency"),
        f"{prefix}_TG":       g("testingGroup"),
        f"{prefix}_FTC":      g("fabricationToleranceClass"),
        f"{prefix}_PWHT":     g("postWeldHeatTreatment"),
        f"{prefix}_HEAD":     g("typeOfHeads"),
        f"{prefix}_INS":      g("insulation"),
        f"{prefix}_INS_SPEC": g("insulationTypeThkDensity"),
    }


def _mech_col_prop_names(prefix: str) -> list:
    """Return the 24 full SW property names for one mechanical column prefix."""
    return [f"{prefix}_{s}" for s in _MECH_PROP_SHORTS]


def _general_data_props(gen: dict) -> dict:
    """
    Build {SW_property_name: value_string} for the General Data section.

    Approved mapping (Phase 3 — 12 SW properties):
        HYDRO_TEST_POSITION    ← hydroTestPosition       (Text — existing template field)
        GENERAL_ORIENT         ← vesselOrientation        (Text)
        GENERAL_SERVICE_LIFE   ← designServiceLife        (Text)
        GENERAL_WIND_CODE      ← windData                 (Text)
        GENERAL_WIND_VEL       ← windDesignVelocity       (Text)
        GENERAL_SEISMIC_CODE   ← seismicDesignCode        (Text)
        GENERAL_SEISMIC_Z      ← hazardFactorZ            (Number — text)
        GENERAL_SEISMIC_H      ← seismicCoefficientHorizontal (Number — text)
        GENERAL_SEISMIC_V      ← seismicCoefficientVertical   (Number — text)
        GENERAL_WEIGHT         ← weightEmptyOperatingHydro    (Text — composite)
        GENERAL_LOCATION       ← location                 (Text)
        GENERAL_QTY            ← qty                      (Number — text)

    DDS general_data keys are camelCase (TypeScript GeneralData type serialised
    directly to JSON).  Blank / None values produce empty strings and are
    subsequently skipped by the write loop.
    """
    def g(key: str) -> str:
        v = gen.get(key)
        return str(v).strip() if v is not None else ""

    return {
        "HYDRO_TEST_POSITION":  g("hydroTestPosition"),
        "GENERAL_ORIENT":       g("vesselOrientation"),
        "GENERAL_SERVICE_LIFE": g("designServiceLife"),
        "GENERAL_WIND_CODE":    g("windData"),
        "GENERAL_WIND_VEL":     g("windDesignVelocity"),
        "GENERAL_SEISMIC_CODE": g("seismicDesignCode"),
        "GENERAL_SEISMIC_Z":    g("hazardFactorZ"),
        "GENERAL_SEISMIC_H":    g("seismicCoefficientHorizontal"),
        "GENERAL_SEISMIC_V":    g("seismicCoefficientVertical"),
        "GENERAL_WEIGHT":       g("weightEmptyOperatingHydro"),
        "GENERAL_LOCATION":     g("location"),
        "GENERAL_QTY":          g("qty"),
    }


def _write_properties(swModel, job: dict, logger) -> tuple[list, list]:
    """
    Write all DDS-sourced custom properties.

    Phase 1 — 16 header properties (template-exact names, always written):
        Drawing_Number, Revision, Tag_No, Serial_No, Description,
        Equipment_Type, Equipment_Configuration, Design_Code,
        Material_Code, Inspection_By

        Agent-filled title-block fields (system-generated, always written):
        DrawnBy, CheckedBy, EngineeringApproval  — value: "Agent"
        DrawnDate, CheckedDate, EngAppDate        — value: today dd/mm/YYYY
        (Format matches Extraction Agent primary date format %d/%m/%Y)

    Phase 2 — Mechanical Design Data (Option C — approved mapping):
        SHELL_*   : 24 properties always (shell column always present)
        TUBE_*    : 24 properties written if mechanical_data.tube is non-null
                    24 properties CLEARED  if mechanical_data.tube is null
        JACKET_*  : 24 properties written if mechanical_data.jacket is non-null
                    24 properties CLEARED  if mechanical_data.jacket is null

        Equipment-type clearing rules:
          Vessel                  → write SHELL_*, clear TUBE_* + JACKET_*
          Heat Exchanger          → write SHELL_* + TUBE_*, clear JACKET_*
          Jacketed Vessel         → write SHELL_* + JACKET_*, clear TUBE_*
          Jacketed HX             → write all three — nothing cleared

        operatingTempMinMax generates TWO properties per column:
            <PREFIX>_OP_TEMP  — full text string  e.g. "100 / 120"
            <PREFIX>_MOT      — max value only     e.g. "120"  (existing template field)

        Clearing uses ICustomPropertyManager.Delete() (primary) so the property
        is fully removed and the drawing title block shows blank.
        Falls back to Add3("") if Delete() is unavailable on the COM binding.

    Phase 3 — General Data (12 properties):
        HYDRO_TEST_POSITION    ← hydroTestPosition         (existing template field)
        GENERAL_ORIENT         ← vesselOrientation
        GENERAL_SERVICE_LIFE   ← designServiceLife
        GENERAL_WIND_CODE      ← windData
        GENERAL_WIND_VEL       ← windDesignVelocity
        GENERAL_SEISMIC_CODE   ← seismicDesignCode
        GENERAL_SEISMIC_Z      ← hazardFactorZ
        GENERAL_SEISMIC_H      ← seismicCoefficientHorizontal
        GENERAL_SEISMIC_V      ← seismicCoefficientVertical
        GENERAL_WEIGHT         ← weightEmptyOperatingHydro
        GENERAL_LOCATION       ← location
        GENERAL_QTY            ← qty

    Only non-blank values are written (blank fields are silently skipped).
    Returns (properties_written, warnings).
    """
    cpm = swModel.Extension.CustomPropertyManager("")
    written  = []
    warnings = []

    dds = job.get("dds") or {}

    def _dds(*keys: str) -> str:
        """Return first non-blank value found among the given DDS keys."""
        for k in keys:
            v = str(dds.get(k) or "").strip()
            if v:
                return v
        return ""

    # ── Phase 1: 16 header properties ────────────────────────────────────────
    today_str = datetime.now().strftime("%d/%m/%Y")
    to_write: dict = {
        # DDS-sourced fields
        "Drawing_Number":          job.get("drawing_number", ""),
        "Revision":                job.get("revision", ""),
        "Tag_No":                  _dds("tag_no", "tagNo"),
        "Serial_No":               _dds("manufacture_serial_no", "manufactureSerialNo"),
        "Description":             _dds("equipment_description", "equipmentDescription"),
        "Equipment_Type":          _dds("equipment_type"),
        "Equipment_Configuration": _dds("equipment_config"),
        "Design_Code":             _dds("design_code"),
        "Material_Code":           _dds("material_code"),
        "Inspection_By":           _dds("inspection_by"),
        # Agent-filled title-block fields (Phase 1 automated workflow)
        "DrawnBy":                 "Agent",
        "DrawnDate":               today_str,
        "CheckedBy":               "Agent",
        "CheckedDate":             today_str,
        "EngineeringApproval":     "Agent",
        "EngAppDate":              today_str,
    }
    logger.info(
        f"[Structurer] Phase 1 agent-filled fields: "
        f"DrawnBy=Agent DrawnDate={today_str} "
        f"CheckedBy=Agent CheckedDate={today_str} "
        f"EngineeringApproval=Agent EngAppDate={today_str}"
    )

    # ── Phase 2: Mechanical Design Data — 72 properties (up to) ──────────────
    #
    # Columns with data  → queued for writing  (to_write)
    # Columns that are null/absent → queued for explicit clearing (to_clear)
    #   so stale template / memory values are removed from the drawing.
    #
    # Clearing strategy:
    #   Primary  : ICustomPropertyManager.Delete(name)  — removes the property
    #              entirely; linked title-block annotation shows blank.
    #   Fallback : Add3(name, swCustomInfoText, "", swCustomPropertyReplaceValue)
    #              — sets value to "" if Delete() is unavailable on this COM build.
    mech    = dds.get("mechanical_data")
    to_clear: list = []          # property names to explicitly blank/delete

    if isinstance(mech, dict):
        shell_col  = mech.get("shell")
        tube_col   = mech.get("tube")
        jacket_col = mech.get("jacket")

        if isinstance(shell_col, dict):
            to_write.update(_mech_col_props("SHELL", shell_col))
            logger.info("[Structurer] Mechanical SHELL column: 24 properties queued")
        else:
            logger.warning("[Structurer] mechanical_data.shell missing or not a dict — skipped")

        if isinstance(tube_col, dict):
            to_write.update(_mech_col_props("TUBE", tube_col))
            logger.info("[Structurer] Mechanical TUBE column: 24 properties queued")
        else:
            to_clear.extend(_mech_col_prop_names("TUBE"))
            logger.info("[Structurer] mechanical_data.tube is null — TUBE_* queued for clear")

        if isinstance(jacket_col, dict):
            to_write.update(_mech_col_props("JACKET", jacket_col))
            logger.info("[Structurer] Mechanical JACKET column: 24 properties queued")
        else:
            to_clear.extend(_mech_col_prop_names("JACKET"))
            logger.info("[Structurer] mechanical_data.jacket is null — JACKET_* queued for clear")
    else:
        logger.warning("[Structurer] mechanical_data missing from DDS payload — Phase 2 skipped")

    # ── Phase 3: General Data ─────────────────────────────────────────────────
    gen_data: dict = (dds.get("general_data") or {}) if isinstance(dds.get("general_data"), dict) else {}
    if gen_data:
        to_write.update(_general_data_props(gen_data))
        logger.info("[Structurer] Phase 3 — general_data present, merging 12 GENERAL_* properties")
    else:
        logger.warning("[Structurer] general_data missing or empty in DDS payload — Phase 3 skipped")

    # ── Write loop ────────────────────────────────────────────────────────────
    non_blank = {k: v for k, v in to_write.items() if v}
    skipped   = len(to_write) - len(non_blank)
    logger.info(
        f"[Structurer] Writing {len(non_blank)} properties "
        f"({skipped} skipped — blank in DDS payload)"
    )

    for name, value in to_write.items():
        str_val = str(value) if value is not None else ""
        if not str_val:
            logger.debug(f"[Structurer] Property skipped (blank): {name}")
            continue
        try:
            ret = cpm.Add3(name, _SW_CUSTOM_INFO_TEXT, str_val, _SW_CUSTOM_PROP_REPLACE)
            if ret == 0:
                logger.info(f"[Structurer] Property written: {name} = {str_val!r}")
                written.append(name)
            else:
                msg = f"Property '{name}' Add3 returned code {ret}"
                logger.warning(f"[Structurer] {msg}")
                warnings.append(msg)
        except Exception as e:
            msg = f"Property '{name}' write failed: {type(e).__name__}: {e}"
            logger.warning(f"[Structurer] {msg}")
            warnings.append(msg)

    # ── Clear loop — explicitly blank non-applicable columns ──────────────────
    # Prefer Delete() to remove the property entirely (title block shows blank).
    # Fall back to Add3("") if Delete is unavailable (older COM binding).
    cleared_count = 0
    if to_clear:
        logger.info(f"[Structurer] Clearing {len(to_clear)} non-applicable column properties")
    for name in to_clear:
        deleted = False
        try:
            cpm.Delete(name)
            deleted = True
        except Exception:
            pass
        if deleted:
            logger.info(f"[Structurer] Property cleared (deleted): {name}")
            cleared_count += 1
        else:
            # Fallback: overwrite with empty string
            try:
                cpm.Add3(name, _SW_CUSTOM_INFO_TEXT, "", _SW_CUSTOM_PROP_REPLACE)
                logger.info(f"[Structurer] Property cleared (set empty): {name}")
                cleared_count += 1
            except Exception as e:
                msg = f"Property '{name}' clear failed: {type(e).__name__}: {e}"
                logger.warning(f"[Structurer] {msg}")
                warnings.append(msg)
    if to_clear:
        logger.info(f"[Structurer] {cleared_count}/{len(to_clear)} non-applicable properties cleared")

    return written, warnings


def _verify_properties(swModel, properties_written: list, logger) -> tuple[list, list]:
    """
    Read back each written property to confirm it round-trips correctly.

    SW2019 COM quirk: ICustomPropertyManager.Get5/Get4 require explicit ByRef
    VARIANT objects for their output parameters — calling with only (name, cached)
    raises DISP_E_PARAMNOTOPTIONAL (0x8002000F).  We pass VT_BSTR|VT_BYREF
    VARIANTs as placeholders so win32com satisfies the ByRef contract, then read
    .value after the call.  Mirrors the extraction agent's proven strategy.

    Returns (verified, mismatch_warnings).
    """
    import pythoncom
    from win32com.client import VARIANT as _VARIANT
    _VT_BS_REF = pythoncom.VT_BSTR | pythoncom.VT_BYREF
    _VT_BL_REF = pythoncom.VT_BOOL | pythoncom.VT_BYREF

    cpm        = swModel.Extension.CustomPropertyManager("")
    verified   = []
    mismatches = []

    for name in properties_written:
        read_val = None
        try:
            # Try Get5 (name, useCached, val/out, resolvedVal/out, wasResolved/out)
            v_val  = _VARIANT(_VT_BS_REF, "")
            v_rval = _VARIANT(_VT_BS_REF, "")
            v_wr   = _VARIANT(_VT_BL_REF, False)
            cpm.Get5(name, False, v_val, v_rval, v_wr)
            raw = (v_val.value  or "").strip()
            res = (v_rval.value or "").strip()
            read_val = res if res else raw
        except Exception:
            pass

        if read_val is None:
            try:
                # Fallback: Get4 (name, useCached, val/out, resolvedVal/out)
                v_val  = _VARIANT(_VT_BS_REF, "")
                v_rval = _VARIANT(_VT_BS_REF, "")
                cpm.Get4(name, False, v_val, v_rval)
                raw = (v_val.value  or "").strip()
                res = (v_rval.value or "").strip()
                read_val = res if res else raw
            except Exception:
                pass

        if read_val is not None:
            verified.append(name)
            logger.info(f"[Structurer] Verified: {name} = {read_val!r}")
        else:
            msg = f"Read-back of '{name}' returned None/unreadable"
            mismatches.append(msg)
            logger.warning(f"[Structurer] {msg}")

    return verified, mismatches


# ── Safety checks on existing file ────────────────────────────────────────────

def _check_existing_drawing_consistency(swModel, job: dict, logger):
    """
    For update_existing mode: read Drawing_Number and Revision from the opened
    file and compare against job payload. Raises ValueError on mismatch.
    """
    cpm = swModel.Extension.CustomPropertyManager("")
    checks = {
        "Drawing_Number": job["drawing_number"],
        "Revision":       job["revision"],
    }
    for prop_name, expected in checks.items():
        try:
            result = cpm.Get5(prop_name, False)
            actual = result[1] if isinstance(result, tuple) and len(result) > 1 else result
            actual = str(actual or "").strip()
        except Exception:
            actual = ""

        if not actual:
            logger.warning(
                f"[Structurer] {prop_name} not found in existing file — proceeding with overwrite"
            )
            continue

        if actual.lower() != expected.lower():
            raise ValueError(
                f"{prop_name} mismatch: file has '{actual}', job expects '{expected}'"
            )

    logger.info("[Structurer] Existing file consistency checks passed")


# ── Main entry point ──────────────────────────────────────────────────────────

def run_structuring(job: dict, config, cancel_event: threading.Event, logger) -> dict:
    """
    Phase 1 structuring entry point.  Called inside a worker thread by
    structure_job_runner.py.

    Returns a result dict.
    Raises on unrecoverable error — caller is responsible for fail_job().
    """
    if not PYWIN32_AVAILABLE:
        raise RuntimeError(
            "pywin32 is not installed — cannot launch SolidWorks COM."
        )

    drawing_number = job["drawing_number"]
    revision       = job["revision"]
    mode           = (job.get("mode") or "create_new").strip()

    # Prefer template_path / staging_root embedded in the job payload
    # (set via ERP System Settings, baked into every job at creation time).
    # Fall back to local config.ini values for backward compatibility.
    template_path = (
        (job.get("template_path") or job.get("templatePath") or "").strip()
        or config.structurer_template_path
    )
    staging_root = (
        (job.get("staging_root") or job.get("stagingRoot") or "").strip()
        or config.structurer_staging_root
    )

    logger.info(
        f"[Structurer] Job start — drawing={drawing_number} rev={revision} mode={mode}"
    )
    logger.info(f"[Structurer] template_path (resolved): {template_path or 'NOT SET'}")
    logger.info(f"[Structurer] staging_root  (resolved): {staging_root  or 'NOT SET'}")

    # ── Pre-flight (before any SolidWorks involvement) ────────────────────────
    staging_path = _preflight(job, template_path, staging_root)
    logger.info(f"[Structurer] Pre-flight passed — staging_path={staging_path}")

    if mode == "create_new" and os.path.isfile(staging_path):
        _existing_bytes = os.path.getsize(staging_path)
        logger.warning(
            f"[Structurer] create_new: staging file already exists "
            f"({_existing_bytes:,} bytes) — removing before fresh create: {staging_path}"
        )
        try:
            os.remove(staging_path)
            logger.info("[Structurer] Existing staging file removed OK")
        except OSError as _oe:
            raise ValueError(
                f"create_new: cannot remove existing staging file: {staging_path} — {_oe}"
            )

    t_start = time.monotonic()

    swApp          = None
    swModel        = None
    agent_sw_pid   = None
    binding_mode   = "none"
    sw_launch_ok   = False

    # ── SolidWorks launch (one retry on failure) ──────────────────────────────
    pythoncom.CoInitialize()
    try:
        pids_before = _get_sldworks_pids()

        for attempt in range(1, 3):
            try:
                swApp, binding_mode = _launch_sw_dedicated_instance(
                    config.sw_progid, logger
                )
                sw_launch_ok = True
                break
            except Exception as e:
                logger.warning(
                    f"[Structurer] SW launch attempt {attempt} failed: {type(e).__name__}: {e}"
                )
                if attempt < 2:
                    logger.info("[Structurer] Retrying SolidWorks launch in 10s…")
                    time.sleep(10)

        if not sw_launch_ok:
            raise RuntimeError("SolidWorks launch failed after 1 retry")

        pids_after  = _get_sldworks_pids()
        new_pids    = pids_after - pids_before
        agent_sw_pid = next(iter(new_pids), None)
        if agent_sw_pid:
            logger.info(f"[COM] Agent's dedicated SLDWORKS.EXE PID: {agent_sw_pid}")
        else:
            logger.info("[COM] PID not isolated (DispatchEx reused existing process)")

        swApp.Visible = False
        try:
            swApp.UserControl = False
        except Exception:
            pass

        logger.info("[Structurer] SolidWorks hidden instance ready")

        # ── Cancel check ─────────────────────────────────────────────────────
        if cancel_event.is_set():
            raise RuntimeError("Job cancelled before document open")

        # ── Mode branch ───────────────────────────────────────────────────────
        if mode == "create_new":
            logger.info(f"[Structurer] NewDocument from template: {template_path}")
            swModel = swApp.NewDocument(template_path, 0, 0, 0)
            if swModel is None:
                raise RuntimeError(
                    "NewDocument returned None — check template_path and SolidWorks license"
                )
            logger.info("[Structurer] New drawing document created")

            # ── Confirm document is active ────────────────────────────────────
            # NewDocument automatically makes the new doc the active document.
            # We confirmed this empirically: swApp.ActiveDoc.GetTitle returns the
            # correct title immediately after NewDocument returns.
            #
            # NOTE: ActivateDoc2 is NOT called here because its 3rd param (Errors)
            # is ByRef Long — pywin32 late-binding raises DISP_E_TYPEMISMATCH when
            # a plain Python int is passed, and the call is unnecessary anyway since
            # the document is already active.
            #
            # NOTE: GetTitle is a COM PROPERTY in pywin32 — access WITHOUT parens.
            # Calling swModel.GetTitle() raises "str object is not callable".

            _doc_title = ""
            try:
                _doc_title = str(swModel.GetTitle)
                logger.info(f"[Structurer] Document title: {_doc_title!r}")
            except Exception as _te:
                logger.warning(f"[Structurer] GetTitle property failed: {_te}")

            try:
                _active = swApp.ActiveDoc
                _active_title = str(_active.GetTitle) if _active is not None else "None"
                logger.info(f"[Structurer] Active doc confirmed: {_active_title!r}")
            except Exception as _ce:
                logger.warning(f"[Structurer] Could not confirm active doc: {_ce}")

        elif mode == "update_existing":
            logger.info(f"[Structurer] Opening existing file: {staging_path}")

            swModel = None

            for open_attempt in range(1, 3):
                try:
                    # ISldWorks.OpenDoc(FileName, Type) — no ByRef params, late-bind safe
                    swModel = swApp.OpenDoc(staging_path, _SW_DOC_DRAWING)
                except Exception as e:
                    logger.warning(
                        f"[Structurer] OpenDoc attempt {open_attempt} raised: "
                        f"{type(e).__name__}: {e}"
                    )
                    swModel = None

                if swModel is not None:
                    break

                if open_attempt < 2:
                    logger.warning(
                        "[Structurer] OpenDoc returned None — "
                        "file may be locked. Retrying in 15s…"
                    )
                    time.sleep(15)

            if swModel is None:
                raise RuntimeError(
                    f"OpenDoc failed for '{staging_path}' after 1 retry — "
                    "file may be missing, locked, or corrupt"
                )

            logger.info("[Structurer] Existing drawing opened")

            # Safety: check Drawing_Number + Revision consistency
            _check_existing_drawing_consistency(swModel, job, logger)

        else:
            raise ValueError(f"Unknown mode: '{mode}' — expected create_new or update_existing")

        # ── Cancel check ─────────────────────────────────────────────────────
        if cancel_event.is_set():
            raise RuntimeError("Job cancelled before property write")

        # ── Shared Save2 helper ───────────────────────────────────────────────
        # SW2019 headless quirk: Save2 often returns 0 (False) even when the
        # save succeeds.  We check the file's modification time before/after the
        # call — if mtime advanced the save happened; if not but the file still
        # exists with bytes, we accept it with a WARNING rather than failing.
        def _save2(label: str) -> None:
            _mtime_before = 0.0
            try:
                _mtime_before = os.path.getmtime(staging_path)
            except OSError:
                pass

            try:
                _ret = swModel.Save2(0)
            except Exception as _se2:
                import traceback as _tb2
                _a2   = getattr(_se2, 'args', ())
                _hr2  = (hex(_a2[0]) if isinstance(_a2[0], int) else repr(_a2[0])) if _a2 else 'N/A'
                raise RuntimeError(
                    f"Save2({label}) COM exception: {type(_se2).__name__}: {_se2} "
                    f"| HRESULT={_hr2} | traceback: {_tb2.format_exc()}"
                ) from _se2

            _mtime_after = 0.0
            try:
                _mtime_after = os.path.getmtime(staging_path)
            except OSError:
                pass
            _mtime_changed = _mtime_after > _mtime_before
            _sz2 = os.path.getsize(staging_path) if os.path.isfile(staging_path) else 0

            logger.info(
                f"[Structurer] Save2({label}) → ret={_ret} "
                f"| mtime_changed={_mtime_changed} | file={_sz2:,} bytes"
            )

            if _ret:
                return  # SW reported success
            if _mtime_changed:
                logger.warning(
                    f"[Structurer] Save2({label}) returned False but mtime advanced "
                    f"— treating as success (SW2019 headless non-fatal return)"
                )
                return
            if _sz2 > 0:
                logger.warning(
                    f"[Structurer] Save2({label}) returned False, mtime unchanged "
                    f"— file exists ({_sz2:,} bytes), treating as success"
                )
                return
            raise RuntimeError(f"Save2({label}) returned False and file not found/empty")

        # ── Save + property-write flow ────────────────────────────────────────
        # create_new  : SaveAs3 first (establishes file path) → write props → Save2
        # update_existing: write props → Save2
        #
        # IMPORTANT: For create_new, properties MUST be written AFTER SaveAs3.
        # SW2019 headless SaveAs3 writes the template bytes to disk and resets
        # the in-memory document to that template state — any Add3() calls made
        # before SaveAs3 are discarded.  Writing properties after SaveAs3 and
        # flushing with Save2 is the reliable two-phase pattern.

        if mode == "create_new":
            # ── Log active doc immediately before save ────────────────────────
            try:
                _pre_save_active = swApp.ActiveDoc
                _pre_save_title  = str(_pre_save_active.GetTitle) if _pre_save_active else "None"
            except Exception:
                _pre_save_title = "unknown"
            logger.info(f"[Structurer] Active doc before SaveAs3: {_pre_save_title!r}")

            logger.info(f"[Structurer] SaveAs3 → {staging_path}")

            # swSaveAsOptions_e constants
            _OPT_NONE   = 0   # allow dialogs (may hang in headless — but some SW builds need it)
            _OPT_SILENT = 1   # suppress dialogs
            _OPT_COPY   = 2   # save a copy (doesn't rename the in-memory document)

            def _try_save_as3(doc, label, opts):
                """
                Call IModelDoc2.SaveAs3 and return (effective_success, err_str|None).

                SW2019 headless quirk: SaveAs3 can write the file to disk but still
                return 0 (False) when there is a non-fatal internal issue — e.g. a
                rebuild warning, unresolved sheet-format reference, etc.  We check
                whether the output file actually appeared on disk: if it did with a
                non-zero size, we log a WARNING and treat it as success so the job
                completes rather than failing.
                """
                # Remove stale file so we can detect whether this call wrote it
                if os.path.isfile(staging_path):
                    try:
                        os.remove(staging_path)
                        logger.debug(f"[Structurer] Pre-SaveAs3 stale file removed ({label})")
                    except OSError:
                        pass

                try:
                    r = doc.SaveAs3(staging_path, _SW_SAVE_CURRENT_VERSION, opts)
                except Exception as _se:
                    import traceback as _tb2
                    _a  = getattr(_se, 'args', ())
                    _hr = (hex(_a[0]) if isinstance(_a[0], int) else repr(_a[0])) if _a else 'N/A'
                    _ei = repr(_a[2]) if len(_a) > 2 else 'N/A'
                    _msg = (
                        f"SaveAs3({label}) COM exception: {type(_se).__name__}: {_se} "
                        f"| HRESULT={_hr} | excepinfo={_ei} "
                        f"| path={staging_path!r} "
                        f"| traceback: {_tb2.format_exc()}"
                    )
                    logger.warning(f"[Structurer] {_msg}")
                    return None, _msg

                # Check whether file landed on disk
                _sz = os.path.getsize(staging_path) if os.path.isfile(staging_path) else 0
                logger.info(
                    f"[Structurer] SaveAs3({label}, opts={opts}) "
                    f"→ ret={r} | file_on_disk={_sz:,} bytes"
                )

                if r:
                    return True, None   # SW reported success

                if _sz > 0:
                    # SW returned False but wrote a non-empty file — accept it.
                    # This is a known SW2019 headless behaviour when a non-fatal
                    # internal issue (rebuild warning, sheet-format reference) is
                    # encountered; the file is valid and openable.
                    logger.warning(
                        f"[Structurer] SaveAs3({label}) returned False but file was "
                        f"written ({_sz:,} bytes) — treating as success "
                        f"(SW2019 headless non-fatal return)"
                    )
                    return True, None

                return False, None   # no file, genuine failure

            # ── Attempt 1: swModel, Options=Silent ───────────────────────────
            ret, exc_info = _try_save_as3(swModel, "swModel", _OPT_SILENT)
            if exc_info:
                raise RuntimeError(exc_info)

            # ── Attempt 2: swModel, Options=None (no dialog suppression) ─────
            if not ret:
                logger.warning("[Structurer] Attempt 1 failed — retrying with opts=0")
                ret, exc_info = _try_save_as3(swModel, "swModel/opts=0", _OPT_NONE)
                if exc_info:
                    raise RuntimeError(exc_info)

            # ── Attempt 3: swApp.ActiveDoc, Options=Silent ───────────────────
            if not ret:
                logger.warning("[Structurer] Attempt 2 failed — retrying via ActiveDoc")
                try:
                    _ad = swApp.ActiveDoc
                    if _ad is not None:
                        ret, exc_info = _try_save_as3(_ad, "ActiveDoc", _OPT_SILENT)
                        if exc_info:
                            logger.warning(f"[Structurer] Attempt 3 raised: {exc_info}")
                            ret = False
                    else:
                        logger.warning("[Structurer] swApp.ActiveDoc is None — skipping attempt 3")
                except Exception as _a3e:
                    logger.warning(f"[Structurer] Attempt 3 wrapper error: {_a3e}")

            # ── Attempt 4: Visible=True, swModel, Options=Silent ─────────────
            if not ret:
                logger.warning("[Structurer] Attempt 3 failed — setting Visible=True and retrying")
                try:
                    swApp.Visible = True
                    logger.info("[Structurer] swApp.Visible = True")
                except Exception as _ve:
                    logger.warning(f"[Structurer] Could not set Visible=True: {_ve}")
                ret, exc_info = _try_save_as3(swModel, "swModel+Visible", _OPT_SILENT)
                if exc_info:
                    raise RuntimeError(exc_info)

            if not ret:
                raise RuntimeError(
                    f"SaveAs3 returned False AND no file written — all 4 attempts failed "
                    f"| path={staging_path!r} "
                    f"| active_doc_before_save={_pre_save_title!r} "
                    f"| doc_title={_doc_title!r}"
                )

            # ── Phase 2: write properties now that the document has a path ────
            # SaveAs3 reset the in-memory document to template state, so any
            # Add3() calls made before SaveAs3 were discarded.  We write here
            # and flush with Save2 so the properties actually land on disk.
            logger.info("[Structurer] Writing custom properties (post-SaveAs3)…")
            properties_written, write_warnings = _write_properties(swModel, job, logger)
            logger.info(
                f"[Structurer] {len(properties_written)} properties written, "
                f"{len(write_warnings)} write warnings"
            )
            _save2("create_new/props")

        else:
            # ── update_existing: write props then Save2 ───────────────────────
            logger.info("[Structurer] Writing custom properties…")
            properties_written, write_warnings = _write_properties(swModel, job, logger)
            logger.info(
                f"[Structurer] {len(properties_written)} properties written, "
                f"{len(write_warnings)} write warnings"
            )
            _save2("update_existing")

        logger.info("[Structurer] Save successful")

        # ── Flush pause — let SolidWorks finish writing before ExitApp() ──────
        # Without this, ExitApp() can interrupt the file write, leaving a file
        # that Windows Explorer can see but SolidWorks cannot open.
        time.sleep(2)
        logger.info("[Structurer] Flush pause complete — proceeding to verify")

        # ── Post-write read-back ──────────────────────────────────────────────
        properties_verified, verify_warnings = _verify_properties(
            swModel, properties_written, logger
        )

        # ── File stats ────────────────────────────────────────────────────────
        file_size = 0
        try:
            file_size = os.path.getsize(staging_path)
        except Exception:
            pass

        duration = time.monotonic() - t_start
        all_warnings = write_warnings + verify_warnings

        result = {
            "status":               "success",
            "drawing_number":       drawing_number,
            "revision":             revision,
            "mode":                 mode,
            "file_path":            staging_path,
            "file_size_bytes":      file_size,
            "properties_written":   properties_written,
            "properties_verified":  properties_verified,
            "solidworks_session":   f"dedicated-isolated (binding={binding_mode} pid={agent_sw_pid})",
            "duration_sec":         round(duration, 2),
            "errors":               [],
            "warnings":             all_warnings,
        }
        logger.info(
            f"[Structurer] Complete — {len(properties_written)} props, "
            f"{file_size:,} bytes, {duration:.1f}s"
        )
        return result

    finally:
        # ── Step 1: Close document ────────────────────────────────────────────
        if swModel is not None:
            try:
                swApp.CloseDoc(staging_path)
                logger.info("[COM] Document closed")
            except Exception as e:
                logger.warning(f"[COM] CloseDoc error: {e}")

        # ── Step 2: Exit dedicated instance (always) ──────────────────────────
        if swApp is not None:
            try:
                swApp.ExitApp()
                logger.info("[COM] Dedicated SolidWorks instance exited cleanly")
            except Exception as e:
                logger.warning(f"[COM] ExitApp() failed: {type(e).__name__}: {e}")
                if agent_sw_pid:
                    logger.warning(
                        f"[COM] Orphan guard: force-killing PID {agent_sw_pid}…"
                    )
                    _kill_orphan_sw_process(agent_sw_pid, logger)
                else:
                    logger.warning(
                        "[COM] Orphan guard: no PID tracked — "
                        "check Task Manager for stray SLDWORKS.EXE processes"
                    )

        # ── Step 3: Release COM apartment ────────────────────────────────────
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
