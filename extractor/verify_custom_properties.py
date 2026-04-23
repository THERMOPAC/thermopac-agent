"""
verify_custom_properties.py — Layer 1: Custom Property Verification

Validates the 21 target custom properties extracted from a SolidWorks drawing.
No DDS table comparison.  Rules are applied purely to the property values themselves.

Input:
    cp_extraction — dict returned by _extract_custom_properties():
        {
            "resolved":    {"prop": {"value": str, "source": str}, ...},
            "bySource":    {"drawing": {...}, "sheet": {...}, "model": {...}},
            "allDetected": {"drawing": [...], "sheet": [...], "model": [...]},
            "totalFound":  int,
        }

Output (returned):
    {
        "customPropertyVerification": {
            "status":          "pass" | "fail" | "hold",
            "equipmentConfig": str,
            "fields": [
                {
                    "property":        str,
                    "source":          "drawing" | "sheet" | "model" | "none",
                    "applicability":   "required" | "not_applicable",
                    "value":           str,
                    "normalizedValue": str,
                    "result":          "pass" | "fail" | "hold" | "not_applicable",
                    "reason":          str,
                },
                ...
            ],
        }
    }
"""

from __future__ import annotations
import re
from datetime import datetime, date


# ─── Constants ────────────────────────────────────────────────────────────────

_ALLOWED_CONFIGS = {
    "vessel",
    "jacketed vessel",
    "heat exchanger",
    "jacketed vessel and heat exchanger",
}

_BLANK_VALUES = frozenset(["", "-", "—", "n/a", "na", "none", "null"])

_DATE_FORMATS = [
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
    "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d",
    "%d/%m/%y", "%d-%m-%y",
    "%m/%d/%Y", "%m-%d-%Y",
    "%d %b %Y", "%d-%b-%Y", "%d %B %Y",
    "%b %d, %Y", "%B %d, %Y",
]


# ─── Normalisation helpers ────────────────────────────────────────────────────

def _norm_text(v: str) -> str:
    """Trim + collapse whitespace + lowercase."""
    return re.sub(r"\s+", " ", str(v or "").strip()).lower()


def _norm_numeric(v: str) -> str:
    """Normalise numeric string: strip trailing zeros. Returns '' if not numeric."""
    try:
        f = float(v.replace(",", ""))
        if f == int(f):
            return str(int(f))
        return f"{f:g}"
    except (ValueError, AttributeError):
        return ""


def _parse_date(raw: str) -> date | None:
    s = str(raw or "").strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _is_blank(v: str | None) -> bool:
    return v is None or _norm_text(v) in _BLANK_VALUES


# ─── Equipment configuration → active column map ─────────────────────────────

def _active_columns(config_norm: str) -> tuple[bool, bool]:
    """Return (tube_active, jacket_active) for a normalised config string."""
    if config_norm == "jacketed vessel and heat exchanger":
        return True, True
    if config_norm == "heat exchanger":
        return True, False
    if config_norm == "jacketed vessel":
        return False, True
    if config_norm == "vessel":
        return False, False
    return False, False


# ─── Field builders ───────────────────────────────────────────────────────────

def _field(
    prop: str,
    source: str,
    applicability: str,
    value: str,
    normalized_value: str,
    result: str,
    reason: str,
) -> dict:
    return {
        "property":        prop,
        "source":          source,
        "applicability":   applicability,
        "value":           value,
        "normalizedValue": normalized_value,
        "result":          result,
        "reason":          reason,
    }


def _not_applicable(prop: str, source: str, value: str) -> dict:
    return _field(
        prop, source, "not_applicable", value, "",
        "not_applicable",
        "Not required for selected Equipment_Configuration",
    )


def _not_applicable_contaminated(prop: str, source: str, value: str, config: str) -> dict:
    """
    Property is not applicable for the current Equipment_Configuration,
    but a value is present in the drawing — this is a data integrity failure.
    applicability stays 'not_applicable' to communicate why the rule fired,
    but result is 'fail' so it enters overall scoring and blocks pass.
    """
    return _field(
        prop, source, "not_applicable", value, "",
        "fail",
        f"Value {value!r} is present but this property is not applicable for "
        f"Equipment_Configuration={config!r}. Remove this value from the drawing.",
    )


def _missing(prop: str) -> dict:
    return _field(prop, "none", "required", "", "", "hold", "Property missing or blank")


def _required_pass(prop: str, source: str, value: str, norm: str = "") -> dict:
    return _field(prop, source, "required", value, norm or value, "pass", "")


def _required_fail(prop: str, source: str, value: str, reason: str) -> dict:
    return _field(prop, source, "required", value, "", "fail", reason)


def _required_hold(prop: str, source: str, value: str, reason: str) -> dict:
    return _field(prop, source, "required", value, "", "hold", reason)


# ─── Main entry point ─────────────────────────────────────────────────────────

def verify_custom_properties(cp_extraction: dict, logger=None) -> dict:
    """
    Verify custom properties against Layer 1 rules.

    Args:
        cp_extraction : dict from _extract_custom_properties()
        logger        : optional logger; falls back to no-op

    Returns:
        {"customPropertyVerification": {"status": str, "equipmentConfig": str, "fields": list}}
    """

    def _log(msg: str) -> None:
        if logger:
            try:
                logger.info(msg)
            except Exception:
                pass

    _log("[CPVerify] Layer 1 verification starting")

    resolved: dict[str, dict] = (cp_extraction or {}).get("resolved") or {}

    def _val(prop: str) -> str:
        return str((resolved.get(prop) or {}).get("value") or "").strip()

    def _src(prop: str) -> str:
        return str((resolved.get(prop) or {}).get("source") or "none")

    fields: list[dict] = []
    today = datetime.now().date()

    # ── Section A: Equipment_Configuration (mandatory gate) ───────────────────
    eq_cfg_raw = _val("Equipment_Configuration")
    eq_cfg_src = _src("Equipment_Configuration")
    # Fallback: Property Tab Builder templates often omit Equipment_Configuration
    # as a separate field but include Equipment_Type with the same allowed values.
    # If Equipment_Configuration is blank but Equipment_Type is a valid config
    # string, promote it so downstream rules can resolve correctly.
    if _is_blank(eq_cfg_raw):
        eq_type_val = _val("Equipment_Type")
        if not _is_blank(eq_type_val) and _norm_text(eq_type_val) in _ALLOWED_CONFIGS:
            eq_cfg_raw = eq_type_val
            eq_cfg_src = _src("Equipment_Type") + " (via Equipment_Type fallback)"
    eq_cfg_norm = _norm_text(eq_cfg_raw)
    eq_cfg_valid = eq_cfg_norm in _ALLOWED_CONFIGS

    if _is_blank(eq_cfg_raw):
        fields.append(_missing("Equipment_Configuration"))
        eq_cfg_display = "unknown"
        config_hold = True
    elif not eq_cfg_valid:
        fields.append(_field(
            "Equipment_Configuration", eq_cfg_src, "required",
            eq_cfg_raw, eq_cfg_norm,
            "hold",
            f"Invalid value {eq_cfg_raw!r}. Allowed: Vessel, Jacketed Vessel, "
            f"Heat Exchanger, Jacketed Vessel and Heat Exchanger",
        ))
        eq_cfg_display = eq_cfg_raw
        config_hold = True
    else:
        fields.append(_required_pass("Equipment_Configuration", eq_cfg_src, eq_cfg_raw, eq_cfg_norm))
        eq_cfg_display = eq_cfg_raw
        config_hold = False

    _log(f"[CPVerify] Equipment_Configuration: {eq_cfg_raw!r} valid={eq_cfg_valid}")

    tube_active, jacket_active = _active_columns(eq_cfg_norm) if eq_cfg_valid else (False, False)

    # ── Section C: Always required fields (excluding Equipment_Configuration) ──
    always_required = [
        "HYDRO_TEST_POSITION",
        "Drawing_Number",
        "Tag_No",
        "Equipment_Type",
        "Design_Code",
        "Material_Code",
        "Inspection_By",
        "Revision",
    ]
    for prop in always_required:
        v = _val(prop)
        s = _src(prop)
        if _is_blank(v):
            fields.append(_missing(prop))
        else:
            fields.append(_required_pass(prop, s, v))

    # ── Section B: Conditional IDP/MOT fields ─────────────────────────────────
    # Shell — always required when equipment config is valid
    for prop in ("SHELL_IDP", "SHELL_MOT"):
        v = _val(prop)
        s = _src(prop)
        if config_hold:
            fields.append(_required_hold(prop, s, v, "Equipment_Configuration invalid — cannot determine applicability"))
        elif _is_blank(v):
            fields.append(_missing(prop))
        else:
            norm = _norm_numeric(v) or v
            fields.append(_required_pass(prop, s, v, norm))

    # Tube — required only for Heat Exchanger / Jacketed Vessel and Heat Exchanger
    for prop in ("TUBE_IDP", "TUBE_MOT"):
        v = _val(prop)
        s = _src(prop)
        if config_hold:
            fields.append(_required_hold(prop, s, v, "Equipment_Configuration invalid — cannot determine applicability"))
        elif tube_active:
            if _is_blank(v):
                fields.append(_missing(prop))
            else:
                norm = _norm_numeric(v) or v
                fields.append(_required_pass(prop, s, v, norm))
        else:
            # Not applicable — but if a value is present the drawing is wrong
            if _is_blank(v):
                fields.append(_not_applicable(prop, s, v))
            else:
                fields.append(_not_applicable_contaminated(prop, s, v, eq_cfg_display))

    # Jacket — required only for Jacketed Vessel / Jacketed Vessel and Heat Exchanger
    for prop in ("JACKET_IDP", "JACKET_MOT"):
        v = _val(prop)
        s = _src(prop)
        if config_hold:
            fields.append(_required_hold(prop, s, v, "Equipment_Configuration invalid — cannot determine applicability"))
        elif jacket_active:
            if _is_blank(v):
                fields.append(_missing(prop))
            else:
                norm = _norm_numeric(v) or v
                fields.append(_required_pass(prop, s, v, norm))
        else:
            # Not applicable — but if a value is present the drawing is wrong
            if _is_blank(v):
                fields.append(_not_applicable(prop, s, v))
            else:
                fields.append(_not_applicable_contaminated(prop, s, v, eq_cfg_display))

    # ── Section D: Rule-based fields ──────────────────────────────────────────

    # DrawnBy
    drawn_by = _val("DrawnBy")
    if _is_blank(drawn_by):
        fields.append(_missing("DrawnBy"))
    else:
        fields.append(_required_pass("DrawnBy", _src("DrawnBy"), drawn_by))

    # DrawnDate — valid date, before today
    drawn_date_raw = _val("DrawnDate")
    drawn_date = _parse_date(drawn_date_raw)
    if _is_blank(drawn_date_raw):
        fields.append(_missing("DrawnDate"))
    elif drawn_date is None:
        fields.append(_required_hold("DrawnDate", _src("DrawnDate"), drawn_date_raw,
                                     f"Cannot parse as date: {drawn_date_raw!r}"))
    elif drawn_date >= today:
        fields.append(_required_fail("DrawnDate", _src("DrawnDate"), drawn_date_raw,
                                     f"DrawnDate {drawn_date} must be before today ({today})"))
    else:
        fields.append(_required_pass("DrawnDate", _src("DrawnDate"), drawn_date_raw,
                                     str(drawn_date)))

    # CheckedBy
    checked_by = _val("CheckedBy")
    if _is_blank(checked_by):
        fields.append(_missing("CheckedBy"))
    else:
        fields.append(_required_pass("CheckedBy", _src("CheckedBy"), checked_by))

    # CheckedDate — valid date, >= DrawnDate
    checked_date_raw = _val("CheckedDate")
    checked_date = _parse_date(checked_date_raw)
    if _is_blank(checked_date_raw):
        fields.append(_missing("CheckedDate"))
    elif checked_date is None:
        fields.append(_required_hold("CheckedDate", _src("CheckedDate"), checked_date_raw,
                                     f"Cannot parse as date: {checked_date_raw!r}"))
    elif drawn_date is None:
        fields.append(_required_hold("CheckedDate", _src("CheckedDate"), checked_date_raw,
                                     "Cannot verify ≥ DrawnDate because DrawnDate is invalid"))
    elif checked_date < drawn_date:
        fields.append(_required_fail("CheckedDate", _src("CheckedDate"), checked_date_raw,
                                     f"CheckedDate {checked_date} is before DrawnDate {drawn_date}"))
    else:
        fields.append(_required_pass("CheckedDate", _src("CheckedDate"), checked_date_raw,
                                     str(checked_date)))

    # EngineeringApproval
    eng_approval = _val("EngineeringApproval")
    if _is_blank(eng_approval):
        fields.append(_missing("EngineeringApproval"))
    else:
        fields.append(_required_pass("EngineeringApproval", _src("EngineeringApproval"), eng_approval))

    # EngAppDate — valid date, >= CheckedDate
    eng_app_date_raw = _val("EngAppDate")
    eng_app_date = _parse_date(eng_app_date_raw)
    if _is_blank(eng_app_date_raw):
        fields.append(_missing("EngAppDate"))
    elif eng_app_date is None:
        fields.append(_required_hold("EngAppDate", _src("EngAppDate"), eng_app_date_raw,
                                     f"Cannot parse as date: {eng_app_date_raw!r}"))
    elif checked_date is None:
        fields.append(_required_hold("EngAppDate", _src("EngAppDate"), eng_app_date_raw,
                                     "Cannot verify ≥ CheckedDate because CheckedDate is invalid"))
    elif eng_app_date < checked_date:
        fields.append(_required_fail("EngAppDate", _src("EngAppDate"), eng_app_date_raw,
                                     f"EngAppDate {eng_app_date} is before CheckedDate {checked_date}"))
    else:
        fields.append(_required_pass("EngAppDate", _src("EngAppDate"), eng_app_date_raw,
                                     str(eng_app_date)))

    # ── Overall status ────────────────────────────────────────────────────────
    # not_applicable fields do not count toward hold/fail/pass
    scorable = [f for f in fields if f["result"] != "not_applicable"]
    any_hold = any(f["result"] == "hold" for f in scorable)
    any_fail = any(f["result"] == "fail" for f in scorable)

    if any_hold:
        overall = "hold"
    elif any_fail:
        overall = "fail"
    else:
        overall = "pass"

    counts = {r: sum(1 for f in fields if f["result"] == r)
              for r in ("pass", "fail", "hold", "not_applicable")}
    _log(
        f"[CPVerify] Complete: overall={overall} equipmentConfig={eq_cfg_display!r} "
        f"pass={counts['pass']} fail={counts['fail']} hold={counts['hold']} "
        f"not_applicable={counts['not_applicable']} total={len(fields)}"
    )

    return {
        "customPropertyVerification": {
            "status":          overall,
            "equipmentConfig": eq_cfg_display,
            "fields":          fields,
        }
    }
