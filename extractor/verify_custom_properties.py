"""
verify_custom_properties.py — Layer 1: Custom Property Verification

Validates custom properties extracted from a SolidWorks drawing.
No DDS table comparison.  Rules are applied purely to the property values themselves.

Sections:
  A  — Equipment_Configuration gate (1 field)
  C  — Always-required header fields (10 fields)
  B  — Conditional IDP / MOT numeric checks (6 fields, config-dependent)
  D  — Engineer-filled fields with date / sequence rules (7 fields)
  E  — Mechanical column properties: SHELL / TUBE / JACKET × 22 fields each
       (up to 66 fields; column applicability mirrors structuring agent logic)
  F  — General Data properties: orientation, service life, wind/seismic codes,
       weights, location, quantity (12 fields; sourced from dds.general_data)

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


# ─── Section E: Mechanical column constants and helpers ───────────────────────

_RT_OPTIONS = frozenset({
    'full radiography (100% rt)',
    'spot radiography (10% rt)',
    'spot radiography (5% rt)',
    'no radiography',
})
_RT_DISPLAY = (
    'FULL RADIOGRAPHY (100% RT)',
    'SPOT RADIOGRAPHY (10% RT)',
    'SPOT RADIOGRAPHY (5% RT)',
    'NO RADIOGRAPHY',
)
_JE_BY_RT = {
    'full radiography (100% rt)':  '1 / 1 / 1',
    'spot radiography (10% rt)':   '1 / 1 / 1',
    'spot radiography (5% rt)':    '0.85 / 0.85 / 0.85',
    'no radiography':              '0.70 / 0.70 / 0.70',
}
_PWHT_OPTIONS = frozenset({'not required', 'required'})
_HEAD_OPTIONS = frozenset({
    'torispherical (10%)', 'ellipsoidal (2:1)', 'hemispherical',
    'flat head', 'conical head', 'dished end (f&d)', 'n.a.',
})
_INS_OPTIONS = frozenset({'yes', 'no', 'n.a.'})

_MECH_SHORTS = (
    'EDP', 'WP', 'HTP', 'MDMT',
    'HT_TEMP', 'OP_TEMP', 'DES_TEMP',
    'STATE', 'VOL', 'FLUID', 'HZ', 'SG',
    'ICA', 'ECA',
    'RT', 'JE', 'TG', 'FTC', 'PWHT',
    'HEAD', 'INS', 'INS_SPEC',
)


def _parse_min_max(v: str):
    """Parse 'min / max' string. Returns (float, float) or None."""
    if '/' not in v:
        return None
    parts = v.split('/')
    if len(parts) < 2:
        return None
    try:
        return (float(parts[0].strip()), float(parts[-1].strip()))
    except (ValueError, AttributeError):
        return None


def _parse_three_slash(v: str):
    """Parse 'a / b / c' string. Returns [float, float, float] or None."""
    parts = v.split('/')
    if len(parts) != 3:
        return None
    try:
        return [float(p.strip()) for p in parts]
    except (ValueError, AttributeError):
        return None


def _opt_pass(prop: str, source: str, value: str,
              norm: str = "", reason: str = "") -> dict:
    return _field(prop, source, "optional", value, norm or value, "pass", reason)


def _opt_hold(prop: str, source: str, value: str, reason: str) -> dict:
    return _field(prop, source, "optional", value, "", "hold", reason)


def _opt_blank(prop: str) -> dict:
    """Optional field absent — non-blocking pass."""
    return _field(prop, "none", "optional", "", "", "pass", "Optional — not present")


def _verify_mech_column(prefix: str, resolved: dict,
                        column_active: bool, eq_cfg_display: str) -> list:
    """
    Verify all 22 Section E mechanical fields for one column prefix
    (SHELL / TUBE / JACKET).

    column_active=False  → all fields marked not_applicable (or contaminated
                           fail if a value is unexpectedly present).
    column_active=True   → full rule set applied.

    Validation tiers:
      WARNING (hold)  — numeric fields with bad format, enum fields with
                        unknown value, cross-field logic violations,
                        conditional INS_SPEC rule.
      INFO    (pass)  — STATE, VOL, FLUID, HZ, SG, TG, FTC always pass;
                        format notes added to reason string only.
    """
    fields = []

    def _v(short: str) -> str:
        return str((resolved.get(f"{prefix}_{short}") or {}).get("value") or "").strip()

    def _s(short: str) -> str:
        return str((resolved.get(f"{prefix}_{short}") or {}).get("source") or "none")

    def _p(short: str) -> str:
        return f"{prefix}_{short}"

    # ── Column not active: mark every field not_applicable / contaminated ──────
    if not column_active:
        for short in _MECH_SHORTS:
            v = _v(short)
            s = _s(short)
            if _is_blank(v):
                fields.append(_not_applicable(_p(short), s, v))
            else:
                fields.append(_not_applicable_contaminated(
                    _p(short), s, v,
                    f"{prefix} column not applicable for "
                    f"Equipment_Configuration={eq_cfg_display!r}",
                ))
        return fields

    # ── Read IDP for cross-checks (already FAIL-checked in Section B) ─────────
    idp_num: float | None = None
    try:
        idp_num = float(_v("IDP")) if _v("IDP") else None
    except ValueError:
        pass

    # ── EDP — numeric (Barg) or "N.A." ───────────────────────────────────────
    v, s = _v("EDP"), _s("EDP")
    if _is_blank(v):
        fields.append(_opt_blank(_p("EDP")))
    elif v.upper() == "N.A.":
        fields.append(_opt_pass(_p("EDP"), s, v, "N.A."))
    else:
        try:
            float(v)
            fields.append(_opt_pass(_p("EDP"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_hold(_p("EDP"), s, v,
                f"Expected numeric Barg or 'N.A.' — got {v!r}"))

    # ── WP — numeric (Barg); cross-check WP < IDP ────────────────────────────
    v, s = _v("WP"), _s("WP")
    if _is_blank(v):
        fields.append(_opt_blank(_p("WP")))
    else:
        try:
            wp_num = float(v)
            if idp_num is not None and wp_num >= idp_num:
                fields.append(_opt_hold(_p("WP"), s, v,
                    f"WP ({v}) >= IDP ({_v('IDP')}) — "
                    "working pressure should be below design pressure"))
            else:
                fields.append(_opt_pass(_p("WP"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_hold(_p("WP"), s, v,
                f"Expected numeric Barg — got {v!r}"))

    # ── HTP — numeric (Barg) or "N.A." ───────────────────────────────────────
    v, s = _v("HTP"), _s("HTP")
    if _is_blank(v):
        fields.append(_opt_blank(_p("HTP")))
    elif v.upper() == "N.A.":
        fields.append(_opt_pass(_p("HTP"), s, v, "N.A."))
    else:
        try:
            float(v)
            fields.append(_opt_pass(_p("HTP"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_hold(_p("HTP"), s, v,
                f"Expected numeric Barg or 'N.A.' — got {v!r}"))

    # ── MDMT — numeric °C (may be negative) ──────────────────────────────────
    v, s = _v("MDMT"), _s("MDMT")
    if _is_blank(v):
        fields.append(_opt_blank(_p("MDMT")))
    else:
        try:
            float(v)
            fields.append(_opt_pass(_p("MDMT"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_hold(_p("MDMT"), s, v,
                f"Expected numeric DEG. C — got {v!r}"))

    # ── HT_TEMP — "min / max" format ─────────────────────────────────────────
    v, s = _v("HT_TEMP"), _s("HT_TEMP")
    if _is_blank(v):
        fields.append(_opt_blank(_p("HT_TEMP")))
    else:
        parsed = _parse_min_max(v)
        if parsed is None:
            fields.append(_opt_hold(_p("HT_TEMP"), s, v,
                f"Expected 'min / max' format — got {v!r}"))
        elif parsed[0] > parsed[1]:
            fields.append(_opt_hold(_p("HT_TEMP"), s, v,
                f"HT_TEMP min ({parsed[0]}) > max ({parsed[1]})"))
        else:
            fields.append(_opt_pass(_p("HT_TEMP"), s, v))

    # ── OP_TEMP — "min / max" format ─────────────────────────────────────────
    v, s = _v("OP_TEMP"), _s("OP_TEMP")
    op_max: float | None = None
    if _is_blank(v):
        fields.append(_opt_blank(_p("OP_TEMP")))
    else:
        parsed = _parse_min_max(v)
        if parsed is None:
            fields.append(_opt_hold(_p("OP_TEMP"), s, v,
                f"Expected 'min / max' format — got {v!r}"))
        elif parsed[0] >= parsed[1]:
            fields.append(_opt_hold(_p("OP_TEMP"), s, v,
                f"OP_TEMP min ({parsed[0]}) >= max ({parsed[1]})"))
        else:
            op_max = parsed[1]
            fields.append(_opt_pass(_p("OP_TEMP"), s, v))

    # ── DES_TEMP — "min / max"; cross-check DES max >= OP max ────────────────
    v, s = _v("DES_TEMP"), _s("DES_TEMP")
    if _is_blank(v):
        fields.append(_opt_blank(_p("DES_TEMP")))
    else:
        parsed = _parse_min_max(v)
        if parsed is None:
            fields.append(_opt_hold(_p("DES_TEMP"), s, v,
                f"Expected 'min / max' format — got {v!r}"))
        elif parsed[0] > parsed[1]:
            fields.append(_opt_hold(_p("DES_TEMP"), s, v,
                f"DES_TEMP min ({parsed[0]}) > max ({parsed[1]})"))
        elif op_max is not None and parsed[1] < op_max:
            fields.append(_opt_hold(_p("DES_TEMP"), s, v,
                f"DES_TEMP max ({parsed[1]}) < OP_TEMP max ({op_max}) — "
                "design temperature must be >= operating temperature"))
        else:
            fields.append(_opt_pass(_p("DES_TEMP"), s, v))

    # ── STATE — INFO (enum soft, not blocking) ────────────────────────────────
    v, s = _v("STATE"), _s("STATE")
    fields.append(_opt_blank(_p("STATE")) if _is_blank(v)
                  else _opt_pass(_p("STATE"), s, v))

    # ── VOL — INFO (positive numeric, not blocking) ───────────────────────────
    v, s = _v("VOL"), _s("VOL")
    if _is_blank(v):
        fields.append(_opt_blank(_p("VOL")))
    else:
        try:
            float(v)
            fields.append(_opt_pass(_p("VOL"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_pass(_p("VOL"), s, v, "",
                f"Non-numeric volume {v!r} (INFO — not blocking)"))

    # ── FLUID — INFO ──────────────────────────────────────────────────────────
    v, s = _v("FLUID"), _s("FLUID")
    fields.append(_opt_blank(_p("FLUID")) if _is_blank(v)
                  else _opt_pass(_p("FLUID"), s, v))

    # ── HZ — INFO (auto-derived, no format rule) ──────────────────────────────
    v, s = _v("HZ"), _s("HZ")
    fields.append(_opt_blank(_p("HZ")) if _is_blank(v)
                  else _opt_pass(_p("HZ"), s, v))

    # ── SG — INFO ("liquid / gas" composite, not blocking) ───────────────────
    v, s = _v("SG"), _s("SG")
    fields.append(_opt_blank(_p("SG")) if _is_blank(v)
                  else _opt_pass(_p("SG"), s, v))

    # ── ICA — numeric mm; WARNING if non-numeric ──────────────────────────────
    v, s = _v("ICA"), _s("ICA")
    if _is_blank(v):
        fields.append(_opt_blank(_p("ICA")))
    else:
        try:
            float(v)
            fields.append(_opt_pass(_p("ICA"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_hold(_p("ICA"), s, v,
                f"Expected numeric mm — got {v!r}"))

    # ── ECA — numeric mm; WARNING if non-numeric ──────────────────────────────
    v, s = _v("ECA"), _s("ECA")
    if _is_blank(v):
        fields.append(_opt_blank(_p("ECA")))
    else:
        try:
            float(v)
            fields.append(_opt_pass(_p("ECA"), s, v, _norm_numeric(v) or v))
        except ValueError:
            fields.append(_opt_hold(_p("ECA"), s, v,
                f"Expected numeric mm — got {v!r}"))

    # ── RT — enum; WARNING if not in allowed list ─────────────────────────────
    v, s = _v("RT"), _s("RT")
    rt_norm = _norm_text(v)
    if _is_blank(v):
        fields.append(_opt_blank(_p("RT")))
    elif rt_norm not in _RT_OPTIONS:
        fields.append(_opt_hold(_p("RT"), s, v,
            f"Unrecognised radiography value {v!r}. "
            f"Allowed: {' | '.join(_RT_DISPLAY)}"))
    else:
        fields.append(_opt_pass(_p("RT"), s, v))

    # ── JE — "x / y / z" format; cross-check with RT ─────────────────────────
    v, s = _v("JE"), _s("JE")
    if _is_blank(v):
        fields.append(_opt_blank(_p("JE")))
    else:
        parts = _parse_three_slash(v)
        if parts is None:
            fields.append(_opt_hold(_p("JE"), s, v,
                f"Expected 'x / y / z' (3 numeric parts) — got {v!r}"))
        else:
            expected_je = _JE_BY_RT.get(rt_norm, "")
            note = ""
            if expected_je and _norm_text(v) != _norm_text(expected_je):
                note = (f"JE {v!r} does not match expected {expected_je!r} "
                        f"for RT {_v('RT')!r} (INFO — not blocking)")
            fields.append(_opt_pass(_p("JE"), s, v, "", note))

    # ── TG — INFO (Testing Group, code-dependent) ─────────────────────────────
    v, s = _v("TG"), _s("TG")
    fields.append(_opt_blank(_p("TG")) if _is_blank(v)
                  else _opt_pass(_p("TG"), s, v))

    # ── FTC — INFO (Fabrication Tolerance Class, code-dependent) ─────────────
    v, s = _v("FTC"), _s("FTC")
    fields.append(_opt_blank(_p("FTC")) if _is_blank(v)
                  else _opt_pass(_p("FTC"), s, v))

    # ── PWHT — enum; WARNING if not in allowed list ───────────────────────────
    v, s = _v("PWHT"), _s("PWHT")
    if _is_blank(v):
        fields.append(_opt_blank(_p("PWHT")))
    elif _norm_text(v) not in _PWHT_OPTIONS:
        fields.append(_opt_hold(_p("PWHT"), s, v,
            f"Unrecognised PWHT value {v!r}. Allowed: NOT REQUIRED | REQUIRED"))
    else:
        fields.append(_opt_pass(_p("PWHT"), s, v))

    # ── HEAD — enum; WARNING if not in allowed list ───────────────────────────
    v, s = _v("HEAD"), _s("HEAD")
    if _is_blank(v):
        fields.append(_opt_blank(_p("HEAD")))
    elif _norm_text(v) not in _HEAD_OPTIONS:
        fields.append(_opt_hold(_p("HEAD"), s, v,
            f"Unrecognised type of heads {v!r}. "
            "Allowed: TORISPHERICAL (10%) | ELLIPSOIDAL (2:1) | HEMISPHERICAL "
            "| FLAT HEAD | CONICAL HEAD | DISHED END (F&D) | N.A."))
    else:
        fields.append(_opt_pass(_p("HEAD"), s, v))

    # ── INS — enum; WARNING if not in allowed list ────────────────────────────
    ins_raw = _v("INS")
    v, s = ins_raw, _s("INS")
    if _is_blank(v):
        fields.append(_opt_blank(_p("INS")))
    elif _norm_text(v) not in _INS_OPTIONS:
        fields.append(_opt_hold(_p("INS"), s, v,
            f"Unrecognised insulation value {v!r}. Allowed: YES | NO | N.A."))
    else:
        fields.append(_opt_pass(_p("INS"), s, v))

    # ── INS_SPEC — conditional: WARNING if INS=YES and blank ─────────────────
    v, s = _v("INS_SPEC"), _s("INS_SPEC")
    if _is_blank(v):
        if _norm_text(ins_raw) == "yes":
            fields.append(_opt_hold(_p("INS_SPEC"), s, v,
                f"INS={ins_raw!r} but INS_SPEC is blank — "
                "insulation type/thickness/density specification is required"))
        else:
            fields.append(_opt_blank(_p("INS_SPEC")))
    else:
        fields.append(_opt_pass(_p("INS_SPEC"), s, v))

    return fields


# ─── Section F: General Data constants ───────────────────────────────────────

_HYDRO_TEST_POSITION_VALUES = frozenset(["vertical", "horizontal"])

_GENERAL_ORIENT_VALUES = frozenset(["vertical", "horizontal"])

_GENERAL_SERVICE_LIFE_VALUES = frozenset([
    "5 years", "10 years", "15 years", "20 years", "25 years", "30 years",
    "5", "10", "15", "20", "25", "30",
])


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
        "Drawing_Number",
        "Tag_No",
        "Serial_No",
        "Description",
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

    # ── HYDRO_TEST_POSITION — required, enum: VERTICAL / HORIZONTAL ───────────
    _htp_raw = _val("HYDRO_TEST_POSITION")
    _htp_src = _src("HYDRO_TEST_POSITION")
    if _is_blank(_htp_raw):
        fields.append(_missing("HYDRO_TEST_POSITION"))
    elif _norm_text(_htp_raw) not in _HYDRO_TEST_POSITION_VALUES:
        fields.append(_required_fail(
            "HYDRO_TEST_POSITION", _htp_src, _htp_raw,
            f"Invalid value {_htp_raw!r}. Allowed: VERTICAL, HORIZONTAL",
        ))
    else:
        fields.append(_required_pass(
            "HYDRO_TEST_POSITION", _htp_src, _htp_raw,
            _norm_text(_htp_raw).upper(),
        ))

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

    # ── Section E: Mechanical column properties ───────────────────────────────
    # SHELL always active; TUBE / JACKET follow Equipment_Configuration.
    # _verify_mech_column handles column_active=False by marking each field
    # not_applicable (blank) or contaminated-fail (unexpected value present).
    _log("[CPVerify] Section E — mechanical column verification starting")
    fields.extend(_verify_mech_column("SHELL",  resolved, True,         eq_cfg_display))
    fields.extend(_verify_mech_column("TUBE",   resolved, tube_active,  eq_cfg_display))
    fields.extend(_verify_mech_column("JACKET", resolved, jacket_active, eq_cfg_display))
    _log(
        f"[CPVerify] Section E complete — SHELL=active "
        f"TUBE={'active' if tube_active else 'not_applicable'} "
        f"JACKET={'active' if jacket_active else 'not_applicable'}"
    )

    # ── Section F: General Data properties ───────────────────────────────────
    _log("[CPVerify] Section F — general data property verification starting")

    # ── GENERAL_ORIENT — required enum: VERTICAL / HORIZONTAL ────────────────
    _go_raw = _val("GENERAL_ORIENT")
    _go_src = _src("GENERAL_ORIENT")
    if _is_blank(_go_raw):
        fields.append(_missing("GENERAL_ORIENT"))
    elif _norm_text(_go_raw) not in _GENERAL_ORIENT_VALUES:
        fields.append(_required_fail(
            "GENERAL_ORIENT", _go_src, _go_raw,
            f"Invalid value {_go_raw!r}. Allowed: VERTICAL, HORIZONTAL",
        ))
    else:
        fields.append(_required_pass(
            "GENERAL_ORIENT", _go_src, _go_raw,
            _norm_text(_go_raw).upper(),
        ))

    # ── GENERAL_SERVICE_LIFE — optional, enum dropdown ────────────────────────
    _gsl_raw = _val("GENERAL_SERVICE_LIFE")
    _gsl_src = _src("GENERAL_SERVICE_LIFE")
    if _is_blank(_gsl_raw):
        fields.append(_opt_blank("GENERAL_SERVICE_LIFE"))
    elif _norm_text(_gsl_raw) not in _GENERAL_SERVICE_LIFE_VALUES:
        fields.append(_opt_hold(
            "GENERAL_SERVICE_LIFE", _gsl_src, _gsl_raw,
            f"Unrecognised service life {_gsl_raw!r}. "
            "Expected one of: 5 years, 10 years, 15 years, 20 years, 25 years, 30 years",
        ))
    else:
        fields.append(_opt_pass("GENERAL_SERVICE_LIFE", _gsl_src, _gsl_raw))

    # ── GENERAL_WIND_CODE — optional, INFO free text ──────────────────────────
    for _gf_prop in ("GENERAL_WIND_CODE", "GENERAL_WIND_VEL", "GENERAL_SEISMIC_CODE",
                     "GENERAL_LOCATION"):
        _gf_raw = _val(_gf_prop)
        _gf_src = _src(_gf_prop)
        if _is_blank(_gf_raw):
            fields.append(_opt_blank(_gf_prop))
        else:
            fields.append(_opt_pass(_gf_prop, _gf_src, _gf_raw))

    # ── GENERAL_SEISMIC_Z / _H / _V — optional, WARNING if non-numeric ───────
    for _gs_prop in ("GENERAL_SEISMIC_Z", "GENERAL_SEISMIC_H", "GENERAL_SEISMIC_V"):
        _gs_raw = _val(_gs_prop)
        _gs_src = _src(_gs_prop)
        if _is_blank(_gs_raw):
            fields.append(_opt_blank(_gs_prop))
        elif not _norm_numeric(_gs_raw):
            fields.append(_opt_hold(
                _gs_prop, _gs_src, _gs_raw,
                f"{_gs_prop} must be numeric — {_gs_raw!r} is not a valid number",
            ))
        else:
            fields.append(_opt_pass(_gs_prop, _gs_src, _gs_raw,
                                    norm=_norm_numeric(_gs_raw)))

    # ── GENERAL_WEIGHT — optional, WARNING if W1/W2/W3 format invalid
    #    or W1 > W2 (empty heavier than operating) or W2 > W3 (operating > test)
    _gw_raw = _val("GENERAL_WEIGHT")
    _gw_src = _src("GENERAL_WEIGHT")
    if _is_blank(_gw_raw):
        fields.append(_opt_blank("GENERAL_WEIGHT"))
    else:
        _gw_parts = _parse_three_slash(_gw_raw)
        if _gw_parts is None:
            fields.append(_opt_hold(
                "GENERAL_WEIGHT", _gw_src, _gw_raw,
                f"Expected 'W1 / W2 / W3' (empty / operating / test) format — "
                f"cannot parse {_gw_raw!r}",
            ))
        else:
            _gw1, _gw2, _gw3 = _gw_parts
            if _gw1 > _gw2:
                fields.append(_opt_hold(
                    "GENERAL_WEIGHT", _gw_src, _gw_raw,
                    f"Empty weight ({_gw1}) must not exceed operating weight ({_gw2})",
                ))
            elif _gw2 > _gw3:
                fields.append(_opt_hold(
                    "GENERAL_WEIGHT", _gw_src, _gw_raw,
                    f"Operating weight ({_gw2}) must not exceed test/hydro weight ({_gw3})",
                ))
            else:
                fields.append(_opt_pass("GENERAL_WEIGHT", _gw_src, _gw_raw))

    # ── GENERAL_QTY — optional, WARNING if non-numeric or < 1 ───────────────
    _gq_raw = _val("GENERAL_QTY")
    _gq_src = _src("GENERAL_QTY")
    if _is_blank(_gq_raw):
        fields.append(_opt_blank("GENERAL_QTY"))
    elif not _norm_numeric(_gq_raw):
        fields.append(_opt_hold(
            "GENERAL_QTY", _gq_src, _gq_raw,
            f"GENERAL_QTY must be numeric — {_gq_raw!r} is not a valid number",
        ))
    else:
        try:
            _gq_num = float(_gq_raw)
        except ValueError:
            _gq_num = 0.0
        if _gq_num < 1:
            fields.append(_opt_hold(
                "GENERAL_QTY", _gq_src, _gq_raw,
                f"GENERAL_QTY must be ≥ 1 — got {_gq_raw!r}",
            ))
        else:
            fields.append(_opt_pass("GENERAL_QTY", _gq_src, _gq_raw,
                                    norm=_norm_numeric(_gq_raw)))

    _log("[CPVerify] Section F complete — 11 GENERAL_* fields evaluated")

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
