"""
verify_custom_properties.py — Layer 1: Drawing Custom Property Verification

Compares SolidWorks drawing custom properties against:
  - Values extracted from the drawing's own DDS table  (check_type="dds_compare")
  - Approval workflow rules                            (check_type="rule")

Custom properties verified:
  DDS compare : HYDRO_TEST_POSITION, SHELL_IDP, SHELL_MOT, TUBE_IDP, TUBE_MOT,
                JACKET_IDP, JACKET_MOT, Drawing_Number, Tag_No, Equipment_Type,
                Design_Code, Material_Code, Inspection_By, Revision
  Rule checks : DrawnBy, DrawnDate, CheckedBy, CheckedDate,
                EngineeringApproval, EngAppDate

Normalization:
  - text   : trim, collapse whitespace, compare case-insensitive
  - numeric: parse float; 3.5 == 3.500
  - dates  : parse to date objects before comparison

Status logic (per field and overall):
  pass  — check succeeded
  fail  — comparison mismatch OR rule violation
  hold  — missing/unparseable value, or dependency (e.g. DrawnDate) is invalid
           hold takes precedence over fail in overall status

Output shape:
  {
    "customPropertyVerification": {
      "status": "pass" | "fail" | "hold",
      "fields": [
        {
          "property":     str,
          "check_type":   "dds_compare" | "rule",
          "status":       "pass" | "fail" | "hold",
          "custom_value": str | None,
          "dds_value":    str | None,    # dds_compare only; None for rule checks
          "reason":       str | None     # populated when status != "pass"
        },
        ...
      ]
    }
  }
"""

from __future__ import annotations
import re
from datetime import datetime, date


# ─── Normalization helpers ────────────────────────────────────────────────────

def _norm_text(val: str | None) -> str:
    if not val:
        return ""
    return re.sub(r"\s+", " ", str(val).strip()).lower()


def _is_blank(val: str | None) -> bool:
    if val is None:
        return True
    s = str(val).strip()
    return s == "" or s == "-"


def _try_parse_numeric(val: str | None) -> float | None:
    if not val:
        return None
    try:
        return float(str(val).strip().replace(",", ""))
    except (ValueError, TypeError):
        return None


_DATE_FORMATS = [
    "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y",
    "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y",
    "%d/%m/%y", "%d-%m-%y", "%m/%d/%Y", "%m-%d-%Y",
    "%d-%b-%Y", "%d-%B-%Y",
]


def _parse_date(val: str | None) -> date | None:
    if not val:
        return None
    s = str(val).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _compare_values(custom_val: str | None, dds_val: str | None) -> tuple[bool, str | None]:
    """
    Compare a custom property value against a DDS table value.
    Returns (match: bool, reason: str | None).
    Numeric comparison is attempted first; falls back to normalised-text comparison.
    """
    cn = _try_parse_numeric(custom_val)
    dn = _try_parse_numeric(dds_val)
    if cn is not None and dn is not None:
        if abs(cn - dn) < 1e-9:
            return True, None
        return False, f"Numeric mismatch: custom={custom_val!r} dds={dds_val!r}"

    ct = _norm_text(custom_val)
    dt = _norm_text(dds_val)
    if ct == dt:
        return True, None
    return False, f"Value mismatch: custom={custom_val!r} dds={dds_val!r}"


# ─── DDS block helpers ────────────────────────────────────────────────────────

_NA_VALUES = frozenset({"", "n.a.", "n/a", "-", "na"})


def _norm_field_key(key: str) -> str:
    return re.sub(r"[\s_]+", " ", key.strip().lower())


def _field_value_from_block(block: dict, field_key: str) -> str | None:
    """
    Lookup VALUE from a FIELD/VALUE block by FIELD name (case/space/underscore
    insensitive). Returns None when not found or block invalid.
    """
    rows = block.get("rows") or []
    norm_key = _norm_field_key(field_key)
    for row in rows:
        if len(row) < 2:
            continue
        row_field = _norm_field_key(str(row[0]))
        if row_field == norm_key:
            v = str(row[1]).strip() if row[1] else None
            return v if v and v.lower() not in _NA_VALUES else None
    return None


def _field_value_from_block_multi(block: dict, *keys: str) -> str | None:
    """Try multiple field key aliases, return first match."""
    for key in keys:
        v = _field_value_from_block(block, key)
        if v is not None:
            return v
    return None


def _mech_col_value(mech_block: dict, param_fragment: str, col_index: int) -> str | None:
    """
    Extract a mechanical data value by PARAMETER name fragment and column index.
    Rows are [GROUP, PARAMETER, SHELL, TUBE, JACKET] (indices 0..4).
    param_fragment is matched as a substring of the PARAMETER cell (case-insensitive).
    """
    rows = mech_block.get("rows") or []
    norm_frag = _norm_text(param_fragment)
    for row in rows:
        if len(row) <= col_index:
            continue
        param_text = _norm_text(row[1]) if len(row) > 1 else ""
        if norm_frag in param_text:
            v = str(row[col_index]).strip() if row[col_index] else ""
            if v and v.lower() not in _NA_VALUES:
                return v
    return None


def _col_is_active(mech_block: dict, col_index: int) -> bool:
    """Return True if at least one data row in col_index has a real value."""
    for row in mech_block.get("rows") or []:
        if len(row) > col_index:
            v = str(row[col_index]).strip()
            if v and v.lower() not in _NA_VALUES:
                return True
    return False


# ─── Equipment configuration helpers ─────────────────────────────────────────

# Canonical config names (lowercased for matching)
_CONFIG_VESSEL          = "vessel"
_CONFIG_JACKETED        = "jacketed vessel"
_CONFIG_HE              = "heat exchanger"
_CONFIG_JACKETED_HE     = "jacketed vessel and heat exchanger"

# Which IDP/MOT columns (tube, jacket) each config activates.
# Shell is ALWAYS active when any config is found.
_CONFIG_TUBE_ACTIVE: dict[str, bool] = {
    _CONFIG_VESSEL:      False,
    _CONFIG_JACKETED:    False,
    _CONFIG_HE:          True,
    _CONFIG_JACKETED_HE: True,
}
_CONFIG_JACKET_ACTIVE: dict[str, bool] = {
    _CONFIG_VESSEL:      False,
    _CONFIG_JACKETED:    True,
    _CONFIG_HE:          False,
    _CONFIG_JACKETED_HE: True,
}


def _resolve_equipment_config(meta_block: dict, meta_valid: bool,
                               custom_properties: dict) -> str | None:
    """
    Determine the equipment configuration string.
    Priority:
      1. DDS METADATA block — "Equipment Configuration" / "Equipment Config" fields
      2. Custom property "Equipment_Config" / "EquipmentConfig" / "Equipment Configuration"
    Returns the raw value (not normalised) or None if not found.
    """
    if meta_valid:
        v = _field_value_from_block_multi(
            meta_block,
            "Equipment Configuration", "Equipment Config",
            "Equipment_Configuration", "Equipment_Config",
            "EquipmentConfig", "EquipmentConfiguration",
        )
        if v:
            return v

    for key in ("Equipment_Config", "EquipmentConfig",
                "Equipment Configuration", "Equipment_Configuration"):
        raw = custom_properties.get(key)
        if raw:
            s = str(raw).strip()
            if s:
                return s
    return None


def _config_active_cols(config_raw: str | None) -> tuple[bool, bool, str]:
    """
    Map equipment configuration to (tube_active, jacket_active, canonical_name).
    Shell is always active when the config is known.
    Falls back to (False, False, 'unknown') when config is unrecognised.
    """
    if config_raw is None:
        return False, False, "unknown"
    norm = re.sub(r"\s+", " ", config_raw.strip().lower())
    # Use the most specific match first (jacketed vessel and heat exchanger)
    for key in (_CONFIG_JACKETED_HE, _CONFIG_JACKETED, _CONFIG_HE, _CONFIG_VESSEL):
        if norm == key or norm.startswith(key):
            return _CONFIG_TUBE_ACTIVE[key], _CONFIG_JACKET_ACTIVE[key], config_raw.strip()
    return False, False, config_raw.strip()


# ─── Field result builders ────────────────────────────────────────────────────

def _dds_result(
    property_name: str,
    custom_val: str | None,
    dds_val: str | None,
    *,
    active: bool = True,
    dds_available: bool = True,
    inactive_reason: str = "Not applicable for this equipment configuration",
) -> dict:
    """Build a dds_compare field result dict."""
    if not active:
        return {
            "property":     property_name,
            "check_type":   "dds_compare",
            "status":       "hold",
            "custom_value": custom_val,
            "dds_value":    None,
            "reason":       inactive_reason,
        }
    if not dds_available:
        return {
            "property":     property_name,
            "check_type":   "dds_compare",
            "status":       "hold",
            "custom_value": custom_val,
            "dds_value":    None,
            "reason":       "DDS block not found or invalid — cannot compare",
        }
    if custom_val is None:
        return {
            "property":     property_name,
            "check_type":   "dds_compare",
            "status":       "hold",
            "custom_value": None,
            "dds_value":    dds_val,
            "reason":       "Custom property missing or blank",
        }
    if dds_val is None:
        return {
            "property":     property_name,
            "check_type":   "dds_compare",
            "status":       "hold",
            "custom_value": custom_val,
            "dds_value":    None,
            "reason":       "DDS table value not found for this field",
        }
    match, reason = _compare_values(custom_val, dds_val)
    return {
        "property":     property_name,
        "check_type":   "dds_compare",
        "status":       "pass" if match else "fail",
        "custom_value": custom_val,
        "dds_value":    dds_val,
        "reason":       reason,
    }


def _rule_result(
    property_name: str,
    status: str,
    custom_val: str | None,
    reason: str | None,
) -> dict:
    """Build a rule check field result dict."""
    return {
        "property":     property_name,
        "check_type":   "rule",
        "status":       status,
        "custom_value": custom_val,
        "dds_value":    None,
        "reason":       reason,
    }


# ─── Main entry point ─────────────────────────────────────────────────────────

def verify_custom_properties(
    custom_properties: dict,
    design_data_table: dict,
    logger=None,
) -> dict:
    """
    Verify drawing custom properties against DDS values and approval rules.

    Args:
        custom_properties : dict from result["properties"]["custom_properties"]
        design_data_table : dict from result["design_data_table"]
        logger            : optional logger object; falls back to no-op

    Returns:
        {"customPropertyVerification": {"status": str, "fields": list}}
    """

    def _log(msg: str) -> None:
        if logger:
            try:
                logger.info(msg)
            except Exception:
                pass

    _log("[CPVerify] Starting custom property verification — Layer 1")

    cp = custom_properties or {}
    dds_blocks = (design_data_table or {}).get("dds_blocks") or {}

    mech_block = dds_blocks.get("mechanical_design_data") or {}
    gen_block  = dds_blocks.get("general_data") or {}
    meta_block = dds_blocks.get("metadata") or {}

    mech_valid = mech_block.get("status") == "valid"
    gen_valid  = gen_block.get("status") == "valid"
    meta_valid = meta_block.get("status") == "valid"

    _log(
        f"[CPVerify] DDS blocks: "
        f"mech={mech_block.get('status','missing')} "
        f"gen={gen_block.get('status','missing')} "
        f"meta={meta_block.get('status','missing')}"
    )

    # ── Equipment configuration — drives which IDP/MOT columns are checked ────
    equip_config_raw = _resolve_equipment_config(meta_block, meta_valid, cp)

    if equip_config_raw is not None:
        tube_active, jacket_active, equip_config_name = _config_active_cols(equip_config_raw)
        shell_active = mech_valid  # shell always required when config is known
        config_source = "dds_metadata" if (meta_valid and _field_value_from_block_multi(
            meta_block,
            "Equipment Configuration", "Equipment Config",
            "Equipment_Configuration", "Equipment_Config",
            "EquipmentConfig", "EquipmentConfiguration",
        )) else "custom_property"
        # Restrict to mech_valid for mechanical fields
        tube_active   = tube_active   and mech_valid
        jacket_active = jacket_active and mech_valid
    else:
        # Config not found — fall back to data-presence detection with a warning
        equip_config_name = "unknown"
        config_source     = "fallback_data_detection"
        shell_active  = mech_valid and _col_is_active(mech_block, 2)
        tube_active   = mech_valid and _col_is_active(mech_block, 3)
        jacket_active = mech_valid and _col_is_active(mech_block, 4)
        _log("[CPVerify] Equipment configuration not found — falling back to data-presence detection")

    _log(
        f"[CPVerify] Equipment config: name={equip_config_name!r} source={config_source} "
        f"shell={shell_active} tube={tube_active} jacket={jacket_active}"
    )

    def _inactive_reason(col: str) -> str:
        return (
            f"Not applicable for equipment configuration: {equip_config_name!r} "
            f"(column {col} not required)"
        )

    fields: list[dict] = []
    today = datetime.now().date()

    def _prop(name: str) -> str | None:
        v = cp.get(name)
        if v is None:
            return None
        s = str(v).strip()
        return s if s else None

    # ── DDS compare fields ────────────────────────────────────────────────────

    # HYDRO_TEST_POSITION — from GENERAL DATA block
    htp_dds = (
        _field_value_from_block_multi(gen_block, "HYDRO TEST POSITION", "HYDRO_TEST_POSITION")
        if gen_valid else None
    )
    fields.append(_dds_result(
        "HYDRO_TEST_POSITION", _prop("HYDRO_TEST_POSITION"), htp_dds,
        active=True, dds_available=gen_valid,
    ))

    # SHELL_IDP / SHELL_MOT — mechanical block, col 2 (SHELL — always active)
    shell_idp_dds = _mech_col_value(mech_block, "internal design pressure", 2) if mech_valid else None
    fields.append(_dds_result(
        "SHELL_IDP", _prop("SHELL_IDP"), shell_idp_dds,
        active=shell_active, dds_available=mech_valid,
    ))

    shell_mot_dds = _mech_col_value(mech_block, "operating", 2) if mech_valid else None
    fields.append(_dds_result(
        "SHELL_MOT", _prop("SHELL_MOT"), shell_mot_dds,
        active=shell_active, dds_available=mech_valid,
    ))

    # TUBE_IDP / TUBE_MOT — mechanical block, col 3 (TUBE)
    tube_idp_dds = _mech_col_value(mech_block, "internal design pressure", 3) if mech_valid else None
    fields.append(_dds_result(
        "TUBE_IDP", _prop("TUBE_IDP"), tube_idp_dds,
        active=tube_active, dds_available=mech_valid,
        inactive_reason=_inactive_reason("TUBE"),
    ))

    tube_mot_dds = _mech_col_value(mech_block, "operating", 3) if mech_valid else None
    fields.append(_dds_result(
        "TUBE_MOT", _prop("TUBE_MOT"), tube_mot_dds,
        active=tube_active, dds_available=mech_valid,
        inactive_reason=_inactive_reason("TUBE"),
    ))

    # JACKET_IDP / JACKET_MOT — mechanical block, col 4 (JACKET)
    jacket_idp_dds = _mech_col_value(mech_block, "internal design pressure", 4) if mech_valid else None
    fields.append(_dds_result(
        "JACKET_IDP", _prop("JACKET_IDP"), jacket_idp_dds,
        active=jacket_active, dds_available=mech_valid,
        inactive_reason=_inactive_reason("JACKET"),
    ))

    jacket_mot_dds = _mech_col_value(mech_block, "operating", 4) if mech_valid else None
    fields.append(_dds_result(
        "JACKET_MOT", _prop("JACKET_MOT"), jacket_mot_dds,
        active=jacket_active, dds_available=mech_valid,
        inactive_reason=_inactive_reason("JACKET"),
    ))

    # Drawing_Number — from METADATA block
    drawing_num_dds = (
        _field_value_from_block_multi(
            meta_block,
            "Drawing Number", "Drawing_Number", "DWG No", "Drawing No",
        ) if meta_valid else None
    )
    fields.append(_dds_result(
        "Drawing_Number", _prop("Drawing_Number"), drawing_num_dds,
        active=True, dds_available=meta_valid,
    ))

    # Tag_No — from METADATA block
    tag_no_dds = (
        _field_value_from_block_multi(meta_block, "Tag No", "Tag_No", "TagNo")
        if meta_valid else None
    )
    fields.append(_dds_result(
        "Tag_No", _prop("Tag_No"), tag_no_dds,
        active=True, dds_available=meta_valid,
    ))

    # Equipment_Type — from METADATA block
    equip_type_dds = (
        _field_value_from_block_multi(
            meta_block, "Equipment Type", "Equipment_Type", "Equipment",
        ) if meta_valid else None
    )
    fields.append(_dds_result(
        "Equipment_Type", _prop("Equipment_Type"), equip_type_dds,
        active=True, dds_available=meta_valid,
    ))

    # Design_Code — try METADATA first, then GENERAL DATA
    design_code_dds = None
    design_code_available = False
    if meta_valid:
        design_code_dds = _field_value_from_block_multi(
            meta_block, "Design Code", "Design_Code",
        )
        design_code_available = True
    if design_code_dds is None and gen_valid:
        design_code_dds = _field_value_from_block_multi(gen_block, "Design Code", "Design_Code")
        design_code_available = True
    fields.append(_dds_result(
        "Design_Code", _prop("Design_Code"), design_code_dds,
        active=True, dds_available=design_code_available,
    ))

    # Material_Code — from METADATA block
    material_code_dds = (
        _field_value_from_block_multi(
            meta_block, "Material Code", "Material_Code", "Material",
        ) if meta_valid else None
    )
    fields.append(_dds_result(
        "Material_Code", _prop("Material_Code"), material_code_dds,
        active=True, dds_available=meta_valid,
    ))

    # Inspection_By — from METADATA block
    insp_by_dds = (
        _field_value_from_block_multi(
            meta_block, "Inspection By", "Inspection_By", "Inspected By",
        ) if meta_valid else None
    )
    fields.append(_dds_result(
        "Inspection_By", _prop("Inspection_By"), insp_by_dds,
        active=True, dds_available=meta_valid,
    ))

    # Revision — from METADATA or GENERAL DATA
    revision_dds = None
    revision_available = False
    if meta_valid:
        revision_dds = _field_value_from_block_multi(meta_block, "Revision", "Rev", "REV")
        revision_available = True
    if revision_dds is None and gen_valid:
        revision_dds = _field_value_from_block_multi(gen_block, "Revision", "Rev")
        revision_available = True
    fields.append(_dds_result(
        "Revision", _prop("Revision"), revision_dds,
        active=True, dds_available=revision_available,
    ))

    # ── Rule check fields ─────────────────────────────────────────────────────

    # DrawnBy — not null, not blank, not "-"
    drawn_by = _prop("DrawnBy")
    if _is_blank(drawn_by):
        fields.append(_rule_result("DrawnBy", "hold", drawn_by, "DrawnBy is missing or blank"))
    else:
        fields.append(_rule_result("DrawnBy", "pass", drawn_by, None))

    # DrawnDate — valid date and < today
    drawn_date_raw = _prop("DrawnDate")
    drawn_date     = _parse_date(drawn_date_raw)
    if _is_blank(drawn_date_raw):
        fields.append(_rule_result("DrawnDate", "hold", drawn_date_raw, "DrawnDate is missing or blank"))
    elif drawn_date is None:
        fields.append(_rule_result("DrawnDate", "hold", drawn_date_raw,
                                   f"DrawnDate cannot be parsed as a date: {drawn_date_raw!r}"))
    elif drawn_date >= today:
        fields.append(_rule_result("DrawnDate", "fail", drawn_date_raw,
                                   f"DrawnDate {drawn_date} must be before today ({today})"))
    else:
        fields.append(_rule_result("DrawnDate", "pass", drawn_date_raw, None))

    # CheckedBy — not null, not blank, not "-"
    checked_by = _prop("CheckedBy")
    if _is_blank(checked_by):
        fields.append(_rule_result("CheckedBy", "hold", checked_by, "CheckedBy is missing or blank"))
    else:
        fields.append(_rule_result("CheckedBy", "pass", checked_by, None))

    # CheckedDate — valid date and >= DrawnDate
    checked_date_raw = _prop("CheckedDate")
    checked_date     = _parse_date(checked_date_raw)
    if _is_blank(checked_date_raw):
        fields.append(_rule_result("CheckedDate", "hold", checked_date_raw,
                                   "CheckedDate is missing or blank"))
    elif checked_date is None:
        fields.append(_rule_result("CheckedDate", "hold", checked_date_raw,
                                   f"CheckedDate cannot be parsed as a date: {checked_date_raw!r}"))
    elif drawn_date is None:
        fields.append(_rule_result("CheckedDate", "hold", checked_date_raw,
                                   "Cannot verify CheckedDate >= DrawnDate because DrawnDate is invalid"))
    elif checked_date < drawn_date:
        fields.append(_rule_result("CheckedDate", "fail", checked_date_raw,
                                   f"CheckedDate {checked_date} is before DrawnDate {drawn_date}"))
    else:
        fields.append(_rule_result("CheckedDate", "pass", checked_date_raw, None))

    # EngineeringApproval — not null, not blank, not "-"
    eng_approval = _prop("EngineeringApproval")
    if _is_blank(eng_approval):
        fields.append(_rule_result("EngineeringApproval", "hold", eng_approval,
                                   "EngineeringApproval is missing or blank"))
    else:
        fields.append(_rule_result("EngineeringApproval", "pass", eng_approval, None))

    # EngAppDate — valid date and >= CheckedDate
    eng_app_date_raw = _prop("EngAppDate")
    eng_app_date     = _parse_date(eng_app_date_raw)
    if _is_blank(eng_app_date_raw):
        fields.append(_rule_result("EngAppDate", "hold", eng_app_date_raw,
                                   "EngAppDate is missing or blank"))
    elif eng_app_date is None:
        fields.append(_rule_result("EngAppDate", "hold", eng_app_date_raw,
                                   f"EngAppDate cannot be parsed as a date: {eng_app_date_raw!r}"))
    elif checked_date is None:
        fields.append(_rule_result("EngAppDate", "hold", eng_app_date_raw,
                                   "Cannot verify EngAppDate >= CheckedDate because CheckedDate is invalid"))
    elif eng_app_date < checked_date:
        fields.append(_rule_result("EngAppDate", "fail", eng_app_date_raw,
                                   f"EngAppDate {eng_app_date} is before CheckedDate {checked_date}"))
    else:
        fields.append(_rule_result("EngAppDate", "pass", eng_app_date_raw, None))

    # ── Overall status ────────────────────────────────────────────────────────
    any_hold = any(f["status"] == "hold" for f in fields)
    any_fail = any(f["status"] == "fail" for f in fields)

    # hold takes precedence over fail
    if any_hold:
        overall = "hold"
    elif any_fail:
        overall = "fail"
    else:
        overall = "pass"

    pass_count = sum(1 for f in fields if f["status"] == "pass")
    fail_count = sum(1 for f in fields if f["status"] == "fail")
    hold_count = sum(1 for f in fields if f["status"] == "hold")

    _log(
        f"[CPVerify] Complete: overall={overall} "
        f"pass={pass_count} fail={fail_count} hold={hold_count} total={len(fields)}"
    )

    return {
        "customPropertyVerification": {
            "status":          overall,
            "equipmentConfig": equip_config_name,
            "configSource":    config_source,
            "fields":          fields,
        }
    }
