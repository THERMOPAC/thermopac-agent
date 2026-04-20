"""
extract_design_data.py — ExtractDesignDataTable()  [MANDATORY]

Reads the Design Data table from an open SolidWorks drawing.
Works in both full-mode and LDR/ViewOnly mode.

Table detection strategy (3 paths in priority order):
  A. GetCurrentSheet() → GetTableAnnotations()
       Active sheet's ISheet object — no ActivateSheet call needed.
  B. GetFirstView() → GetNextView() sweep
       GetFirstView() returns the sheet view first (special view that holds
       all sheet-level annotations/tables before any model views).
       Iterates all views on current sheet.
  C. For all sheet names: try sw_call ActivateSheet + GetCurrentSheet fallback

Candidate scoring:
  - Content-signature: count DDS field labels in cell text. Accept if score >= 3.
  - Title match: table title contains "design data" or "design_data".
  - Last resort fallback: first General Table whose first cell looks like "Parameter".
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list

SW_TABLE_ANNOTATION_GENERAL  = 11
SW_TABLE_ANNOTATION_BOM      = 0

_DDS_FIELD_SIGNATURES = [
    "internal design pressure",
    "external design pressure",
    "design temperature",
    "hydrotest pressure",
    "capacity",
    "fluid group",
    "design code",
    "ped category",
    "serial no",
    "year of manufacture",
    "mawp",
    "operating pressure",
    "test pressure",
    "corrosion allowance",
    "design pressure",
    "working pressure",
    "insulation",
    "nozzle",
]

_DDS_SCORE_THRESHOLD = 3


class DesignDataNotFoundError(Exception):
    """Raised when no Design Data table is found in the drawing."""


def ExtractDesignDataTable(swApp, swModel, swDraw, logger) -> dict:
    rows = _find_design_data_table(swDraw, logger)
    if not rows:
        logger.error("[DesignData] HARD FAIL — Design Data table not found in drawing")
        raise DesignDataNotFoundError(
            "Design Data table not found in drawing. "
            "Ensure the drawing contains a table with 'Design Data' in the title."
        )
    logger.info(f"[DesignData] Found {len(rows)} row(s)")
    return {"found": True, "rows": rows}


def _score_dds_candidate(table_ann, logger) -> int:
    """Score a table by counting DDS field-label matches across all cells."""
    try:
        nrows = table_ann.RowCount
        ncols = table_ann.ColumnCount
    except Exception:
        return 0
    text_blob = ""
    for r in range(min(nrows, 40)):
        for c in range(min(ncols, 4)):
            try:
                cell = str(sw_call(table_ann, "Text", r, c) or "").lower()
                text_blob += " " + cell
            except Exception:
                pass
    return sum(1 for sig in _DDS_FIELD_SIGNATURES if sig in text_blob)


def _check_table(table_ann, label: str, logger, fallback_list: list) -> list | None:
    """Evaluate one table; return rows on match, update fallback_list as side-effect."""
    try:
        t_type = table_ann.Type
    except Exception:
        t_type = -1
    try:
        title = str(table_ann.Title or "").strip()
    except Exception:
        title = ""
    try:
        nrows, ncols = table_ann.RowCount, table_ann.ColumnCount
    except Exception:
        nrows = ncols = "?"

    logger.info(f"[DesignData]   [{label}] type={t_type} size={nrows}x{ncols} title='{title}'")

    # Primary: content-signature scoring
    score = _score_dds_candidate(table_ann, logger)
    logger.info(f"[DesignData]   [{label}] DDS score={score}")
    if score >= _DDS_SCORE_THRESHOLD:
        logger.info(f"[DesignData] DDS match via content-signature (score={score}, label='{label}')")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            return rows

    # Secondary: title match
    if "design data" in title.lower() or "design_data" in title.lower():
        logger.info(f"[DesignData] DDS match via title (title='{title}', label='{label}')")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            return rows

    # Accumulate fallback candidate (first general table with parameter-like first cell)
    if t_type == SW_TABLE_ANNOTATION_GENERAL and not fallback_list:
        try:
            header = str(sw_call(table_ann, "Text", 0, 0) or "").lower()
            if any(k in header for k in ("parameter", "description", "item")):
                fallback_list.append((table_ann, label, title))
        except Exception:
            pass

    return None


def _iter_view_tables(swView, path: str, logger, fallback_list: list) -> list | None:
    """Call GetTableAnnotations on a view; score each table; return rows on match."""
    if swView is None:
        return None
    try:
        view_name = str(swView.Name or "").strip() or "?"
    except Exception:
        view_name = "?"
    try:
        table_anns = to_list(sw_call(swView, "GetTableAnnotations"))
    except Exception as e:
        logger.debug(f"[DesignData] {path} GetTableAnnotations error: {e}")
        table_anns = []
    n = len(table_anns) if table_anns else 0
    logger.info(f"[DesignData] {path} '{view_name}': {n} table(s)")
    for i, ta in enumerate(table_anns or []):
        result = _check_table(ta, f"{path}/t{i}", logger, fallback_list)
        if result is not None:
            return result
    return None


def _find_design_data_table(swDraw, logger) -> list | None:
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        logger.info(f"[DesignData] Sheets: {sheet_names}")
    except Exception as e:
        logger.error(f"[DesignData] GetSheetNames failed: {e}")
        sheet_names = []

    fallback_list: list = []

    # ─── Path A: GetCurrentSheet() → GetTableAnnotations() ────────────────
    # Gets the active ISheet without ActivateSheet (which fails on IModelDoc2 dispatch).
    logger.info("[DesignData] Path A: GetCurrentSheet()")
    try:
        swSheet = sw_call(swDraw, "GetCurrentSheet")
        if swSheet is not None:
            try:
                sheet_tbl_anns = to_list(sw_call(swSheet, "GetTableAnnotations"))
                logger.info(f"[DesignData] Path A: {len(sheet_tbl_anns)} table(s) on current sheet")
                for i, ta in enumerate(sheet_tbl_anns or []):
                    result = _check_table(ta, f"A/sheet/t{i}", logger, fallback_list)
                    if result is not None:
                        return result
            except Exception as e:
                logger.warning(f"[DesignData] Path A GetTableAnnotations error: {e}")
        else:
            logger.warning("[DesignData] Path A: GetCurrentSheet returned None")
    except Exception as e:
        logger.warning(f"[DesignData] Path A error: {e}")

    # ─── Path B: GetFirstView() → GetNextView() sweep ──────────────────────
    # GetFirstView() returns the sheet view (special view) FIRST, then model views.
    # Sheet-level tables (General Tables not inside a model view) live on the sheet view.
    logger.info("[DesignData] Path B: GetFirstView/GetNextView sweep")
    try:
        swView = sw_call(swDraw, "GetFirstView")
        view_num = 0
        while swView is not None and view_num < 50:
            result = _iter_view_tables(swView, f"B/v{view_num}", logger, fallback_list)
            if result is not None:
                return result
            view_num += 1
            try:
                swView = swView.GetNextView()
            except Exception:
                break
        logger.info(f"[DesignData] Path B: traversed {view_num} view(s)")
    except Exception as e:
        logger.warning(f"[DesignData] Path B error: {e}")

    # ─── Path C: per-sheet ActivateSheet → GetCurrentSheet → view sweep ───
    # ActivateSheet(str) fails on some dispatch interfaces but try anyway for each sheet.
    logger.info("[DesignData] Path C: per-sheet ActivateSheet sweep")
    for sname in (sheet_names or []):
        try:
            swDraw.ActivateSheet(sname)
            logger.info(f"[DesignData] Path C: activated sheet '{sname}'")
            # After activation, try GetCurrentSheet
            swSheet2 = sw_call(swDraw, "GetCurrentSheet")
            if swSheet2 is not None:
                tbl_anns = to_list(sw_call(swSheet2, "GetTableAnnotations"))
                logger.info(f"[DesignData] Path C sheet '{sname}': {len(tbl_anns)} table(s)")
                for i, ta in enumerate(tbl_anns or []):
                    result = _check_table(ta, f"C/{sname}/t{i}", logger, fallback_list)
                    if result is not None:
                        return result
            # Also sweep views after activation
            swView = sw_call(swDraw, "GetFirstView")
            vnum = 0
            while swView is not None and vnum < 30:
                result = _iter_view_tables(swView, f"C/{sname}/v{vnum}", logger, fallback_list)
                if result is not None:
                    return result
                vnum += 1
                try:
                    swView = swView.GetNextView()
                except Exception:
                    break
        except Exception as e:
            logger.warning(f"[DesignData] Path C sheet '{sname}' error: {e}")

    # ─── Fallback candidate ────────────────────────────────────────────────
    if fallback_list:
        table_ann, label, title = fallback_list[0]
        logger.warning(f"[DesignData] Using fallback general table at '{label}' title='{title}'")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            return rows

    return None


def _parse_table(table_ann, logger, label: str) -> list:
    rows = []
    try:
        row_count = table_ann.RowCount
        col_count = table_ann.ColumnCount
    except Exception as e:
        logger.warning(f"[DesignData] cannot read dimensions for '{label}': {e}")
        return rows
    logger.info(f"[DesignData] Parsing '{label}' — {row_count}R x {col_count}C")
    for r in range(row_count):
        try:
            cells = []
            for c in range(col_count):
                try:
                    cells.append(str(sw_call(table_ann, "Text", r, c) or "").strip())
                except Exception:
                    cells.append("")
            if not any(cells):
                continue
            param = cells[0] if len(cells) > 0 else ""
            value = cells[1] if len(cells) > 1 else ""
            unit  = cells[2] if len(cells) > 2 else ""
            if r == 0 and param.lower() in ("parameter", "description", "item", "no", "#"):
                continue
            if not param:
                continue
            rows.append({"parameter": param, "value": value, "unit": unit})
        except Exception as e:
            logger.debug(f"[DesignData] row {r} parse error: {e}")
    return rows
