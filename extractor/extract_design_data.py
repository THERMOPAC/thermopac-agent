"""
extract_design_data.py — ExtractDesignDataTable()  [MANDATORY]

Reads the Design Data table or best-effort Design Data notes from an open SolidWorks drawing.
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
_NOTE_SCORE_THRESHOLD = 2
_TITLE_MATCHES = [
    "design data",
    "design_data",
    "process data",
    "equipment data",
    "design",
    "data",
]


class DesignDataNotFoundError(Exception):
    """Raised when no Design Data table is found in the drawing."""


def ExtractDesignDataTable(swApp, swModel, swDraw, logger) -> dict:
    warnings = []
    table_result = _find_design_data_table(swDraw, logger)
    if table_result.get("rows"):
        rows = table_result["rows"]
        logger.info(f"[DesignData] Found {len(rows)} row(s) from table")
        return {
            "found": True,
            "status": "found",
            "source": "table",
            "rows": rows,
            "table_titles_found": table_result.get("table_titles_found", []),
            "fallback_text": [],
            "warnings": table_result.get("warnings", []),
        }

    warnings.extend(table_result.get("warnings", []))
    notes_result = _find_design_data_notes(swModel, swDraw, logger)
    warnings.extend(notes_result.get("warnings", []))
    if notes_result.get("accepted_text"):
        logger.warning("[DesignData] Structured Design Data table missing; accepted likely Design Data note/text fallback")
        return {
            "found": False,
            "status": "found",
            "source": "notes",
            "rows": [],
            "table_titles_found": table_result.get("table_titles_found", []),
            "fallback_text": notes_result["accepted_text"],
            "note_candidates": notes_result.get("candidates", []),
            "warnings": warnings,
        }

    warning = "Design Data table not found or inaccessible; continuing with design_data.status=missing"
    logger.warning(f"[DesignData] {warning}")
    warnings.append(warning)
    return {
        "found": False,
        "status": "missing",
        "source": "missing",
        "rows": [],
        "table_titles_found": table_result.get("table_titles_found", []),
        "fallback_text": [],
        "note_candidates": notes_result.get("candidates", []),
        "warnings": warnings,
    }


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


def _check_table(table_ann, label: str, logger, fallback_list: list, titles_found: list) -> list | None:
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

    titles_found.append({"label": label, "type": t_type, "rows": nrows, "cols": ncols, "title": title})
    logger.info(f"[DesignData]   [{label}] table title found: '{title}' type={t_type} size={nrows}x{ncols}")

    # Primary: content-signature scoring
    score = _score_dds_candidate(table_ann, logger)
    logger.info(f"[DesignData]   [{label}] DDS score={score}")
    if score >= _DDS_SCORE_THRESHOLD:
        logger.info(f"[DesignData] DDS match via content-signature (score={score}, label='{label}')")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            logger.info(f"[DesignData]   [{label}] ACCEPT table: DDS score {score} and {len(rows)} parsed row(s)")
            return rows
        logger.info(f"[DesignData]   [{label}] REJECT table: DDS score {score} but no parseable rows")

    title_lower = title.lower()
    matched_title = next((k for k in _TITLE_MATCHES if k in title_lower), None)
    if matched_title:
        logger.info(f"[DesignData] DDS match via similar title keyword '{matched_title}' (title='{title}', label='{label}')")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            logger.info(f"[DesignData]   [{label}] ACCEPT table: similar title and {len(rows)} parsed row(s)")
            return rows
        logger.info(f"[DesignData]   [{label}] REJECT table: similar title but no parseable rows")
    else:
        logger.info(f"[DesignData]   [{label}] REJECT table: title does not match Design/Data keywords and DDS score {score} < {_DDS_SCORE_THRESHOLD}")

    # Accumulate fallback candidate (first general table with parameter-like first cell)
    if t_type == SW_TABLE_ANNOTATION_GENERAL and not fallback_list:
        try:
            header = str(sw_call(table_ann, "Text", 0, 0) or "").lower()
            if any(k in header for k in ("parameter", "description", "item")):
                fallback_list.append((table_ann, label, title))
        except Exception:
            pass

    return None


def _iter_view_tables(swView, path: str, logger, fallback_list: list, titles_found: list) -> list | None:
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
        result = _check_table(ta, f"{path}/t{i}", logger, fallback_list, titles_found)
        if result is not None:
            return result
    return None


def _find_design_data_table(swDraw, logger) -> dict:
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        logger.info(f"[DesignData] Sheets: {sheet_names}")
    except Exception as e:
        logger.error(f"[DesignData] GetSheetNames failed: {e}")
        sheet_names = []

    fallback_list: list = []
    titles_found: list = []
    warnings: list = []

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
                    result = _check_table(ta, f"A/sheet/t{i}", logger, fallback_list, titles_found)
                    if result is not None:
                        return {"rows": result, "table_titles_found": titles_found, "warnings": warnings}
            except Exception as e:
                logger.warning(f"[DesignData] Path A GetTableAnnotations error: {e}")
        else:
            logger.warning("[DesignData] Path A: GetCurrentSheet returned None")
    except Exception as e:
        msg = f"Path A GetCurrentSheet failed: {e}"
        logger.warning(f"[DesignData] {msg}")
        warnings.append(msg)

    # ─── Path B: GetFirstView() → GetNextView() sweep ──────────────────────
    # GetFirstView() returns the sheet view (special view) FIRST, then model views.
    # Sheet-level tables (General Tables not inside a model view) live on the sheet view.
    logger.info("[DesignData] Path B: GetFirstView/GetNextView sweep")
    try:
        swView = sw_call(swDraw, "GetFirstView")
        view_num = 0
        while swView is not None and view_num < 50:
            result = _iter_view_tables(swView, f"B/v{view_num}", logger, fallback_list, titles_found)
            if result is not None:
                return {"rows": result, "table_titles_found": titles_found, "warnings": warnings}
            view_num += 1
            try:
                swView = sw_call(swView, "GetNextView")
            except Exception:
                break
        logger.info(f"[DesignData] Path B: traversed {view_num} view(s)")
    except Exception as e:
        msg = f"Path B GetFirstView/GetNextView failed: {e}"
        logger.warning(f"[DesignData] {msg}")
        warnings.append(msg)

    # ─── Path C: per-sheet ActivateSheet → GetCurrentSheet → view sweep ───
    # ActivateSheet(str) fails on some dispatch interfaces but try anyway for each sheet.
    logger.info("[DesignData] Path C: per-sheet ActivateSheet sweep")
    for sname in (sheet_names or []):
        try:
            sw_call(swDraw, "ActivateSheet", sname)
            logger.info(f"[DesignData] Path C: activated sheet '{sname}'")
            # After activation, try GetCurrentSheet
            swSheet2 = sw_call(swDraw, "GetCurrentSheet")
            if swSheet2 is not None:
                tbl_anns = to_list(sw_call(swSheet2, "GetTableAnnotations"))
                logger.info(f"[DesignData] Path C sheet '{sname}': {len(tbl_anns)} table(s)")
                for i, ta in enumerate(tbl_anns or []):
                    result = _check_table(ta, f"C/{sname}/t{i}", logger, fallback_list, titles_found)
                    if result is not None:
                        return {"rows": result, "table_titles_found": titles_found, "warnings": warnings}
            # Also sweep views after activation
            swView = sw_call(swDraw, "GetFirstView")
            vnum = 0
            while swView is not None and vnum < 30:
                result = _iter_view_tables(swView, f"C/{sname}/v{vnum}", logger, fallback_list, titles_found)
                if result is not None:
                    return {"rows": result, "table_titles_found": titles_found, "warnings": warnings}
                vnum += 1
                try:
                    swView = sw_call(swView, "GetNextView")
                except Exception:
                    break
        except Exception as e:
            msg = f"Path C sheet '{sname}' failed: {e}"
            logger.warning(f"[DesignData] {msg}")
            warnings.append(msg)

    # ─── Fallback candidate ────────────────────────────────────────────────
    if fallback_list:
        table_ann, label, title = fallback_list[0]
        logger.warning(f"[DesignData] Using fallback general table at '{label}' title='{title}'")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            return {"rows": rows, "table_titles_found": titles_found, "warnings": warnings}

    if titles_found:
        logger.info(f"[DesignData] Table scan complete: {len(titles_found)} table candidate(s) found but none accepted")
    else:
        logger.info("[DesignData] Table scan complete: no table titles accessible/found")
    return {"rows": [], "table_titles_found": titles_found, "warnings": warnings}


def _find_design_data_notes(swModel, swDraw, logger) -> dict:
    candidates = []
    accepted = []
    warnings = []
    seen = set()

    def add_candidate(label: str, text: str):
        cleaned = " ".join(str(text or "").replace("\r", "\n").split())
        if not cleaned or cleaned in seen:
            return
        seen.add(cleaned)
        lower = cleaned.lower()
        score = sum(1 for sig in _DDS_FIELD_SIGNATURES if sig in lower)
        title_hit = next((k for k in _TITLE_MATCHES if k in lower), None)
        reason = ""
        accepted_candidate = False
        if title_hit and score >= 1:
            accepted_candidate = True
            reason = f"accepted: title/text keyword '{title_hit}' and DDS score {score}"
        elif score >= _NOTE_SCORE_THRESHOLD:
            accepted_candidate = True
            reason = f"accepted: DDS score {score}"
        elif title_hit:
            reason = f"rejected: keyword '{title_hit}' found but DDS score {score} < 1"
        else:
            reason = f"rejected: no Design/Data keyword and DDS score {score} < {_NOTE_SCORE_THRESHOLD}"
        sample = cleaned[:500]
        logger.info(f"[DesignData] Note/text candidate [{label}] score={score}: {reason}; text='{sample[:180]}'")
        item = {"label": label, "score": score, "reason": reason, "text": sample}
        candidates.append(item)
        if accepted_candidate:
            accepted.append(item)

    def ann_text(ann):
        for method in ("GetText", "Text"):
            try:
                value = sw_call(ann, method)
                if value:
                    return str(value)
            except Exception:
                pass
        try:
            note = sw_call(ann, "GetSpecificAnnotation2")
            if note is not None:
                for method in ("GetText", "Text", "GetTextAtIndex"):
                    try:
                        value = sw_call(note, method, 0) if method == "GetTextAtIndex" else sw_call(note, method)
                        if value:
                            return str(value)
                    except Exception:
                        pass
        except Exception:
            pass
        return ""

    logger.info("[DesignData] Note/text fallback scan starting")
    try:
        ann = sw_call(swModel, "GetFirstAnnotation")
        count = 0
        while ann is not None and count < 300:
            text = ann_text(ann)
            add_candidate(f"model/annotation/{count}", text)
            count += 1
            try:
                ann = sw_call(ann, "GetNext")
            except Exception as e:
                warnings.append(f"Model annotation GetNext failed: {e}")
                break
        logger.info(f"[DesignData] Model annotation note/text scan visited {count} annotation(s)")
    except Exception as e:
        msg = f"ModelDoc2.GetFirstAnnotation failed: {e}"
        logger.warning(f"[DesignData] {msg}")
        warnings.append(msg)

    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
    except Exception as e:
        sheet_names = []
        warnings.append(f"Note fallback GetSheetNames failed: {e}")

    for sheet_name in sheet_names:
        try:
            sw_call(swDraw, "ActivateSheet", sheet_name)
            swSheet = sw_call(swDraw, "GetCurrentSheet")
            views = to_list(sw_call(swSheet, "GetViews")) if swSheet is not None else []
            logger.info(f"[DesignData] Note/text sheet '{sheet_name}': {len(views)} view candidate(s)")
            for v_idx, view in enumerate(views):
                try:
                    anns = to_list(sw_call(view, "GetAnnotations"))
                    for a_idx, ann in enumerate(anns):
                        add_candidate(f"{sheet_name}/view{v_idx}/annotation{a_idx}", ann_text(ann))
                except Exception as e:
                    warnings.append(f"Note fallback GetAnnotations failed on sheet '{sheet_name}' view {v_idx}: {e}")
        except Exception as e:
            msg = f"Note fallback sheet '{sheet_name}' failed: {e}"
            logger.warning(f"[DesignData] {msg}")
            warnings.append(msg)

    logger.info(f"[DesignData] Note/text fallback scan complete: {len(candidates)} candidate(s), {len(accepted)} accepted")
    return {"accepted_text": accepted, "candidates": candidates, "warnings": warnings}


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
