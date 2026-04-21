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
  - Content-signature: count DDS field labels in cell text. Accept if score >= 1.
  - Title match: table title contains "design data" or "design_data".
  - Last resort fallback: first General Table whose first cell looks like "Parameter".
"""

from __future__ import annotations
import html
import re
from extractor._com_helper import get_com_value, sw_call, to_list, activate_sheet_and_get_current_sheet, iter_drawing_views

try:
    import pythoncom
except Exception:
    pythoncom = None

SW_TABLE_ANNOTATION_GENERAL_TYPES  = {0, 11}
SW_TABLE_ANNOTATION_BOM_TYPES      = {2}

MECHANICAL_DDS_TITLE_PREFIX = "DESIGN DATA SHEET"
GENERAL_DATA_TITLE = "GENERAL DATA"
METADATA_TITLE = "METADATA"
MECHANICAL_DDS_HEADERS = ["GROUP", "PARAMETER", "SHELL", "TUBE", "JACKET"]
FIELD_VALUE_HEADERS = ["FIELD", "VALUE"]

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

_DDS_SCORE_THRESHOLD = 1
_NOTE_SCORE_THRESHOLD = 2
_TITLE_MATCHES = [
    "design data",
    "design_data",
    "process data",
    "equipment data",
    "design",
    "data",
    "general table",
]


def _invoke_member(obj, name: str, *args):
    try:
        attr = getattr(obj, name)
        if callable(attr):
            return attr(*args)
        if not args:
            return attr
    except Exception:
        pass
    if pythoncom is not None:
        try:
            ole = getattr(obj, "_oleobj_", None)
            if ole is not None:
                dispid = ole.GetIDsOfNames(name)
                for flag in (pythoncom.DISPATCH_PROPERTYGET, pythoncom.DISPATCH_METHOD):
                    try:
                        return ole.Invoke(dispid, 0, flag, 1, *args)
                    except Exception:
                        pass
        except Exception:
            pass
    raise AttributeError(f"COM member {name} unavailable")


def _cell_text(table_ann, row: int, col: int) -> str:
    attempts = (
        ("Text", (row, col)),
        ("Text2", (row, col, False)),
        ("DisplayedText", (row, col)),
        ("DisplayedText2", (row, col, False)),
        ("GetCellText", (row, col)),
        ("GetCellText2", (row, col, False)),
    )
    for name, args in attempts:
        try:
            value = _invoke_member(table_ann, name, *args)
            if value is not None:
                text = str(value).strip()
                if text and not text.lower().startswith("<comobject"):
                    return text
        except Exception:
            pass
    return ""


class DesignDataNotFoundError(Exception):
    """Raised when no Design Data table is found in the drawing."""


def ExtractDesignDataTable(swApp, swModel, swDraw, logger) -> dict:
    strict_result = _extract_strict_dds_blocks(swApp, swDraw, logger)
    blocks = strict_result["dds_blocks"]
    found = any(block.get("found") for block in blocks.values())
    valid = any(block.get("status") == "valid" for block in blocks.values())
    if valid:
        logger.info("[DesignData] Strict DDS block extraction found valid table block(s)")
    elif found:
        logger.warning("[DesignData] Strict DDS block extraction found section(s), but header validation failed")
    else:
        logger.warning("[DesignData] Strict DDS block extraction found no matching DDS sections")
    return {
        "found": found,
        "status": "table" if found else "missing",
        "source": "table" if found else "missing",
        "rows": [],
        "dds_blocks": blocks,
        "table_titles_found": strict_result.get("table_titles_found", []),
        "raw_tables": strict_result.get("raw_tables", []),
        "fallback_text": [],
        "warnings": strict_result.get("warnings", []),
    }


def _strip_cell_markup(value) -> str:
    text = "" if value is None else str(value)
    text = html.unescape(text).replace("\xa0", " ")
    text = re.sub(r"<[^>]*>", "", text)
    return text.strip()


def _match_text(value: str) -> str:
    return re.sub(r"\s+", " ", _strip_cell_markup(value)).strip().upper()


def _empty_dds_block(expected_headers: list[str], status: str = "missing") -> dict:
    return {
        "found": False,
        "header_match": False,
        "status": status,
        "missing_headers": list(expected_headers),
        "row_count": 0,
        "headers": [],
        "rows": [],
        "title": "",
    }


def _title_candidates(table_title: str, rows: list[list[str]]) -> list[dict]:
    candidates = []
    if table_title:
        candidates.append({"title": table_title, "source": "table_title", "row_index": None})
    for idx, row in enumerate(rows[:5]):
        joined = " ".join(cell for cell in row if cell).strip()
        if joined:
            candidates.append({"title": joined, "source": "row", "row_index": idx})
    return candidates


def _locate_header_sequence(rows: list[list[str]], expected_headers: list[str]) -> tuple[int | None, int | None, list[str]]:
    expected_norm = [_match_text(header) for header in expected_headers]
    for row_index, row in enumerate(rows):
        row_norm = [_match_text(cell) for cell in row]
        max_start = max(0, len(row_norm) - len(expected_norm) + 1)
        for start in range(max_start):
            segment = row_norm[start:start + len(expected_norm)]
            if segment == expected_norm:
                return row_index, start, row[start:start + len(expected_norm)]
    return None, None, []


def _missing_headers(rows: list[list[str]], expected_headers: list[str]) -> list[str]:
    available = {_match_text(cell) for row in rows for cell in row if cell}
    return [header for header in expected_headers if _match_text(header) not in available]


def _rows_after_header(rows: list[list[str]], header_index: int, start_col: int, width: int) -> list[list[str]]:
    parsed = []
    for row in rows[header_index + 1:]:
        cells = row[start_col:start_col + width]
        if len(cells) < width:
            cells = cells + [""] * (width - len(cells))
        if any(cell != "" for cell in cells):
            parsed.append(cells)
    return parsed


def _build_block(rows: list[list[str]], expected_headers: list[str], title: str, found: bool) -> dict:
    header_index, start_col, headers = _locate_header_sequence(rows, expected_headers)
    header_match = header_index is not None and start_col is not None
    if not found:
        return _empty_dds_block(expected_headers)
    if not header_match:
        block = _empty_dds_block(expected_headers, status="invalid")
        block["found"] = True
        block["title"] = title
        block["missing_headers"] = _missing_headers(rows, expected_headers)
        return block
    parsed_rows = _rows_after_header(rows, header_index, start_col, len(expected_headers))
    return {
        "found": True,
        "header_match": True,
        "status": "valid",
        "missing_headers": [],
        "row_count": len(parsed_rows),
        "headers": headers,
        "rows": parsed_rows,
        "title": title,
    }


def _evaluate_table_for_strict_blocks(table_rows: list[list[str]], table_title: str, blocks: dict) -> None:
    candidates = _title_candidates(table_title, table_rows)
    title_texts = [(candidate["title"], _match_text(candidate["title"])) for candidate in candidates]

    mechanical_title = next((title for title, norm in title_texts if norm.startswith(MECHANICAL_DDS_TITLE_PREFIX)), "")
    if mechanical_title and not blocks["mechanical_design_data"].get("found"):
        blocks["mechanical_design_data"] = _build_block(
            table_rows,
            MECHANICAL_DDS_HEADERS,
            mechanical_title,
            found=True,
        )

    general_title = next((title for title, norm in title_texts if norm == GENERAL_DATA_TITLE), "")
    if general_title and not blocks["general_data"].get("found"):
        blocks["general_data"] = _build_block(
            table_rows,
            FIELD_VALUE_HEADERS,
            general_title,
            found=True,
        )

    metadata_title = next((title for title, norm in title_texts if norm == METADATA_TITLE), "")
    if metadata_title and not blocks["metadata"].get("found"):
        blocks["metadata"] = _build_block(
            table_rows,
            FIELD_VALUE_HEADERS,
            metadata_title,
            found=True,
        )


def _read_strict_table_raw(table_ann) -> list[list[str]]:
    rows = []
    try:
        row_count = table_ann.RowCount
        col_count = table_ann.ColumnCount
    except Exception:
        return rows
    for r in range(row_count):
        row = []
        for c in range(col_count):
            try:
                row.append(_strip_cell_markup(_cell_text(table_ann, r, c)))
            except Exception:
                row.append("")
        rows.append(row)
    return rows


def _scan_strict_table(table_ann, label: str, blocks: dict, titles_found: list, raw_tables: list, logger) -> None:
    try:
        t_type = table_ann.Type
    except Exception:
        t_type = -1
    try:
        title = _strip_cell_markup(table_ann.Title or "")
    except Exception:
        title = ""
    try:
        nrows, ncols = table_ann.RowCount, table_ann.ColumnCount
    except Exception:
        nrows = ncols = "?"
    rows = _read_strict_table_raw(table_ann)
    candidates = _title_candidates(title, rows)
    titles_found.append({
        "label": label,
        "type": t_type,
        "rows": nrows,
        "cols": ncols,
        "title": title,
        "title_candidates": [candidate["title"] for candidate in candidates],
    })
    raw_tables.append({"label": label, "type": t_type, "rows": nrows, "cols": ncols, "title": title, "content": rows})
    logger.info(f"[DesignData] Strict scan [{label}] title='{title}' type={t_type} size={nrows}x{ncols}")
    _evaluate_table_for_strict_blocks(rows, title, blocks)


def _extract_strict_dds_blocks(swApp, swDraw, logger) -> dict:
    blocks = {
        "mechanical_design_data": _empty_dds_block(MECHANICAL_DDS_HEADERS),
        "general_data": _empty_dds_block(FIELD_VALUE_HEADERS),
        "metadata": _empty_dds_block(FIELD_VALUE_HEADERS),
    }
    titles_found = []
    raw_tables = []
    warnings = []

    def done() -> bool:
        return all(block.get("found") for block in blocks.values())

    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
    except Exception as e:
        sheet_names = []
        warnings.append(f"GetSheetNames failed: {e}")

    try:
        swSheet = sw_call(swDraw, "GetCurrentSheet")
        if swSheet is not None:
            table_anns = to_list(sw_call(swSheet, "GetTableAnnotations"))
            for i, ta in enumerate(table_anns or []):
                _scan_strict_table(ta, f"A/current_sheet/t{i}", blocks, titles_found, raw_tables, logger)
                if done():
                    return {"dds_blocks": blocks, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
    except Exception as e:
        warnings.append(f"Current sheet strict table scan failed: {e}")

    try:
        swView = sw_call(swDraw, "GetFirstView")
        view_num = 0
        while swView is not None and view_num < 50:
            try:
                table_anns = []
                for method in ("GetTableAnnotations", "TableAnnotations"):
                    try:
                        table_anns = to_list(sw_call(swView, method))
                        if table_anns:
                            break
                    except Exception:
                        pass
                for i, ta in enumerate(table_anns or []):
                    _scan_strict_table(ta, f"B/v{view_num}/t{i}", blocks, titles_found, raw_tables, logger)
                    if done():
                        return {"dds_blocks": blocks, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
            except Exception as e:
                warnings.append(f"View {view_num} strict table scan failed: {e}")
            view_num += 1
            try:
                swView = sw_call(swView, "GetNextView")
            except Exception:
                break
    except Exception as e:
        warnings.append(f"GetFirstView strict table scan failed: {e}")

    for sname in (sheet_names or []):
        try:
            swDraw, swSheet = activate_sheet_and_get_current_sheet(swApp, swDraw, sname, logger)
            if swSheet is not None:
                table_anns = to_list(sw_call(swSheet, "GetTableAnnotations"))
                for i, ta in enumerate(table_anns or []):
                    _scan_strict_table(ta, f"C/{sname}/t{i}", blocks, titles_found, raw_tables, logger)
                    if done():
                        return {"dds_blocks": blocks, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
        except Exception as e:
            warnings.append(f"Sheet '{sname}' strict table scan failed: {e}")

    for vnum, (_, view) in enumerate(iter_drawing_views(swDraw, sheet_names)):
        try:
            table_anns = to_list(sw_call(view, "GetTableAnnotations"))
            for i, ta in enumerate(table_anns or []):
                _scan_strict_table(ta, f"D/GetViews/v{vnum}/t{i}", blocks, titles_found, raw_tables, logger)
                if done():
                    return {"dds_blocks": blocks, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
        except Exception as e:
            warnings.append(f"GetViews view {vnum} strict table scan failed: {e}")

    return {"dds_blocks": blocks, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}


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
                cell = _cell_text(table_ann, r, c).lower()
                text_blob += " " + cell
            except Exception:
                pass
    return sum(1 for sig in _DDS_FIELD_SIGNATURES if sig in text_blob)


def _read_table_raw(table_ann, logger, label: str, max_rows: int = 30, max_cols: int = 12) -> list:
    rows = []
    try:
        row_count = table_ann.RowCount
        col_count = table_ann.ColumnCount
    except Exception:
        return rows
    for r in range(min(row_count, max_rows)):
        cells = []
        for c in range(min(col_count, max_cols)):
            try:
                cells.append(_cell_text(table_ann, r, c))
            except Exception:
                cells.append("")
        if any(cells):
            rows.append(cells)
    return rows


def _check_table(table_ann, label: str, logger, fallback_list: list, titles_found: list, raw_tables: list) -> list | None:
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

    raw_rows = _read_table_raw(table_ann, logger, label)
    titles_found.append({"label": label, "type": t_type, "rows": nrows, "cols": ncols, "title": title, "raw_preview": raw_rows[:8]})
    raw_tables.append({"label": label, "type": t_type, "rows": nrows, "cols": ncols, "title": title, "content": raw_rows})
    logger.info(f"[DesignData]   [{label}] table title found: '{title}' type={t_type} size={nrows}x{ncols}")

    # Primary: content-signature scoring
    score = _score_dds_candidate(table_ann, logger)
    logger.info(f"[DesignData]   [{label}] DDS score={score}")
    is_general = t_type in SW_TABLE_ANNOTATION_GENERAL_TYPES or "general table" in title.lower()
    is_bom = t_type in SW_TABLE_ANNOTATION_BOM_TYPES or "bom" in title.lower() or "bill of material" in title.lower()
    if score >= _DDS_SCORE_THRESHOLD and not is_bom:
        logger.info(f"[DesignData] DDS match via content-signature (score={score}, label='{label}')")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            logger.info(f"[DesignData]   [{label}] ACCEPT table: DDS score {score} and {len(rows)} parsed row(s)")
            return rows
        logger.info(f"[DesignData]   [{label}] REJECT table: DDS score {score} but no parseable rows")

    title_lower = title.lower()
    matched_title = next((k for k in _TITLE_MATCHES if k in title_lower), None)
    if matched_title and not is_bom:
        logger.info(f"[DesignData] DDS match via similar title keyword '{matched_title}' (title='{title}', label='{label}')")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            logger.info(f"[DesignData]   [{label}] ACCEPT table: similar title and {len(rows)} parsed row(s)")
            return rows
        logger.info(f"[DesignData]   [{label}] REJECT table: similar title but no parseable rows")
    else:
        logger.info(f"[DesignData]   [{label}] REJECT table: title does not match Design/Data keywords and DDS score {score} < {_DDS_SCORE_THRESHOLD}")

    # Accumulate fallback candidate (first general table with parameter-like first cell)
    if is_general and not fallback_list:
        try:
            header = _cell_text(table_ann, 0, 0).lower()
            blob = " ".join(" ".join(row) for row in raw_rows).lower()
            if any(k in blob for k in ("parameter", "description", "item", "design", "pressure", "temperature", "capacity", "material", "shell", "dish")):
                fallback_list.append((table_ann, label, title))
        except Exception:
            pass

    return None


def _iter_view_tables(swView, path: str, logger, fallback_list: list, titles_found: list, raw_tables: list) -> list | None:
    """Call GetTableAnnotations on a view; score each table; return rows on match."""
    if swView is None:
        return None
    try:
        view_name = str(get_com_value(swView, ("Name", "GetName2", "GetName")) or "").strip() or "?"
    except Exception:
        view_name = "?"
    try:
        table_anns = []
        for method in ("GetTableAnnotations", "TableAnnotations"):
            try:
                table_anns = to_list(sw_call(swView, method))
                if table_anns:
                    break
            except Exception:
                pass
    except Exception as e:
        logger.debug(f"[DesignData] {path} GetTableAnnotations error: {e}")
        table_anns = []
    n = len(table_anns) if table_anns else 0
    logger.info(f"[DesignData] {path} '{view_name}': {n} table(s)")
    for i, ta in enumerate(table_anns or []):
        result = _check_table(ta, f"{path}/t{i}", logger, fallback_list, titles_found, raw_tables)
        if result is not None:
            return result
    return None


def _find_design_data_table(swApp, swDraw, logger) -> dict:
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        logger.info(f"[DesignData] Sheets: {sheet_names}")
    except Exception as e:
        logger.error(f"[DesignData] GetSheetNames failed: {e}")
        sheet_names = []

    fallback_list: list = []
    titles_found: list = []
    raw_tables: list = []
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
                    result = _check_table(ta, f"A/sheet/t{i}", logger, fallback_list, titles_found, raw_tables)
                    if result is not None:
                        return {"rows": result, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
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
            result = _iter_view_tables(swView, f"B/v{view_num}", logger, fallback_list, titles_found, raw_tables)
            if result is not None:
                return {"rows": result, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
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
            swDraw, swSheet2 = activate_sheet_and_get_current_sheet(swApp, swDraw, sname, logger)
            logger.info(f"[DesignData] Path C: activated sheet '{sname}'")
            if swSheet2 is not None:
                tbl_anns = to_list(sw_call(swSheet2, "GetTableAnnotations"))
                logger.info(f"[DesignData] Path C sheet '{sname}': {len(tbl_anns)} table(s)")
                for i, ta in enumerate(tbl_anns or []):
                    result = _check_table(ta, f"C/{sname}/t{i}", logger, fallback_list, titles_found, raw_tables)
                    if result is not None:
                        return {"rows": result, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
            # Also sweep views after activation
            swView = sw_call(swDraw, "GetFirstView")
            vnum = 0
            while swView is not None and vnum < 30:
                result = _iter_view_tables(swView, f"C/{sname}/v{vnum}", logger, fallback_list, titles_found, raw_tables)
                if result is not None:
                    return {"rows": result, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}
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
    for vnum, (_, view) in enumerate(iter_drawing_views(swDraw, sheet_names)):
        result = _iter_view_tables(view, f"D/GetViews/v{vnum}", logger, fallback_list, titles_found, raw_tables)
        if result is not None:
            return {"rows": result, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}

    if fallback_list:
        table_ann, label, title = fallback_list[0]
        logger.warning(f"[DesignData] Using fallback general table at '{label}' title='{title}'")
        rows = _parse_table(table_ann, logger, label=label)
        if rows:
            return {"rows": rows, "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}

    if titles_found:
        logger.info(f"[DesignData] Table scan complete: {len(titles_found)} table candidate(s) found but none accepted")
    else:
        logger.info("[DesignData] Table scan complete: no table titles accessible/found")
    return {"rows": [], "table_titles_found": titles_found, "raw_tables": raw_tables, "warnings": warnings}


def _find_design_data_notes(swApp, swModel, swDraw, logger) -> dict:
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
            swDraw, swSheet = activate_sheet_and_get_current_sheet(swApp, swDraw, sheet_name, logger)
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

    for v_idx, view in enumerate([v for _, v in iter_drawing_views(swDraw, sheet_names)]):
        try:
            anns = to_list(sw_call(view, "GetAnnotations"))
            for a_idx, ann in enumerate(anns):
                add_candidate(f"GetViews/view{v_idx}/annotation{a_idx}", ann_text(ann))
        except Exception as e:
            warnings.append(f"Note fallback GetViews annotations failed on view {v_idx}: {e}")

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
                    cells.append(_cell_text(table_ann, r, c))
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
            rows.append({"parameter": param, "value": value, "unit": unit, "raw_cells": cells})
        except Exception as e:
            logger.debug(f"[DesignData] row {r} parse error: {e}")
    return rows
