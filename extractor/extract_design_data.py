"""
extract_design_data.py — ExtractDesignDataTable()  [MANDATORY]

Reads the Design Data table from an open SolidWorks drawing.
Works in both full-mode and LDR/ViewOnly mode.  Drawing General Tables
are stored in the drawing sheet, not in referenced 3-D part files, so
LDR mode has full read access to them.
If no table is found, raises DesignDataNotFoundError (hard fail).

Table detection logic (in order):
  1. Scan all sheets → all table annotations
  2. Match any table whose Title contains "design data" (case-insensitive)
     — accepted for any table type (GENERAL, BOM, or custom)
  3. If no titled match → fall back to a General Table whose first column header
     contains "parameter", "description", or "item" (case-insensitive)
  4. Parse rows as { parameter, value, unit } triples
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list

SW_TABLE_ANNOTATION_GENERAL  = 11
SW_TABLE_ANNOTATION_BOM      = 0
SW_TABLE_ANNOTATION_REVISION = 7


class DesignDataNotFoundError(Exception):
    """Raised when no Design Data table is found in the drawing."""


def ExtractDesignDataTable(swApp, swModel, swDraw, logger) -> dict:
    """
    Returns the 'design_data_table' section of the extraction JSON.
    Always raises DesignDataNotFoundError if no table found (hard fail).
    """
    rows = _find_design_data_table(swDraw, logger)

    if not rows:
        logger.error("[DesignData] HARD FAIL — Design Data table not found in drawing")
        raise DesignDataNotFoundError(
            "Design Data table not found in drawing. "
            "Ensure the drawing contains a table with 'Design Data' in the title."
        )

    logger.info(f"[DesignData] Found {len(rows)} row(s)")
    return {
        "found": True,
        "rows":  rows,
    }


def _find_design_data_table(swDraw, logger) -> list | None:
    """Iterate all sheets and table annotations, return parsed rows or None."""
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        if not sheet_names:
            return None
    except Exception as e:
        logger.error(f"[DesignData] cannot get sheet names: {e}")
        return None

    fallback_candidate = None

    for sheet_name in sheet_names:
        try:
            swDraw.ActivateSheet(sheet_name)
            swSheet = sw_call(swDraw, "GetCurrentSheet")
            if swSheet is None:
                continue
        except Exception as e:
            logger.debug(f"[DesignData] cannot activate sheet '{sheet_name}': {e}")
            continue

        try:
            table_anns = to_list(sw_call(swSheet, "GetTableAnnotations"))
        except Exception as e:
            logger.debug(f"[DesignData] cannot get table annotations on '{sheet_name}': {e}")
            continue

        if not table_anns:
            continue

        for table_ann in table_anns:
            try:
                t_type = table_ann.Type
            except Exception:
                t_type = -1

            try:
                title = str(table_ann.Title or "").strip()
            except Exception:
                title = ""

            # Title match: accept any table type
            if "design data" in title.lower() or "design_data" in title.lower():
                rows = _parse_table(table_ann, logger, label=f"{sheet_name}/{title}")
                if rows:
                    return rows

            # Fallback: first GENERAL table whose first column header looks like parameters
            if t_type == SW_TABLE_ANNOTATION_GENERAL and fallback_candidate is None:
                try:
                    header = str(sw_call(table_ann, "Text", 0, 0) or "").lower()
                    if any(k in header for k in ("parameter", "description", "item")):
                        fallback_candidate = (table_ann, sheet_name, title)
                except Exception:
                    pass

    if fallback_candidate:
        table_ann, sheet_name, title = fallback_candidate
        logger.warning(f"[DesignData] No 'Design Data' title found; using fallback "
                       f"table on sheet '{sheet_name}' title='{title}'")
        rows = _parse_table(table_ann, logger, label=f"{sheet_name}/{title}")
        if rows:
            return rows

    return None


def _parse_table(table_ann, logger, label: str) -> list:
    rows = []
    try:
        row_count = table_ann.RowCount
        col_count = table_ann.ColumnCount
    except Exception as e:
        logger.warning(f"[DesignData] cannot read table dimensions for '{label}': {e}")
        return rows

    logger.debug(f"[DesignData] Parsing '{label}' — {row_count}R x {col_count}C")

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
            continue

    return rows
