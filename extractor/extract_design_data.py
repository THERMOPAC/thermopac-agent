"""
extract_design_data.py — ExtractDesignDataTable()  [MANDATORY]

Reads the Design Data table from an open SolidWorks drawing.
Works in both full-mode and LDR/ViewOnly mode.  Drawing General Tables
are stored in the drawing sheet, not in referenced 3-D part files, so
LDR mode has full read access to them.
If no table is found, raises DesignDataNotFoundError (hard fail).

Table detection strategy (in order):
  1. Traverse IDrawingDoc.GetViews() — returns per-sheet view arrays;
     first element of each array IS the sheet view (contains sheet-level tables).
  2. On every view call GetTableAnnotations().
  3. Score each candidate table by content-signature match against DDS field labels.
  4. Accept if score >= 3 DDS fields found in table text.
  5. Also accept if table title contains "design data" or "design_data".
  6. Fallback: first General Table whose row-0 cell contains "parameter"/"description"/"item".
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list

SW_TABLE_ANNOTATION_GENERAL  = 11
SW_TABLE_ANNOTATION_BOM      = 0
SW_TABLE_ANNOTATION_REVISION = 7

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

_DDS_SCORE_THRESHOLD = 3  # need >= this many field matches to accept a candidate


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
    return {
        "found": True,
        "rows":  rows,
    }


def _score_dds_candidate(table_ann, logger) -> int:
    """
    Score a table against DDS field-content signatures.
    Reads all cells (up to 40 rows × 4 cols), lower-cases, and counts
    how many DDS field substrings appear in the combined text.
    """
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

    score = sum(1 for sig in _DDS_FIELD_SIGNATURES if sig in text_blob)
    return score


def _find_design_data_table(swDraw, logger) -> list | None:
    """
    Primary traversal: IDrawingDoc.GetViews()
    Returns a per-sheet array of view objects; first element of each
    sub-array is the sheet view itself (carries sheet-level table annotations).
    This avoids ActivateSheet / GetCurrentSheet which are not available via
    late-bound COM dispatch on IModelDoc2.
    """
    try:
        raw_views = sw_call(swDraw, "GetViews")
        sheet_arrays = to_list(raw_views)
        if not sheet_arrays:
            logger.error("[DesignData] GetViews returned empty — no sheets found")
            return None
        logger.info(f"[DesignData] GetViews: {len(sheet_arrays)} sheet(s)")
    except Exception as e:
        logger.error(f"[DesignData] GetViews failed: {e}")
        return None

    fallback_candidate = None

    for sheet_idx, sheet_data in enumerate(sheet_arrays):
        views = to_list(sheet_data)
        if not views:
            logger.warning(f"[DesignData] Sheet {sheet_idx}: GetViews returned no views")
            continue

        logger.info(f"[DesignData] Sheet {sheet_idx}: {len(views)} view(s) (view[0]=sheet view)")

        for view_idx, swView in enumerate(views):
            if swView is None:
                continue

            view_name = "?"
            try:
                view_name = str(swView.Name or f"view_{view_idx}")
            except Exception:
                pass

            # ── GetTableAnnotations on this view ───────────────────────────
            try:
                table_anns = to_list(sw_call(swView, "GetTableAnnotations"))
            except Exception as e:
                logger.debug(f"[DesignData]   view[{view_idx}] '{view_name}': GetTableAnnotations error: {e}")
                table_anns = []

            n_tables = len(table_anns) if table_anns else 0
            logger.info(f"[DesignData]   view[{view_idx}] '{view_name}': {n_tables} table(s)")

            for tbl_idx, table_ann in enumerate(table_anns or []):
                try:
                    t_type = table_ann.Type
                except Exception:
                    t_type = -1
                try:
                    title = str(table_ann.Title or "").strip()
                except Exception:
                    title = ""
                try:
                    nrows = table_ann.RowCount
                    ncols = table_ann.ColumnCount
                except Exception:
                    nrows = ncols = "?"

                label = f"s{sheet_idx}/v{view_idx}/{title or f'tbl{tbl_idx}'}"
                logger.info(f"[DesignData]     [{label}] type={t_type} size={nrows}x{ncols} title='{title}'")

                # Primary: content-signature match
                score = _score_dds_candidate(table_ann, logger)
                logger.info(f"[DesignData]     [{label}] DDS content score={score}")
                if score >= _DDS_SCORE_THRESHOLD:
                    logger.info(f"[DesignData] DDS candidate via content-signature (score={score})")
                    rows = _parse_table(table_ann, logger, label=label)
                    if rows:
                        return rows

                # Secondary: title match
                if "design data" in title.lower() or "design_data" in title.lower():
                    logger.info(f"[DesignData] DDS candidate via title match (title='{title}')")
                    rows = _parse_table(table_ann, logger, label=label)
                    if rows:
                        return rows

                # Fallback accumulator
                if t_type == SW_TABLE_ANNOTATION_GENERAL and fallback_candidate is None:
                    try:
                        header = str(sw_call(table_ann, "Text", 0, 0) or "").lower()
                        if any(k in header for k in ("parameter", "description", "item")):
                            fallback_candidate = (table_ann, label, title)
                    except Exception:
                        pass

            # ── Annotation sweep via GetFirstAnnotation2 (if tables missed) ─
            if n_tables == 0:
                try:
                    ann = swView.GetFirstAnnotation2(0)  # 0 = all annotation types
                    sweep_count = 0
                    while ann is not None and sweep_count < 300:
                        sweep_count += 1
                        try:
                            specific = ann.GetSpecificAnnotation()
                            if specific is not None:
                                try:
                                    t_type = specific.Type
                                    title = str(specific.Title or "").strip()
                                    score = _score_dds_candidate(specific, logger)
                                    logger.info(f"[DesignData]     [sweep#{sweep_count}] type={t_type} title='{title}' score={score}")
                                    if score >= _DDS_SCORE_THRESHOLD:
                                        logger.info(f"[DesignData] DDS via annotation sweep (score={score})")
                                        rows = _parse_table(specific, logger, label=f"sweep/{title}")
                                        if rows:
                                            return rows
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        try:
                            ann = ann.GetNext5(0)
                        except Exception:
                            break
                    if sweep_count > 0:
                        logger.info(f"[DesignData]   view[{view_idx}] sweep found {sweep_count} annotation(s)")
                except Exception as e:
                    logger.debug(f"[DesignData]   view[{view_idx}] annotation sweep error: {e}")

    if fallback_candidate:
        table_ann, label, title = fallback_candidate
        logger.warning(f"[DesignData] No DDS match; using fallback general table at '{label}'")
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
        logger.warning(f"[DesignData] cannot read table dimensions for '{label}': {e}")
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
            continue

    return rows
