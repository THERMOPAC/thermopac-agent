"""
extract_nozzles.py — ExtractNozzles()
Searches for a nozzle schedule / nozzle table on any sheet.
Detection: General Table with title containing "nozzle" (case-insensitive).
Columns expected: Tag | Size | Rating | Service | Facing (flexible order).
Soft failure — returns found=False if no table found.
"""

from __future__ import annotations

SW_TABLE_GENERAL = 11

NOZZLE_TITLE_KEYWORDS = ("nozzle", "nozzle schedule", "connection schedule")
NOZZLE_COL_HEADERS = {
    "tag":     ("tag", "no", "#", "mark"),
    "size":    ("size", "dn", "nps", "nominal"),
    "rating":  ("rating", "class", "flange class"),
    "service": ("service", "fluid", "description", "function"),
    "facing":  ("facing", "face", "face type"),
}


def ExtractNozzles(swApp, swModel, swDraw, logger) -> dict:
    result = {"found": False, "nozzle_count": 0, "nozzles": []}
    try:
        sheet_names = swDraw.GetSheetNames()
        if not sheet_names:
            return result
        if not hasattr(sheet_names, "__iter__"):
            sheet_names = [sheet_names]

        for sheet_name in sheet_names:
            try:
                swDraw.ActivateSheet(sheet_name)
                swSheet = swDraw.GetCurrentSheet()
                table_anns = swSheet.GetTableAnnotations()
                if not table_anns:
                    continue
                if not hasattr(table_anns, "__iter__"):
                    table_anns = [table_anns]

                for ta in table_anns:
                    try:
                        if ta.Type != SW_TABLE_GENERAL:
                            continue
                        title = str(ta.Title or "").lower().strip()
                        if not any(k in title for k in NOZZLE_TITLE_KEYWORDS):
                            continue

                        # Parse column map from header row (row 0)
                        col_count = ta.ColumnCount
                        row_count = ta.RowCount
                        col_map = {}  # field_name → column_index
                        for c in range(col_count):
                            try:
                                header = str(ta.Text(0, c) or "").lower().strip()
                                for field, aliases in NOZZLE_COL_HEADERS.items():
                                    if any(a in header for a in aliases):
                                        col_map[field] = c
                                        break
                            except Exception:
                                pass

                        nozzles = []
                        for r in range(1, row_count):
                            row = {}
                            for field, col_idx in col_map.items():
                                try:
                                    row[field] = str(ta.Text(r, col_idx) or "").strip()
                                except Exception:
                                    row[field] = ""
                            if any(row.values()):
                                nozzles.append(row)

                        if nozzles:
                            result["found"]         = True
                            result["nozzle_count"]  = len(nozzles)
                            result["nozzles"]       = nozzles
                            logger.info(f"[Nozzles] Found {len(nozzles)} nozzle(s) "
                                        f"in '{sheet_name}' table '{title}'")
                            return result

                    except Exception as e:
                        logger.debug(f"[Nozzles] table error on '{sheet_name}': {e}")

            except Exception as e:
                logger.debug(f"[Nozzles] sheet error '{sheet_name}': {e}")

    except Exception as e:
        logger.error(f"[Nozzles] unexpected error: {e}")

    if not result["found"]:
        logger.info("[Nozzles] No nozzle schedule table found")
    return result
