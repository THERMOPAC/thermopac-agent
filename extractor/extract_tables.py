"""
extract_tables.py — ExtractTables()
Reads BOM, Revision Table, and General Tolerance Table presence.
Soft failure.
"""

from __future__ import annotations

SW_TABLE_BOM       = 0
SW_TABLE_GENERAL   = 11
SW_TABLE_REVISION  = 7


def ExtractTables(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "bom_found":                    False,
        "bom_rows":                     0,
        "revision_table_found":         False,
        "revision_rows":                [],
        "general_tolerance_table_found": False,
    }
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
                        t_type = ta.Type
                    except Exception:
                        continue

                    if t_type == SW_TABLE_BOM:
                        result["bom_found"] = True
                        try:
                            result["bom_rows"] = max(result["bom_rows"], ta.RowCount - 1)
                        except Exception:
                            pass

                    elif t_type == SW_TABLE_REVISION:
                        result["revision_table_found"] = True
                        try:
                            row_count = ta.RowCount
                            col_count = ta.ColumnCount
                            for r in range(1, row_count):  # skip header row 0
                                row_data = {}
                                labels = ["rev", "date", "description", "by"]
                                for c in range(min(col_count, 4)):
                                    try:
                                        row_data[labels[c]] = str(ta.Text(r, c) or "").strip()
                                    except Exception:
                                        row_data[labels[c]] = ""
                                if any(row_data.values()):
                                    result["revision_rows"].append(row_data)
                        except Exception as e:
                            logger.debug(f"[Tables] revision table parse error: {e}")

                    elif t_type == SW_TABLE_GENERAL:
                        try:
                            title = str(ta.Title or "").lower()
                            if "tolerance" in title or "general tol" in title:
                                result["general_tolerance_table_found"] = True
                        except Exception:
                            pass

            except Exception as e:
                logger.debug(f"[Tables] error on sheet '{sheet_name}': {e}")

    except Exception as e:
        logger.error(f"[Tables] unexpected error: {e}")

    logger.info(f"[Tables] bom={result['bom_found']} "
                f"revision={result['revision_table_found']} "
                f"tolerance={result['general_tolerance_table_found']}")
    return result
