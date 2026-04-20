"""
extract_sheets.py — ExtractSheets()
Reads sheet names, scale, paper size, and view count from the drawing.
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list

SW_PAPER_SIZES = {
    0:  "A0", 1:  "A1", 2:  "A2",  3:  "A3",  4:  "A4",
    5:  "B",  6:  "C",  7:  "D",   8:  "E",
    12: "A0-Landscape", 13: "A1-Landscape",
}


def ExtractSheets(swApp, swModel, swDraw, logger) -> list:
    result = []
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        if not sheet_names:
            logger.warning("[Sheets] No sheets found")
            return result

        for name in sheet_names:
            entry = {"sheet_name": str(name), "scale": "", "paper_size": "", "view_count": 0}
            try:
                swDraw.ActivateSheet(name)
                swSheet = sw_call(swDraw, "GetCurrentSheet")
                if swSheet is None:
                    result.append(entry)
                    continue

                try:
                    scale_ratio = sw_call(swSheet, "GetScale2", True)
                    if scale_ratio and scale_ratio > 0:
                        denom = round(1.0 / scale_ratio)
                        entry["scale"] = f"1:{denom}"
                except Exception:
                    pass

                try:
                    size_idx = sw_call(swSheet, "GetSize")
                    entry["paper_size"] = SW_PAPER_SIZES.get(size_idx, f"size_{size_idx}")
                except Exception:
                    pass

                try:
                    views = to_list(sw_call(swSheet, "GetViews"))
                    entry["view_count"] = len(views)
                except Exception:
                    pass

            except Exception as e:
                logger.debug(f"[Sheets] error on sheet '{name}': {e}")

            result.append(entry)
            logger.debug(f"[Sheets] {entry}")

    except Exception as e:
        logger.error(f"[Sheets] unexpected error: {e}")

    logger.info(f"[Sheets] {len(result)} sheet(s) extracted")
    return result
