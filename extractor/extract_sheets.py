"""
extract_sheets.py — ExtractSheets()
Reads sheet names, scale, paper size, and view count from the drawing.
"""

from __future__ import annotations

SW_PAPER_SIZES = {
    0:  "A0", 1:  "A1", 2:  "A2",  3:  "A3",  4:  "A4",
    5:  "B",  6:  "C",  7:  "D",   8:  "E",
    12: "A0-Landscape", 13: "A1-Landscape",
}


def ExtractSheets(swApp, swModel, swDraw, logger) -> list:
    result = []
    try:
        sheet_names = swDraw.GetSheetNames()
        if not sheet_names:
            logger.warning("[Sheets] No sheets found")
            return result

        if not hasattr(sheet_names, "__iter__"):
            sheet_names = [sheet_names]

        for name in sheet_names:
            entry = {"sheet_name": str(name), "scale": "", "paper_size": "", "view_count": 0}
            try:
                swDraw.ActivateSheet(name)
                swSheet = swDraw.GetCurrentSheet()

                # Scale: GetScale2(True) returns scale as ratio (e.g. 0.1 = 1:10)
                try:
                    scale_ratio = swSheet.GetScale2(True)
                    if scale_ratio and scale_ratio > 0:
                        denom = round(1.0 / scale_ratio)
                        entry["scale"] = f"1:{denom}"
                except Exception:
                    pass

                # Paper size
                try:
                    size_idx = swSheet.GetSize()
                    entry["paper_size"] = SW_PAPER_SIZES.get(size_idx, f"size_{size_idx}")
                except Exception:
                    pass

                # View count
                try:
                    views = swSheet.GetViews()
                    entry["view_count"] = len(views) if views else 0
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
