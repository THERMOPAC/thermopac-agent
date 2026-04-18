"""
extract_views.py — ExtractViews()
Reads view names, types, scale, and model references from all sheets.
"""

from __future__ import annotations

SW_VIEW_TYPES = {
    1: "base",
    2: "projected",
    3: "section",
    4: "detail",
    5: "auxiliary",
    6: "standard_3view",
    7: "relative",
    8: "predefined",
    9: "empty",
}


def ExtractViews(swApp, swModel, swDraw, logger) -> list:
    result = []
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
                views = swSheet.GetViews()
                if not views:
                    continue
                if not hasattr(views, "__iter__"):
                    views = [views]

                for view in views:
                    entry = {
                        "sheet":           str(sheet_name),
                        "view_name":       "",
                        "view_type":       "unknown",
                        "scale":           "",
                        "model_reference": "",
                    }
                    try:
                        entry["view_name"] = str(view.GetName2() or "")
                    except Exception:
                        pass
                    try:
                        v_type = view.Type
                        entry["view_type"] = SW_VIEW_TYPES.get(v_type, f"type_{v_type}")
                    except Exception:
                        pass
                    try:
                        scale_ratio = view.ScaleDecimal
                        if scale_ratio and scale_ratio > 0:
                            denom = round(1.0 / scale_ratio)
                            entry["scale"] = f"1:{denom}"
                    except Exception:
                        pass
                    try:
                        ref_model = view.GetReferencedModelName()
                        if ref_model:
                            import os
                            entry["model_reference"] = os.path.basename(ref_model)
                    except Exception:
                        pass

                    result.append(entry)

            except Exception as e:
                logger.debug(f"[Views] error on sheet '{sheet_name}': {e}")

    except Exception as e:
        logger.error(f"[Views] unexpected error: {e}")

    logger.info(f"[Views] {len(result)} view(s) extracted")
    return result
