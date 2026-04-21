"""
extract_views.py — ExtractViews()
Reads view names, types, scale, and model references from all sheets.
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list, activate_sheet_and_get_current_sheet, iter_drawing_views

SW_VIEW_TYPES = {
    1: "base", 2: "projected", 3: "section", 4: "detail",
    5: "auxiliary", 6: "standard_3view", 7: "relative",
    8: "predefined", 9: "empty",
}


def ExtractViews(swApp, swModel, swDraw, logger) -> list:
    result = []
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        if not sheet_names:
            return result

        for sheet_name in sheet_names:
            try:
                swDraw, swSheet = activate_sheet_and_get_current_sheet(swApp, swDraw, sheet_name, logger)
                if swSheet is None:
                    continue
                views = to_list(sw_call(swSheet, "GetViews"))
                if not views:
                    continue

                for view in views:
                    entry = {
                        "sheet":           str(sheet_name),
                        "view_name":       "",
                        "view_type":       "unknown",
                        "scale":           "",
                        "model_reference": "",
                    }
                    try:
                        entry["view_name"] = str(sw_call(view, "GetName2") or "")
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
                        ref_model = sw_call(view, "GetReferencedModelName")
                        if ref_model:
                            import os
                            entry["model_reference"] = os.path.basename(ref_model)
                    except Exception:
                        pass

                    result.append(entry)

            except Exception as e:
                logger.debug(f"[Views] error on sheet '{sheet_name}': {e}")

        if not result:
            for sheet_name, view in iter_drawing_views(swDraw, sheet_names):
                entry = {
                    "sheet":           str(sheet_name),
                    "view_name":       "",
                    "view_type":       "unknown",
                    "scale":           "",
                    "model_reference": "",
                }
                try:
                    entry["view_name"] = str(sw_call(view, "GetName2") or getattr(view, "Name", "") or "")
                except Exception:
                    pass
                try:
                    v_type = getattr(view, "Type", None)
                    entry["view_type"] = SW_VIEW_TYPES.get(v_type, f"type_{v_type}")
                except Exception:
                    pass
                try:
                    scale_ratio = getattr(view, "ScaleDecimal", None)
                    if scale_ratio and scale_ratio > 0:
                        denom = round(1.0 / scale_ratio)
                        entry["scale"] = f"1:{denom}"
                except Exception:
                    pass
                try:
                    ref_model = sw_call(view, "GetReferencedModelName")
                    if ref_model:
                        import os
                        entry["model_reference"] = os.path.basename(ref_model)
                except Exception:
                    pass
                result.append(entry)

        if not result:
            swView = sw_call(swDraw, "GetFirstView")
            sheet_name = ""
            view_index = 0
            while swView is not None and view_index < 100:
                entry = {
                    "sheet": sheet_name,
                    "view_name": "",
                    "view_type": "unknown",
                    "scale": "",
                    "model_reference": "",
                }
                try:
                    entry["view_name"] = str(sw_call(swView, "GetName2") or getattr(swView, "Name", "") or "")
                except Exception:
                    pass
                try:
                    if view_index == 0 and not sheet_name:
                        sheet_name = entry["view_name"]
                    entry["sheet"] = sheet_name
                except Exception:
                    pass
                try:
                    v_type = getattr(swView, "Type", None)
                    entry["view_type"] = SW_VIEW_TYPES.get(v_type, f"type_{v_type}")
                except Exception:
                    pass
                try:
                    ref_model = sw_call(swView, "GetReferencedModelName")
                    if ref_model:
                        import os
                        entry["model_reference"] = os.path.basename(ref_model)
                except Exception:
                    pass
                result.append(entry)
                view_index += 1
                swView = sw_call(swView, "GetNextView")

    except Exception as e:
        logger.error(f"[Views] unexpected error: {e}")

    logger.info(f"[Views] {len(result)} view(s) extracted")
    return result
