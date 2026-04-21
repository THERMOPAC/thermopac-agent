"""
extract_views.py — ExtractViews()
Reads view names, types, scale, and model references from all sheets.
"""

from __future__ import annotations
from extractor._com_helper import (
    get_com_value,
    iter_drawing_views,
    log_view_object_debug,
    sw_call,
    to_list,
    activate_sheet_and_get_current_sheet,
)

SW_VIEW_TYPES = {
    1: "base", 2: "section", 3: "detail", 4: "projected",
    5: "auxiliary", 6: "standard_3view", 7: "relative",
    8: "predefined", 9: "empty",
}


def _view_name(view) -> str:
    value = get_com_value(view, ("Name", "GetName2", "GetName"))
    return str(value or "").strip()


def _view_type(view) -> str:
    value = get_com_value(view, ("Type", "GetType"))
    if value is None:
        return "unknown"
    try:
        value_int = int(value)
        return SW_VIEW_TYPES.get(value_int, f"type_{value_int}")
    except Exception:
        return str(value)


def _view_scale(view) -> str:
    value = get_com_value(view, ("ScaleDecimal", "ScaleRatio"))
    try:
        if isinstance(value, (list, tuple)) and len(value) >= 2 and value[0]:
            return f"{value[0]}:{value[1]}"
        if value and value > 0:
            denom = round(1.0 / float(value))
            return f"1:{denom}"
    except Exception:
        pass
    return ""


def _view_model_reference(view) -> str:
    import os
    for value in (
        get_com_value(view, ("GetReferencedModelName", "ReferencedModelName")),
        get_com_value(view, ("ReferencedDocument", "GetReferencedDocument")),
    ):
        try:
            if not value:
                continue
            if isinstance(value, str):
                return os.path.basename(value)
            path = get_com_value(value, ("GetPathName", "PathName", "GetTitle", "Title"))
            if path:
                return os.path.basename(str(path))
        except Exception:
            pass
    return ""


def _view_entry(sheet_name: str, view, logger, label: str) -> dict:
    log_view_object_debug(view, logger, label)
    return {
        "sheet":           str(sheet_name or ""),
        "view_name":       _view_name(view),
        "view_type":       _view_type(view),
        "scale":           _view_scale(view),
        "model_reference": _view_model_reference(view),
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

                for index, view in enumerate(views):
                    result.append(_view_entry(str(sheet_name), view, logger, f"sheet/{sheet_name}/v{index}"))

            except Exception as e:
                logger.debug(f"[Views] error on sheet '{sheet_name}': {e}")

        if not result:
            for index, (sheet_name, view) in enumerate(iter_drawing_views(swDraw, sheet_names)):
                result.append(_view_entry(sheet_name, view, logger, f"GetViews/v{index}"))

        if not result:
            swView = sw_call(swDraw, "GetFirstView")
            sheet_name = ""
            view_index = 0
            while swView is not None and view_index < 100:
                entry = _view_entry(sheet_name, swView, logger, f"GetFirstView/v{view_index}")
                try:
                    if view_index == 0 and not sheet_name:
                        sheet_name = entry["view_name"]
                    entry["sheet"] = sheet_name
                except Exception:
                    pass
                result.append(entry)
                view_index += 1
                swView = sw_call(swView, "GetNextView")

    except Exception as e:
        logger.error(f"[Views] unexpected error: {e}")

    logger.info(f"[Views] {len(result)} view(s) extracted")
    return result
