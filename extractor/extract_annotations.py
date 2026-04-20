"""
extract_annotations.py — ExtractAnnotations()
Counts and samples notes, weld symbols, surface finish marks, and GD&T callouts.
Soft failure.
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list

SW_ANN_NOTE         = 5
SW_ANN_WELD_SYMBOL  = 28
SW_ANN_SURFACE      = 22
SW_ANN_GTOL         = 11


def ExtractAnnotations(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "notes_count":          0,
        "weld_symbols_count":   0,
        "surface_finish_count": 0,
        "gd_t_count":           0,
        "notes_sample":         [],
    }
    try:
        sheet_names = to_list(sw_call(swDraw, "GetSheetNames"))
        if not sheet_names:
            return result

        for sheet_name in sheet_names:
            try:
                swDraw.ActivateSheet(sheet_name)
                swSheet = sw_call(swDraw, "GetCurrentSheet")
                if swSheet is None:
                    continue
                ann_views = to_list(sw_call(swSheet, "GetViews"))
                if not ann_views:
                    continue

                for view in ann_views:
                    try:
                        anns = to_list(sw_call(view, "GetAnnotations"))
                        if not anns:
                            continue

                        for ann in anns:
                            try:
                                t = sw_call(ann, "GetType")
                                if t == SW_ANN_NOTE:
                                    result["notes_count"] += 1
                                    if len(result["notes_sample"]) < 30:
                                        try:
                                            text = str(sw_call(ann, "GetText") or "").strip()
                                            if text:
                                                result["notes_sample"].append(text)
                                        except Exception:
                                            pass
                                elif t == SW_ANN_WELD_SYMBOL:
                                    result["weld_symbols_count"] += 1
                                elif t == SW_ANN_SURFACE:
                                    result["surface_finish_count"] += 1
                                elif t == SW_ANN_GTOL:
                                    result["gd_t_count"] += 1
                            except Exception:
                                continue
                    except Exception:
                        continue
            except Exception as e:
                logger.debug(f"[Annotations] error on sheet '{sheet_name}': {e}")

    except Exception as e:
        logger.error(f"[Annotations] unexpected error: {e}")

    logger.info(f"[Annotations] notes={result['notes_count']} "
                f"welds={result['weld_symbols_count']} "
                f"surface={result['surface_finish_count']} "
                f"gdt={result['gd_t_count']}")
    return result
