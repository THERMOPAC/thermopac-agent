"""
extract_annotations.py — ExtractAnnotations()
Counts and samples notes, weld symbols, surface finish marks, and GD&T callouts.
Soft failure.
"""

from __future__ import annotations
from extractor._com_helper import sw_call, to_list, activate_sheet_and_get_current_sheet, iter_drawing_views

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
                swDraw, swSheet = activate_sheet_and_get_current_sheet(swApp, swDraw, sheet_name, logger)
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

    def consume_annotation(ann):
        try:
            t = sw_call(ann, "GetType")
            if t == SW_ANN_NOTE:
                result["notes_count"] += 1
                if len(result["notes_sample"]) < 30:
                    for method in ("GetText", "Text"):
                        try:
                            text = str(sw_call(ann, method) or "").strip()
                            if text:
                                result["notes_sample"].append(text)
                                break
                        except Exception:
                            pass
            elif t == SW_ANN_WELD_SYMBOL:
                result["weld_symbols_count"] += 1
            elif t == SW_ANN_SURFACE:
                result["surface_finish_count"] += 1
            elif t == SW_ANN_GTOL:
                result["gd_t_count"] += 1
        except Exception:
            pass

    if (
        result["notes_count"] == 0
        and result["weld_symbols_count"] == 0
        and result["surface_finish_count"] == 0
        and result["gd_t_count"] == 0
    ):
        try:
            for _, view in iter_drawing_views(swDraw):
                anns = []
                for method in ("GetAnnotations", "Annotations"):
                    try:
                        anns = to_list(sw_call(view, method))
                        if anns:
                            break
                    except Exception:
                        pass
                for ann in anns:
                    consume_annotation(ann)
                ann = None
                for method in ("GetFirstAnnotation3", "GetFirstAnnotation2", "GetFirstAnnotation"):
                    try:
                        ann = sw_call(view, method)
                        if ann is not None:
                            break
                    except Exception:
                        pass
                count = 0
                while ann is not None and count < 500:
                    consume_annotation(ann)
                    next_ann = None
                    for method in ("GetNext3", "GetNext2", "GetNext"):
                        try:
                            next_ann = sw_call(ann, method)
                            break
                        except Exception:
                            pass
                    ann = next_ann
                    count += 1
        except Exception as e:
            logger.debug(f"[Annotations] drawing GetViews traversal failed: {e}")

    if (
        result["notes_count"] == 0
        and result["weld_symbols_count"] == 0
        and result["surface_finish_count"] == 0
        and result["gd_t_count"] == 0
    ):
        try:
            ann = None
            for method in ("GetFirstAnnotation2", "GetFirstAnnotation"):
                try:
                    ann = sw_call(swModel, method)
                    if ann is not None:
                        break
                except Exception:
                    pass
            count = 0
            while ann is not None and count < 1000:
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
                    pass
                next_ann = None
                for method in ("GetNext3", "GetNext2", "GetNext"):
                    try:
                        next_ann = sw_call(ann, method)
                        break
                    except Exception:
                        pass
                ann = next_ann
                count += 1
            logger.info(f"[Annotations] model annotation traversal count={count}")
        except Exception as e:
            logger.debug(f"[Annotations] model annotation traversal failed: {e}")

    logger.info(f"[Annotations] notes={result['notes_count']} "
                f"welds={result['weld_symbols_count']} "
                f"surface={result['surface_finish_count']} "
                f"gdt={result['gd_t_count']}")
    return result
