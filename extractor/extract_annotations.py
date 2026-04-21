"""
extract_annotations.py — ExtractAnnotations()
Counts and samples notes, weld symbols, surface finish marks, and GD&T callouts.
Soft failure.
"""

from __future__ import annotations
from extractor._com_helper import get_com_value, sw_call, to_list, activate_sheet_and_get_current_sheet, iter_drawing_views

NOTE_TYPES = {1}
GTOL_TYPES = {5, 11}
SURFACE_TYPES = {6, 7, 22}
WELD_TYPES = {8, 28}


def ExtractAnnotations(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "notes_count":          0,
        "weld_symbols_count":   0,
        "surface_finish_count": 0,
        "gd_t_count":           0,
        "notes_sample":         [],
        "weld_symbols_sample":  [],
        "surface_finish_sample": [],
        "gd_t_sample":          [],
        "other_annotations_count": 0,
        "other_annotations_sample": [],
        "total_annotations_seen": 0,
    }
    seen_annotations = set()

    def annotation_key(ann):
        try:
            return str(get_com_value(ann, ("GetName", "Name", "GetNameForSelection")) or repr(ann))
        except Exception:
            return repr(ann)

    def specific_annotation(ann):
        for method in ("GetSpecificAnnotation2", "GetSpecificAnnotation"):
            try:
                value = sw_call(ann, method)
                if value is not None:
                    return value
            except Exception:
                pass
        return None

    def text_from_obj(obj):
        if obj is None:
            return ""
        for method in ("GetText", "Text", "GetTextAtIndex", "GetText2", "GetString"):
            try:
                value = sw_call(obj, method, 0) if method == "GetTextAtIndex" else sw_call(obj, method)
                if value:
                    return str(value).strip()
            except Exception:
                pass
        return ""

    def annotation_text(ann):
        text = text_from_obj(ann)
        if text:
            return text
        return text_from_obj(specific_annotation(ann))

    def sample_specific(ann):
        text = annotation_text(ann)
        if text:
            return text[:300]
        spec = specific_annotation(ann)
        for method in ("GetText", "Text", "GetSymbol", "Symbol", "GetFrameSymbols", "GetSurfaceFinishSymbol"):
            try:
                value = sw_call(spec, method) if spec is not None else sw_call(ann, method)
                if value:
                    return str(value)[:300]
            except Exception:
                pass
        return ""

    def consume_annotation(ann, sheet_name="", view_name=""):
        key = annotation_key(ann)
        if key in seen_annotations:
            return
        seen_annotations.add(key)
        result["total_annotations_seen"] += 1
        t = get_com_value(ann, ("GetType", "Type"))
        try:
            t_int = int(t) if t is not None else None
        except Exception:
            t_int = None
        text = annotation_text(ann)
        sample = {
            "sheet": str(sheet_name or ""),
            "view": str(view_name or ""),
            "type": t_int,
            "text": text[:500] if text else sample_specific(ann),
        }
        if t_int in WELD_TYPES:
            result["weld_symbols_count"] += 1
            if len(result["weld_symbols_sample"]) < 20:
                result["weld_symbols_sample"].append(sample)
            return
        if t_int in SURFACE_TYPES:
            result["surface_finish_count"] += 1
            if len(result["surface_finish_sample"]) < 20:
                result["surface_finish_sample"].append(sample)
            return
        if t_int in GTOL_TYPES:
            result["gd_t_count"] += 1
            if len(result["gd_t_sample"]) < 20:
                result["gd_t_sample"].append(sample)
            return
        if text or t_int in NOTE_TYPES:
            result["notes_count"] += 1
            if text and len(result["notes_sample"]) < 30:
                result["notes_sample"].append(text[:500])
            return
        result["other_annotations_count"] += 1
        if len(result["other_annotations_sample"]) < 20:
            result["other_annotations_sample"].append(sample)

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
                        view_name = str(get_com_value(view, ("Name", "GetName2", "GetName")) or "")
                        anns = to_list(sw_call(view, "GetAnnotations"))
                        if not anns:
                            continue

                        for ann in anns:
                            consume_annotation(ann, sheet_name, view_name)
                    except Exception:
                        continue
            except Exception as e:
                logger.debug(f"[Annotations] error on sheet '{sheet_name}': {e}")

    except Exception as e:
        logger.error(f"[Annotations] unexpected error: {e}")

    if result["total_annotations_seen"] == 0:
        try:
            for sheet_name, view in iter_drawing_views(swDraw):
                view_name = str(get_com_value(view, ("Name", "GetName2", "GetName")) or "")
                anns = []
                for method in ("GetAnnotations", "Annotations"):
                    try:
                        anns = to_list(sw_call(view, method))
                        if anns:
                            break
                    except Exception:
                        pass
                for ann in anns:
                    consume_annotation(ann, sheet_name, view_name)
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
                    consume_annotation(ann, sheet_name, view_name)
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

    if result["total_annotations_seen"] == 0:
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
                consume_annotation(ann, "", "model")
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

    logger.info(f"[Annotations] seen={result['total_annotations_seen']} "
                f"notes={result['notes_count']} "
                f"welds={result['weld_symbols_count']} "
                f"surface={result['surface_finish_count']} "
                f"gdt={result['gd_t_count']} "
                f"other={result['other_annotations_count']}")
    return result
