"""
extract_dimensions.py — ExtractDimensions()
Counts total, driven, and tolerance dimensions. Returns a sample of the first 20.
Soft failure: errors logged, partial data returned.
"""

from __future__ import annotations
from extractor._com_helper import get_com_value, sw_call, iter_drawing_views, to_list


def ExtractDimensions(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "total_count":     0,
        "driven_count":    0,
        "tolerance_count": 0,
        "sample":          [],
    }
    try:
        all_dims = sw_call(swModel, "GetDimensionNames")
        if not all_dims:
            logger.info("[Dimensions] No dimensions found")
            all_dims = []

        if not hasattr(all_dims, "__iter__"):
            all_dims = [all_dims]

        total     = 0
        driven    = 0
        tolerance = 0
        sample    = []

        for dim_name in all_dims:
            total += 1
            try:
                dim = sw_call(swModel, "Parameter", dim_name)
                if dim is None:
                    continue

                # IsDriven: True = reference/driven dimension
                try:
                    if dim.IsReference():
                        driven += 1
                except Exception:
                    pass

                # Tolerance check
                try:
                    tol = dim.GetTolerance()
                    if tol and tol.Type != 0:  # 0 = swTolNone
                        tolerance += 1
                except Exception:
                    pass

                # Build sample (first 20)
                if len(sample) < 20:
                    try:
                        val  = dim.GetSystemValue2("") if hasattr(dim, "GetSystemValue2") else None
                        unit = "mm"
                        entry = {
                            "value": round(val * 1000, 4) if val is not None else None,
                            "unit":  unit,
                            "name":  str(dim_name),
                        }
                        sample.append(entry)
                    except Exception:
                        sample.append({"name": str(dim_name), "value": None, "unit": ""})

            except Exception as e:
                logger.debug(f"[Dimensions] dim '{dim_name}' error: {e}")

        result.update({
            "total_count":     total,
            "driven_count":    driven,
            "tolerance_count": tolerance,
            "sample":          sample,
        })
        logger.info(f"[Dimensions] total={total} driven={driven} tolerance={tolerance}")

    except Exception as e:
        logger.warning(f"[Dimensions] GetDimensionNames path unavailable: {e}")

    if result["total_count"] == 0:
        try:
            total = 0
            sample = []
            view_queue = [view for _, view in iter_drawing_views(swDraw)]
            swView = view_queue.pop(0) if view_queue else sw_call(swDraw, "GetFirstView")
            view_index = 0
            while swView is not None and view_index < 100:
                display_dims = []
                for method in ("GetDisplayDimensions", "DisplayDimensions"):
                    try:
                        display_dims = to_list(sw_call(swView, method))
                        if display_dims:
                            break
                    except Exception:
                        pass
                disp_dim = None
                for method in ("GetFirstDisplayDimension5", "GetFirstDisplayDimension", "GetFirstDisplayDimension4"):
                    try:
                        disp_dim = sw_call(swView, method)
                        if disp_dim is not None:
                            break
                    except Exception:
                        pass
                linked_dims = []
                dim_index = 0
                while disp_dim is not None and dim_index < 500:
                    linked_dims.append(disp_dim)
                    next_dim = None
                    for method in ("GetNext5", "GetNext4", "GetNext"):
                        try:
                            next_dim = sw_call(disp_dim, method)
                            break
                        except Exception:
                            pass
                    disp_dim = next_dim
                    dim_index += 1
                for disp_dim in display_dims + linked_dims:
                    total += 1
                    if len(sample) < 20:
                        entry = {"name": "", "value": None, "unit": "mm"}
                        dim = get_com_value(disp_dim, ("GetDimension2", "GetDimension"), 0)
                        if dim is not None:
                            name = get_com_value(dim, ("FullName", "Name", "GetNameForSelection"))
                            if name:
                                entry["name"] = str(name)
                            val = get_com_value(dim, ("GetSystemValue2", "SystemValue"), "")
                            try:
                                entry["value"] = round(float(val) * 1000, 4) if val is not None else None
                            except Exception:
                                pass
                        else:
                            name = get_com_value(disp_dim, ("GetNameForSelection", "Name"))
                            if name:
                                entry["name"] = str(name)
                        sample.append(entry)
                if view_queue:
                    swView = view_queue.pop(0)
                else:
                    try:
                        swView = sw_call(swView, "GetNextView")
                    except Exception:
                        break
                view_index += 1
            result["total_count"] = total
            result["sample"] = sample
            logger.info(f"[Dimensions] display-dimension traversal total={total}")
        except Exception as e:
            logger.error(f"[Dimensions] display-dimension traversal failed: {e}")

    return result
