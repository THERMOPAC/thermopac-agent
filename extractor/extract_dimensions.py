"""
extract_dimensions.py — ExtractDimensions()
Counts total, driven, and tolerance dimensions. Returns a sample of the first 20.
Soft failure: errors logged, partial data returned.
"""

from __future__ import annotations
from extractor._com_helper import get_com_value, sw_call, iter_drawing_views, to_list


def _bool_value(value) -> bool:
    try:
        return bool(value() if callable(value) else value)
    except Exception:
        return False


def _dimension_name(dim, fallback="") -> str:
    value = get_com_value(dim, ("FullName", "Name", "GetNameForSelection"))
    return str(value or fallback or "")


def _dimension_value_mm(dim):
    value = get_com_value(dim, ("GetSystemValue2", "SystemValue"), "")
    try:
        return round(float(value) * 1000, 4) if value is not None else None
    except Exception:
        return None


def _dimension_tolerance(dim) -> dict:
    out = {"type": None, "min_value_mm": None, "max_value_mm": None, "text": ""}
    tol = get_com_value(dim, ("GetTolerance", "Tolerance"))
    if tol is None:
        return out
    t = get_com_value(tol, ("Type", "GetType"))
    try:
        out["type"] = int(t) if t is not None else None
    except Exception:
        out["type"] = str(t) if t is not None else None
    for key, names in (
        ("min_value_mm", ("MinValue", "GetMinValue", "LowerLimit", "GetLowerLimit")),
        ("max_value_mm", ("MaxValue", "GetMaxValue", "UpperLimit", "GetUpperLimit")),
    ):
        value = get_com_value(tol, names)
        try:
            out[key] = round(float(value) * 1000, 4) if value is not None else None
        except Exception:
            pass
    text = get_com_value(tol, ("GetText", "Text", "DisplayValue"))
    if text:
        out["text"] = str(text)
    return out


def _is_driven(dim, disp_dim=None) -> bool:
    for obj in (dim, disp_dim):
        if obj is None:
            continue
        for name in ("IsReference", "Driven", "IsDriven", "Reference"):
            try:
                value = getattr(obj, name)
                if _bool_value(value):
                    return True
            except Exception:
                pass
    return False


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
                            "view": "",
                            "sheet": "",
                        }
                        entry["driven"] = _is_driven(dim)
                        entry["tolerance"] = _dimension_tolerance(dim)
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
            driven = 0
            tolerance = 0
            sample = []
            view_queue = iter_drawing_views(swDraw)
            current = view_queue.pop(0) if view_queue else ("", sw_call(swDraw, "GetFirstView"))
            sheet_name, swView = current
            view_index = 0
            while swView is not None and view_index < 100:
                view_name = str(get_com_value(swView, ("Name", "GetName2", "GetName")) or "")
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
                seen_dims = set()
                for disp_dim in display_dims + linked_dims:
                    dim_key = repr(disp_dim)
                    if dim_key in seen_dims:
                        continue
                    seen_dims.add(dim_key)
                    total += 1
                    dim = get_com_value(disp_dim, ("GetDimension2", "GetDimension"), 0)
                    is_driven = _is_driven(dim, disp_dim)
                    tol = _dimension_tolerance(dim) if dim is not None else {"type": None, "min_value_mm": None, "max_value_mm": None, "text": ""}
                    if is_driven:
                        driven += 1
                    if tol.get("type") not in (None, 0, "0"):
                        tolerance += 1
                    if len(sample) < 20:
                        entry = {
                            "name": "",
                            "value": None,
                            "unit": "mm",
                            "sheet": str(sheet_name or ""),
                            "view": view_name,
                            "driven": is_driven,
                            "tolerance": tol,
                        }
                        if dim is not None:
                            entry["name"] = _dimension_name(dim)
                            entry["value"] = _dimension_value_mm(dim)
                        else:
                            name = get_com_value(disp_dim, ("GetNameForSelection", "Name"))
                            if name:
                                entry["name"] = str(name)
                        sample.append(entry)
                if view_queue:
                    sheet_name, swView = view_queue.pop(0)
                else:
                    try:
                        swView = sw_call(swView, "GetNextView")
                    except Exception:
                        break
                view_index += 1
            result["total_count"] = total
            result["driven_count"] = driven
            result["tolerance_count"] = tolerance
            result["sample"] = sample
            logger.info(f"[Dimensions] display-dimension traversal total={total} driven={driven} tolerance={tolerance}")
        except Exception as e:
            logger.error(f"[Dimensions] display-dimension traversal failed: {e}")

    return result
