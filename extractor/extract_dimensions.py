"""
extract_dimensions.py — ExtractDimensions()
Counts total, driven, and tolerance dimensions. Returns a sample of the first 20.
Soft failure: errors logged, partial data returned.
"""

from __future__ import annotations


def ExtractDimensions(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "total_count":     0,
        "driven_count":    0,
        "tolerance_count": 0,
        "sample":          [],
    }
    try:
        all_dims = swModel.GetDimensionNames()
        if not all_dims:
            logger.info("[Dimensions] No dimensions found")
            return result

        if not hasattr(all_dims, "__iter__"):
            all_dims = [all_dims]

        total     = 0
        driven    = 0
        tolerance = 0
        sample    = []

        for dim_name in all_dims:
            total += 1
            try:
                dim = swModel.Parameter(dim_name)
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
        logger.error(f"[Dimensions] unexpected error: {e}")

    return result
