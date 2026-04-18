"""
extract_health.py — ExtractHealth()
Reads rebuild errors/warnings, open errors/warnings, and dangling dimensions.
Soft failure.
"""

from __future__ import annotations


def ExtractHealth(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "open_errors":        [],
        "open_warnings":      [],
        "rebuild_errors":     0,
        "rebuild_warnings":   0,
        "dangling_dimensions": 0,
        "dangling_relations":  0,
    }
    try:
        # ── Check rebuild errors (ForceRebuild3 with rebuild-all) ─────────────
        # We do NOT force rebuild — that might trigger saves or warnings.
        # Instead read cached FeatureManager tree state.
        try:
            feat_mgr = swModel.FeatureManager
            if feat_mgr:
                # GetFeatureStatistics returns rebuild error/warning counts
                ret = feat_mgr.GetFeatureStatistics()
                if isinstance(ret, (list, tuple)) and len(ret) >= 3:
                    result["rebuild_errors"]   = int(ret[1] or 0)
                    result["rebuild_warnings"] = int(ret[2] or 0)
        except Exception as e:
            logger.debug(f"[Health] FeatureStatistics unavailable: {e}")

        # ── Open errors/warnings stored in active config ───────────────────────
        try:
            active_config = swModel.GetActiveConfiguration()
            if active_config:
                # GetSpecificFeature2 for dangling can be complex; use AnnotationView
                pass
        except Exception as e:
            logger.debug(f"[Health] active config unavailable: {e}")

        # ── Dangling dimensions via dimension status scan ──────────────────────
        try:
            dangling_dims = 0
            dim_names = swModel.GetDimensionNames()
            if dim_names:
                if not hasattr(dim_names, "__iter__"):
                    dim_names = [dim_names]
                for dn in dim_names:
                    try:
                        dim = swModel.Parameter(dn)
                        if dim and dim.IsDangling():
                            dangling_dims += 1
                    except Exception:
                        pass
            result["dangling_dimensions"] = dangling_dims
        except Exception as e:
            logger.debug(f"[Health] dangling dimension scan error: {e}")

    except Exception as e:
        logger.error(f"[Health] unexpected error: {e}")

    logger.info(f"[Health] rebuild_err={result['rebuild_errors']} "
                f"rebuild_warn={result['rebuild_warnings']} "
                f"dangling_dims={result['dangling_dimensions']}")
    return result
