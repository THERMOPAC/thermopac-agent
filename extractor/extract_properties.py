"""
extract_properties.py — ExtractProperties()
Reads document-level and custom properties from an open SolidWorks drawing.
"""

from __future__ import annotations
import datetime
from typing import Any


def ExtractProperties(swApp, swModel, logger) -> dict:
    """
    Returns a dict matching the 'properties' section of the extraction JSON schema.
    Never raises — logs errors and returns partial data.
    """
    result = {
        "drawing_number":    "",
        "revision":          "",
        "title":             "",
        "description":       "",
        "author":            "",
        "created_date":      "",
        "last_saved_date":   "",
        "solidworks_version": "",
        "custom_properties": {},
    }

    try:
        # ── Standard file summary info ────────────────────────────────────────
        try:
            # GetFileSummaryInfo returns IFileSummaryInfo interface
            summaryInfo = swModel.GetSummaryInfo()
            if summaryInfo:
                result["title"]       = _safe_str(summaryInfo.Title)
                result["description"] = _safe_str(summaryInfo.Subject)
                result["author"]      = _safe_str(summaryInfo.Author)
        except Exception as e:
            logger.debug(f"[Properties] summary info unavailable: {e}")

        # ── SolidWorks version ────────────────────────────────────────────────
        try:
            result["solidworks_version"] = _safe_str(swApp.RevisionNumber())
        except Exception as e:
            logger.debug(f"[Properties] SW version unavailable: {e}")

        # ── Custom properties (configuration-independent: "") ─────────────────
        try:
            mgr = swModel.Extension.CustomPropertyManager("")
            names = mgr.GetNames()
            if names:
                for name in names:
                    try:
                        # Get5: (retval, valOut, resolvedValOut, wasResolved, linkToProp)
                        ret = mgr.Get5(name, False)
                        # ret is a tuple; resolved value is index 2
                        resolved = ret[2] if isinstance(ret, (list, tuple)) and len(ret) > 2 else ""
                        result["custom_properties"][name] = _safe_str(resolved)
                    except Exception as ex:
                        logger.debug(f"[Properties] cannot read property '{name}': {ex}")
        except Exception as e:
            logger.debug(f"[Properties] custom property manager error: {e}")

        # ── Derive key fields from custom properties ──────────────────────────
        cp = result["custom_properties"]
        _map = {
            "drawing_number": ["DrwNo", "DrawingNo", "DrawingNumber", "DWG No",
                                "DWG_NO", "Drawing Number"],
            "revision":       ["Rev", "Revision", "REV", "REVISION"],
            "title":          ["Title", "TITLE", "Description"],
            "author":         ["Author", "DrawnBy", "Drawn By", "AUTHOR"],
        }
        for field, keys in _map.items():
            if not result[field]:
                for k in keys:
                    if k in cp and cp[k]:
                        result[field] = cp[k]
                        break

        # ── File dates via GetPathName + os.stat ─────────────────────────────
        try:
            import os
            path = swModel.GetPathName()
            if path and os.path.exists(path):
                stat = os.stat(path)
                result["created_date"]    = _format_ts(stat.st_ctime)
                result["last_saved_date"] = _format_ts(stat.st_mtime)
        except Exception as e:
            logger.debug(f"[Properties] file dates unavailable: {e}")

    except Exception as e:
        logger.error(f"[Properties] unexpected error: {e}")

    logger.info(f"[Properties] drawing_number={result['drawing_number']!r} "
                f"revision={result['revision']!r} "
                f"custom_count={len(result['custom_properties'])}")
    return result


def _safe_str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _format_ts(ts: float) -> str:
    try:
        return datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
    except Exception:
        return ""
