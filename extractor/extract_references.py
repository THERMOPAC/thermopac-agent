"""
extract_references.py — ExtractReferences()
Lists referenced model filenames and counts broken external references.
Soft failure.
"""

from __future__ import annotations
import os


def ExtractReferences(swApp, swModel, swDraw, logger) -> dict:
    result = {
        "referenced_models":         [],
        "external_references_broken": 0,
        "total_references":           0,
    }
    try:
        # GetExternalReferences2 returns arrays of paths, statuses, feature types
        # Signature: GetExternalReferences2(CompNames, VarStatus, VarFeatureTypes)
        comp_names  = []
        var_status  = []
        var_ftype   = []
        try:
            ret = swModel.Extension.GetExternalReferences2(comp_names, var_status, var_ftype)
            # On success ret is True; comp_names, var_status populated by ref
            # In win32com these are returned as the last elements of the return tuple
            if isinstance(ret, tuple) and len(ret) >= 3:
                comp_names = ret[1] or []
                var_status = ret[2] or []
        except Exception as e:
            logger.debug(f"[References] GetExternalReferences2 error: {e}")

        if comp_names:
            if not hasattr(comp_names, "__iter__"):
                comp_names = [comp_names]
            if not hasattr(var_status, "__iter__"):
                var_status = [var_status]

            seen = set()
            for i, path in enumerate(comp_names):
                path = str(path or "")
                if not path:
                    continue
                result["total_references"] += 1
                basename = os.path.basename(path)
                if basename not in seen:
                    result["referenced_models"].append(basename)
                    seen.add(basename)

                # Status 2 = out of date / broken
                try:
                    if int(var_status[i]) == 2:
                        result["external_references_broken"] += 1
                except Exception:
                    pass

        # Fallback: GetReferencedDocuments
        if not result["referenced_models"]:
            try:
                docs = swModel.GetReferencedDocuments()
                if docs:
                    if not hasattr(docs, "__iter__"):
                        docs = [docs]
                    for doc in docs:
                        try:
                            name = str(doc.GetPathName() or "")
                            if name:
                                result["referenced_models"].append(os.path.basename(name))
                                result["total_references"] += 1
                        except Exception:
                            pass
            except Exception as e:
                logger.debug(f"[References] GetReferencedDocuments error: {e}")

    except Exception as e:
        logger.error(f"[References] unexpected error: {e}")

    logger.info(f"[References] total={result['total_references']} "
                f"broken={result['external_references_broken']} "
                f"models={len(result['referenced_models'])}")
    return result
