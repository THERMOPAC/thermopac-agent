"""
fix_title_block.py — Replace hardcoded title block text with $PRP property links.

Usage
-----
1. Open your .slddrw file in SolidWorks (keep SolidWorks running).
2. Run this script from the local-agent directory:

       python tools/fix_title_block.py --list
       python tools/fix_title_block.py --apply

Modes
-----
--list   Scan every annotation on the active drawing sheet and print its
         current text plus its position.  Use this to inspect what is there
         before making any change.

--apply  Apply the FIELD_MAP replacements below.  For each annotation whose
         current text exactly matches a key in FIELD_MAP, the text is replaced
         with the corresponding $PRP:"FieldName" link.

         Only exact whole-string matches are replaced.  Annotations that
         already contain $PRP are left untouched.

--save   (optional, combine with --apply) Save the document after applying.

Safety
------
- Only text notes (INote) are touched.  Dimensions, balloons etc. are skipped.
- The script never saves unless --save is explicitly passed.
- A dry-run summary is printed before any change is committed.

Customise FIELD_MAP
-------------------
Edit the dictionary below.  Keys are the CURRENT hardcoded text visible in
the drawing.  Values are the exact property names from the 21-property list.
"""

from __future__ import annotations
import argparse
import sys

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed.  Run:  pip install pywin32")
    sys.exit(1)


# ── Map: current hardcoded text → target property name ────────────────────────
# Edit this to match what your title block actually shows.
# Keys   = the exact text currently in the annotation (case-sensitive).
# Values = the property name from the 21-target list.
# Annotations already containing $PRP are skipped automatically.
FIELD_MAP: dict[str, str] = {
    # Equipment & configuration
    "VERTICAL":                                 "Equipment_Type",
    "Jacketed Vessel":                          "Equipment_Configuration",
    "Vessel":                                   "Equipment_Configuration",
    "Heat Exchanger":                           "Equipment_Configuration",
    "Jacketed Vessel and Heat Exchanger":       "Equipment_Configuration",

    # Design & materials
    "ASME SEC VIII Div-1":                      "Design_Code",
    "ASME SEC VIII Div-2":                      "Design_Code",
    "TUV":                                      "Inspection_By",
    "BV":                                       "Inspection_By",
    "DNV":                                      "Inspection_By",
    "LR":                                       "Inspection_By",

    # Identification
    # "ENTER YOUR TAG NUMBER HERE":             "Tag_No",
    # "ENTER YOUR DRAWING NUMBER HERE":         "Drawing_Number",

    # Approval block — map any placeholder text you use
    # "Drawn By":                               "DrawnBy",
    # "Checked By":                             "CheckedBy",
    # "Approved By":                            "EngineeringApproval",

    # Revision
    "A":                                        "Revision",
    "B":                                        "Revision",
    "0":                                        "Revision",

    # Hydro-test position
    # "HORIZONTAL":                             "HYDRO_TEST_POSITION",
    # "VERTICAL (upright)":                     "HYDRO_TEST_POSITION",
}

# ── Properties that should use numeric $PRP links ─────────────────────────────
# (No difference in syntax; kept for documentation clarity.)
_NUMERIC_PROPS = {
    "SHELL_IDP", "SHELL_MOT",
    "TUBE_IDP",  "TUBE_MOT",
    "JACKET_IDP","JACKET_MOT",
}

SW_ANN_NOTE = 5   # swAnnotationType_e: swNote


def _prp_text(prop_name: str) -> str:
    return f'$PRP:"{prop_name}"'


def _get_annotation_text(ann) -> str:
    try:
        note = ann.GetSpecificAnnotation()
        raw  = getattr(note, "GetText", None)
        if callable(raw):
            return str(raw()).strip()
        return str(getattr(note, "Text", "")).strip()
    except Exception:
        return ""


def _set_annotation_text(ann, text: str) -> bool:
    try:
        note = ann.GetSpecificAnnotation()
        set_fn = getattr(note, "SetText", None)
        if callable(set_fn):
            set_fn(text)
            return True
        # Fallback: direct property
        note.Text = text
        return True
    except Exception as e:
        print(f"    ✗ SetText failed: {e}")
        return False


def _iter_annotations(swModel):
    """Iterate all IAnnotation objects on the active sheet (model space)."""
    ann = swModel.GetFirstAnnotation2(SW_ANN_NOTE)
    while ann is not None:
        yield ann
        try:
            ann = ann.GetNext3()
        except Exception:
            break


def _iter_sheet_format_annotations(swDraw):
    """Iterate annotations inside the sheet format (title block area)."""
    try:
        sheet = swDraw.GetCurrentSheet()
        fmt   = sheet.SheetFormatName if hasattr(sheet, "SheetFormatName") else None
    except Exception:
        fmt = None

    try:
        # Attempt to enter sheet format editing context
        swDraw.EditSheet()
    except Exception:
        pass

    ann = None
    try:
        ann_raw = swDraw.GetFirstAnnotation2(SW_ANN_NOTE)
        ann = ann_raw
    except Exception:
        pass

    collected = []
    while ann is not None:
        collected.append(ann)
        try:
            ann = ann.GetNext3()
        except Exception:
            break

    # Return to sheet (drawing) editing context
    try:
        swDraw.EditSheet()
    except Exception:
        pass

    return collected


def run(mode: str, save: bool) -> None:
    pythoncom.CoInitialize()

    # Connect to running SolidWorks session
    try:
        raw = win32com.client.GetActiveObject("SldWorks.Application.27")
    except Exception:
        try:
            raw = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception as e:
            print(f"ERROR: Cannot connect to SolidWorks. Is it running? ({e})")
            sys.exit(1)

    swApp = win32com.client.Dispatch(raw)
    print(f"[SW] Connected to SolidWorks")

    swModel = swApp.ActiveDoc
    if swModel is None:
        print("ERROR: No active document.  Open your .slddrw file first.")
        sys.exit(1)

    # Verify it is a drawing (type 3)
    try:
        raw_type = swModel.GetType
        doc_type = raw_type() if callable(raw_type) else int(raw_type)
    except Exception:
        doc_type = -1

    if doc_type != 3:
        print(f"ERROR: Active document is not a Drawing (type={doc_type}).  Open the .slddrw file first.")
        sys.exit(1)

    path = ""
    try:
        path = swModel.GetPathName()
    except Exception:
        pass
    print(f"[SW] Active drawing: {path or '(unsaved)'}")

    swDraw = swApp.ActiveDoc   # IDrawingDoc is the same COM object

    # Collect annotations from both model sheet and sheet format
    all_annotations: list = []
    try:
        ann = swModel.GetFirstAnnotation2(SW_ANN_NOTE)
        while ann is not None:
            all_annotations.append(("sheet", ann))
            try:
                ann = ann.GetNext3()
            except Exception:
                break
    except Exception as e:
        print(f"[WARN] Sheet annotation walk failed: {e}")

    print(f"[SW] Found {len(all_annotations)} Note annotations on the active sheet")

    if mode == "list":
        print("\n── Annotation listing ───────────────────────────────────────────")
        print(f"{'#':>4}  {'Source':<8}  {'Text'}")
        print("-" * 72)
        for i, (src, ann) in enumerate(all_annotations, 1):
            text = _get_annotation_text(ann)
            already = "$PRP" in text or "$PRPSHEET" in text
            flag = "  [linked]" if already else ""
            print(f"{i:>4}  {src:<8}  {text!r}{flag}")
        print()
        print("Run with --apply to replace hardcoded text with $PRP links.")
        return

    # ── Apply mode ────────────────────────────────────────────────────────────
    print("\n── Dry-run — planned replacements ──────────────────────────────────")
    planned: list[tuple] = []
    skipped_already_linked = 0
    skipped_no_match       = 0

    for src, ann in all_annotations:
        text = _get_annotation_text(ann)
        if "$PRP" in text or "$PRPSHEET" in text:
            skipped_already_linked += 1
            continue
        if text in FIELD_MAP:
            prop  = FIELD_MAP[text]
            new_t = _prp_text(prop)
            planned.append((src, ann, text, new_t))
            print(f"  REPLACE  {text!r}  →  {new_t}")
        else:
            skipped_no_match += 1

    print(f"\nTotal planned replacements : {len(planned)}")
    print(f"Already using $PRP links  : {skipped_already_linked}")
    print(f"No match in FIELD_MAP     : {skipped_no_match}")

    if not planned:
        print("\nNothing to change.  Edit FIELD_MAP in this script to add more mappings.")
        return

    confirm = input("\nApply replacements? [y/N] ").strip().lower()
    if confirm != "y":
        print("Aborted.")
        return

    print("\n── Applying ─────────────────────────────────────────────────────────")
    success = 0
    for src, ann, old_text, new_text in planned:
        ok = _set_annotation_text(ann, new_text)
        status = "✓" if ok else "✗"
        print(f"  {status}  {old_text!r}  →  {new_text}")
        if ok:
            success += 1

    print(f"\nApplied {success}/{len(planned)} replacements.")

    if save:
        try:
            swModel.Save()
            print("[SW] Drawing saved.")
        except Exception as e:
            print(f"[WARN] Save failed: {e}")
    else:
        print("[INFO] Drawing not saved (pass --save to save automatically).")

    pythoncom.CoUninitialize()


def main() -> None:
    parser = argparse.ArgumentParser(description="Fix SolidWorks title block: replace hardcoded text with $PRP links")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--list",  action="store_true", help="List all Note annotations with their current text")
    group.add_argument("--apply", action="store_true", help="Apply FIELD_MAP replacements")
    parser.add_argument("--save", action="store_true", help="Save the drawing after applying (use with --apply)")
    args = parser.parse_args()

    mode = "list" if args.list else "apply"
    run(mode, save=args.save)


if __name__ == "__main__":
    main()
