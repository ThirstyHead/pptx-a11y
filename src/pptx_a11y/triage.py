"""Interactive terminal triage workflow for PowerPoint presentation remediation."""
import re
from pathlib import Path
from typing import Callable, Optional
from pptx import Presentation
from pptx.presentation import Presentation as PresentationType
from pptx.oxml import parse_xml
from pptx.util import Inches
from .audit import audit_file


def _parse_slide_idx(location: str) -> Optional[int]:
    m = re.search(r"Slide\s+(\d+)", location, re.IGNORECASE)
    return int(m.group(1)) if m else None


def _parse_shape_name_or_id(location: str) -> tuple[Optional[str], Optional[int]]:
    name_m = re.search(r"Shape\s+'([^']+)'", location)
    id_m = re.search(r"\(ID\s+(\d+)\)", location)
    name = name_m.group(1) if name_m else None
    s_id = int(id_m.group(1)) if id_m else None
    return name, s_id


def _find_target_shape(slide, name: Optional[str], s_id: Optional[int]):
    for shape in slide.shapes:
        if s_id is not None and shape.shape_id == s_id:
            return shape
        if name and shape.name == name:
            return shape
    return None


def _set_shape_alt_text(prs: PresentationType, location: str, alt_text: str):
    s_idx = _parse_slide_idx(location)
    if not s_idx or s_idx > len(prs.slides):
        return
    slide = prs.slides[s_idx - 1]
    name, s_id = _parse_shape_name_or_id(location)
    shape = _find_target_shape(slide, name, s_id)
    if shape:
        cNvPr = shape._element.xpath(".//p:cNvPr")
        if cNvPr:
            cNvPr[0].set("descr", alt_text)


def _mark_shape_decorative(prs: PresentationType, location: str):
    s_idx = _parse_slide_idx(location)
    if not s_idx or s_idx > len(prs.slides):
        return
    slide = prs.slides[s_idx - 1]
    name, s_id = _parse_shape_name_or_id(location)
    shape = _find_target_shape(slide, name, s_id)
    if shape:
        cNvPr = shape._element.xpath(".//p:cNvPr")
        if cNvPr:
            cNvPr[0].set("descr", "")
            dec_elem = parse_xml(
                '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2021/oembed" val="1"/>'
            )
            cNvPr[0].append(dec_elem)


def _set_slide_title(prs: PresentationType, location: str, title_text: str):
    s_idx = _parse_slide_idx(location)
    if not s_idx or s_idx > len(prs.slides):
        return
    slide = prs.slides[s_idx - 1]
    if slide.shapes.title:
        slide.shapes.title.text = title_text
    else:
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
        tb.text_frame.text = title_text


def run_interactive_triage(
    in_path: str | Path,
    out_path: Optional[str | Path] = None,
    input_func: Optional[Callable[[str], str]] = None,
    print_func: Optional[Callable[..., None]] = None,
) -> int:
    """Walk user interactively through author-intent accessibility barriers."""
    if input_func is None:
        input_func = input
    if print_func is None:
        print_func = print

    in_p = Path(in_path)
    out_p = Path(out_path) if out_path else in_p.parent / f"{in_p.stem}-triaged.pptx"

    prs = Presentation(str(in_p))
    res = audit_file(in_p)
    findings = res.get("findings", [])

    items_triaged = 0
    print_func(f"\n=== pptx-a11y Interactive Accessibility Triage: {in_p.name} ===\n")

    for f in findings:
        rule_id = f["rule_id"]
        location = f["location"]
        desc = f["description"]

        if rule_id in ("image-alt-missing", "chart-missing-alt"):
            print_func(f"\n[Visual Asset Lacks Alternative Text]")
            print_func(f"Location: {location}")
            print_func(f"Issue:    {desc}")
            ans = input_func("Enter alt text description (or 'd' for decorative, 's' to skip): ").strip()
            if ans.lower() == "s" or not ans:
                continue
            elif ans.lower() == "d":
                _mark_shape_decorative(prs, location)
                items_triaged += 1
                print_func("-> Marked shape as decorative.")
            else:
                _set_shape_alt_text(prs, location, ans)
                items_triaged += 1
                print_func(f"-> Set alt text: '{ans}'")

        elif rule_id == "slide-title-missing":
            print_func(f"\n[Slide Missing Title]")
            print_func(f"Location: {location}")
            ans = input_func("Enter a title for this slide (or 's' to skip): ").strip()
            if ans.lower() != "s" and ans:
                _set_slide_title(prs, location, ans)
                items_triaged += 1
                print_func(f"-> Added slide title: '{ans}'")

        elif rule_id == "slide-title-duplicate":
            print_func(f"\n[Duplicate Slide Title]")
            print_func(f"Location: {location}")
            print_func(f"Issue:    {desc}")
            ans = input_func("Enter a distinct title (or 's' to skip): ").strip()
            if ans.lower() != "s" and ans:
                _set_slide_title(prs, location, ans)
                items_triaged += 1
                print_func(f"-> Updated slide title: '{ans}'")

    prs.save(str(out_p))
    print_func(f"\nTriage complete! {items_triaged} items updated. Saved to: {out_p}\n")
    return items_triaged
