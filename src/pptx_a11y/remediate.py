"""Deterministic remediation engine for PowerPoint presentations."""
from pathlib import Path
from typing import Dict
from pptx import Presentation


def remediate_presentation(
    in_path: str | Path,
    out_path: str | Path,
    default_lang: str = "en-US",
) -> Dict[str, int]:
    prs = Presentation(str(in_path))
    fixes = {
        "title_added": 0,
        "table_headers_set": 0,
        "language_tagged": 0,
        "sections_renamed": 0,
    }

    # 1. Remediate core title if missing
    if not (prs.core_properties.title or "").strip():
        inferred_title = Path(in_path).stem.replace("-", " ").replace("_", " ").title()
        try:
            if prs.slides and prs.slides[0].shapes.title and prs.slides[0].shapes.title.text.strip():
                inferred_title = prs.slides[0].shapes.title.text.strip()
        except Exception:
            pass
        prs.core_properties.title = inferred_title
        fixes["title_added"] += 1

    # 2. Iterate slides for table headers and language tags
    for slide in prs.slides:
        for shape in slide.shapes:
            # Fix table headers
            if shape.has_table:
                tblPr = shape._element.xpath(".//a:tblPr")
                if tblPr and tblPr[0].get("firstRow") not in ("1", "true"):
                    tblPr[0].set("firstRow", "1")
                    fixes["table_headers_set"] += 1

            # Fix language tags on text runs
            tf = getattr(shape, "text_frame", None)
            if tf:
                for p in tf.paragraphs:
                    for run in p.runs:
                        rPr = run._r.xpath(".//a:rPr")
                        if rPr and not rPr[0].get("lang"):
                            rPr[0].set("lang", default_lang)
                            fixes["language_tagged"] += 1
                        elif not rPr:
                            new_rPr = run._r.get_or_add_rPr()
                            new_rPr.set("lang", default_lang)
                            fixes["language_tagged"] += 1

    # 3. Remediate section names (default and duplicates)
    from .rules import _get_sections, DEFAULT_SECTION_PATTERNS
    seen_sec = {}
    for s_idx, name, s_elem in _get_sections(prs):
        clean_name = name
        if name.lower() in DEFAULT_SECTION_PATTERNS:
            clean_name = f"Topic Section {s_idx}"
            s_elem.set("name", clean_name)
            fixes["sections_renamed"] += 1
        lower_clean = clean_name.lower()
        if lower_clean in seen_sec:
            seen_sec[lower_clean] += 1
            disambiguated = f"{clean_name} (Part {seen_sec[lower_clean]})"
            s_elem.set("name", disambiguated)
            fixes["sections_renamed"] += 1
        else:
            seen_sec[lower_clean] = 1

    prs.save(str(out_path))
    return fixes
