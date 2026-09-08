"""Base classes and registry for PresentationML WCAG 2.2 audit rules."""
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import List, Type
from pptx.presentation import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from .findings import Finding, Severity
from .reading_order import check_slide_reading_order


@dataclass
class AuditContext:
    source_name: str
    default_lang: str = "en-US"


class Rule(ABC):
    rule_id: str
    sc: str
    severity: Severity
    title: str

    @abstractmethod
    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        """Evaluate rule against presentation and return findings."""
        pass


RULES: List[Rule] = []


def register_rule(rule_cls: Type[Rule]) -> Type[Rule]:
    RULES.append(rule_cls())
    return rule_cls


# ---------------------------------------------------------------------------
# Principle 2: Operable
# ---------------------------------------------------------------------------

DEFAULT_SECTION_PATTERNS = {
    "default section",
    "untitled section",
    "new section",
    "section",
    "section 1",
    "section 2",
    "section 3",
    "section 4",
    "section 5",
}


def _get_sections(prs: Presentation):
    sec_elems = prs._element.xpath(".//p:sectionLst/p:section | .//*[local-name()='section']")
    results = []
    for idx, elem in enumerate(sec_elems, start=1):
        name = elem.get("name", "").strip()
        results.append((idx, name, elem))
    return results


@register_rule
class DefaultSectionNameRule(Rule):
    rule_id = "section-name-default"
    sc = "2.4.2"
    severity = "moderate"
    title = "Default Section Name Used"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for idx, name, elem in _get_sections(prs):
            if name.lower() in DEFAULT_SECTION_PATTERNS:
                findings.append(Finding(
                    rule_id=self.rule_id,
                    sc=self.sc,
                    severity=self.severity,
                    location=f"Section {idx} ('{name}')",
                    description=f"Section {idx} uses a default, non-descriptive name '{name}'.",
                    evidence=f"Section element in presentation.xml has name='{name}'",
                    fixable=True,
                    fix="Rename the section to describe the topic covered by its slides.",
                    why_unfixable="Automated tools can only infer section names from slide titles; manual naming preferred.",
                    manual_steps=[
                        f"In PowerPoint's thumbnail pane, locate Section {idx} ('{name}').",
                        "Right-click the section header and select 'Rename Section'.",
                        "Provide a clear, descriptive title representing the group of slides.",
                    ],
                ))
        return findings


@register_rule
class DuplicateSectionNameRule(Rule):
    rule_id = "section-name-duplicate"
    sc = "2.4.2"
    severity = "moderate"
    title = "Duplicate Section Name"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        seen = {}
        for idx, name, elem in _get_sections(prs):
            if not name:
                continue
            lower_name = name.lower()
            if lower_name in seen:
                prev_idx = seen[lower_name]
                findings.append(Finding(
                    rule_id=self.rule_id,
                    sc=self.sc,
                    severity=self.severity,
                    location=f"Section {idx} ('{name}')",
                    description=f"Section {idx} repeats the section name '{name}' previously used for Section {prev_idx}.",
                    evidence=f"Duplicate section name '{name}' across sections {prev_idx} and {idx}",
                    fixable=True,
                    fix=f"Differentiate section names (e.g., '{name} (Part 2)' or '{name} - Continued').",
                    why_unfixable="Disambiguating section titles requires human editorial judgment.",
                    manual_steps=[
                        f"In PowerPoint's thumbnail pane, right-click the second '{name}' section header.",
                        "Select 'Rename Section' and differentiate it (e.g. '{name} (Part 2)').",
                    ],
                ))
            else:
                seen[lower_name] = idx
        return findings


@register_rule
class PresentationTitleRule(Rule):
    rule_id = "title-missing"
    sc = "2.4.2"
    severity = "serious"
    title = "Presentation Metadata Title Missing"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        title = (prs.core_properties.title or "").strip()
        if not title:
            findings.append(Finding(
                rule_id=self.rule_id,
                sc=self.sc,
                severity=self.severity,
                location="Presentation Metadata (core.xml)",
                description="The presentation does not declare a title in its document properties.",
                evidence="<dc:title> is empty or missing in docProps/core.xml",
                fixable=True,
                fix="Set a clear, descriptive title in the presentation properties or from the first slide title.",
            ))
        return findings


@register_rule
class SlideTitleRule(Rule):
    rule_id = "slide-title-missing"
    sc = "2.4.2"
    severity = "serious"
    title = "Slide Title Placeholder Missing"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for idx, slide in enumerate(prs.slides, start=1):
            has_title = False
            try:
                if slide.shapes.title and slide.shapes.title.text.strip():
                    has_title = True
            except Exception:
                has_title = False

            if not has_title:
                findings.append(Finding(
                    rule_id=self.rule_id,
                    sc=self.sc,
                    severity=self.severity,
                    location=f"Slide {idx}",
                    description=f"Slide {idx} does not have an explicit, structured title placeholder.",
                    evidence=f"Slide {idx} shape tree has no populated p:ph type='title' or 'ctrTitle'",
                    fixable=True,
                    fix="Use a slide layout containing an explicit Title placeholder and provide descriptive text.",
                ))
        return findings


@register_rule
class DuplicateSlideTitleRule(Rule):
    rule_id = "slide-title-duplicate"
    sc = "2.4.2"
    severity = "moderate"
    title = "Duplicate Slide Title"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        seen_titles: dict[str, int] = {}
        for idx, slide in enumerate(prs.slides, start=1):
            title_text = ""
            try:
                if slide.shapes.title:
                    title_text = slide.shapes.title.text.strip()
            except Exception:
                pass

            if title_text:
                if title_text in seen_titles:
                    prev_idx = seen_titles[title_text]
                    findings.append(Finding(
                        rule_id=self.rule_id,
                        sc=self.sc,
                        severity=self.severity,
                        location=f"Slide {idx}",
                        description=f"Slide {idx} repeats the title '{title_text}' previously used on Slide {prev_idx}.",
                        evidence=f"Duplicate title string '{title_text}' across slides {prev_idx} and {idx}",
                        fixable=False,
                        fix=f"Differentiate slide titles (e.g., '{title_text} (Part 2)' or '{title_text} - Continued').",
                    ))
                else:
                    seen_titles[title_text] = idx
        return findings


VAGUE_LINK_TEXTS = {"click here", "here", "read more", "more", "learn more", "link"}


@register_rule
class LinkTextRule(Rule):
    rule_id = "link-text-vague"
    sc = "2.4.4"
    severity = "serious"
    title = "Ambiguous Link Text"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                tf = getattr(shape, "text_frame", None)
                if not tf:
                    continue
                for p in tf.paragraphs:
                    for run in p.runs:
                        if run.hyperlink and run.hyperlink.address:
                            text = run.text.strip().lower()
                            if text in VAGUE_LINK_TEXTS:
                                findings.append(Finding(
                                    rule_id=self.rule_id,
                                    sc=self.sc,
                                    severity=self.severity,
                                    location=f"Slide {s_idx}, Text '{run.text.strip()}'",
                                    description=f"Hyperlink text '{run.text.strip()}' on Slide {s_idx} does not describe its destination.",
                                    evidence=f"Run has address '{run.hyperlink.address}' with generic anchor text '{run.text.strip()}'",
                                    fixable=False,
                                    fix="Replace vague phrases like 'click here' with meaningful link text describing the target.",
                                ))
        return findings


# ---------------------------------------------------------------------------
# Principle 1: Perceivable
# ---------------------------------------------------------------------------

@register_rule
class VisualAltTextRule(Rule):
    rule_id = "image-alt-missing"
    sc = "1.1.1"
    severity = "critical"
    title = "Image / Graphic Missing Text Alternative"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                is_visual = shape.shape_type in (
                    MSO_SHAPE_TYPE.PICTURE,
                    MSO_SHAPE_TYPE.MEDIA,
                    MSO_SHAPE_TYPE.AUTO_SHAPE,
                    MSO_SHAPE_TYPE.FREEFORM,
                )
                if not is_visual:
                    continue

                cNvPr = shape._element.xpath(".//p:cNvPr")
                if not cNvPr:
                    continue
                node = cNvPr[0]
                descr = node.get("descr", "").strip()
                title = node.get("title", "").strip()

                is_decorative = any(
                    elem.tag.endswith("decorative") and elem.get("val") == "1"
                    for elem in node.iter()
                )

                if not descr and not title and not is_decorative:
                    if shape.is_placeholder:
                        continue
                    findings.append(Finding(
                        rule_id=self.rule_id,
                        sc=self.sc,
                        severity=self.severity,
                        location=f"Slide {s_idx}, Shape '{shape.name}' (ID {shape.shape_id})",
                        description=f"Visual asset '{shape.name}' on Slide {s_idx} has no alternative text or decorative flag.",
                        evidence=f"<p:cNvPr id='{shape.shape_id}' name='{shape.name}'> has no descr attribute and no adec:decorative tag",
                        fixable=False,
                        fix="Provide concise, meaningful alternative text describing the image's content or mark it as decorative.",
                        why_unfixable="Automated tools cannot guess visual content or context without human authorial intent.",
                        manual_steps=[
                            f"Select '{shape.name}' on Slide {s_idx}.",
                            "Right-click and select 'View Alt Text' (or navigate to Picture Format -> Alt Text).",
                            "Enter a concise 1-2 sentence description conveying the essential information.",
                            "If the asset is purely visual fluff, check 'Mark as decorative'.",
                        ],
                    ))
        return findings


@register_rule
class MediaSubtitlesRule(Rule):
    rule_id = "media-subtitles-missing"
    sc = "1.2.2"
    severity = "critical"
    title = "Audio / Video Missing Subtitles or Captions"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                is_media = False
                if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                    is_media = True
                else:
                    # Inspect oxml for embedded media references
                    media_tags = shape._element.xpath(
                        ".//a:videoFile | .//p:videoFile | .//a:quickTimeFile | .//p:quickTimeFile | .//p:media | .//a:audioFile"
                    )
                    if media_tags:
                        is_media = True

                if not is_media:
                    continue

                # Check if closed captions or timed text track is attached
                has_captions = False
                caption_elems = shape._element.xpath(".//p:custDataLst | .//a:extLst")
                for elem in caption_elems:
                    for sub in elem.iter():
                        if "vtt" in (sub.get("val", "") + sub.get("href", "")).lower():
                            has_captions = True
                            break
                    if has_captions:
                        break

                if not has_captions:
                    findings.append(Finding(
                        rule_id=self.rule_id,
                        sc=self.sc,
                        severity=self.severity,
                        location=f"Slide {s_idx}, Media '{shape.name}'",
                        description=f"Audio or video element '{shape.name}' on Slide {s_idx} does not have synchronized closed captions or subtitles.",
                        evidence=f"Shape '{shape.name}' contains media stream with 0 attached timed text/WebVTT caption tracks",
                        fixable=False,
                        fix="Attach a synchronized WebVTT (.vtt) caption track or provide a complete verbatim text transcript in speaker notes.",
                        why_unfixable="Software cannot reliably transcribe spoken dialogue or synchronize timestamps without human verification.",
                        manual_steps=[
                            f"Select media player '{shape.name}' on Slide {s_idx}.",
                            "On the PowerPoint ribbon, select the 'Playback' tab.",
                            "Click 'Insert Captions' and select your WebVTT (.vtt) subtitles file.",
                            "Alternatively, type or paste the complete spoken transcript into the slide's Speaker Notes.",
                        ],
                    ))
        return findings


@register_rule
class TableHeaderRule(Rule):
    rule_id = "table-header-missing"
    sc = "1.3.1"
    severity = "serious"
    title = "Table Header Row Not Declared"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                if shape.has_table:
                    tblPr = shape._element.xpath(".//a:tblPr")
                    is_header = False
                    if tblPr:
                        is_header = tblPr[0].get("firstRow") in ("1", "true")

                    if not is_header:
                        findings.append(Finding(
                            rule_id=self.rule_id,
                            sc=self.sc,
                            severity=self.severity,
                            location=f"Slide {s_idx}, Table '{shape.name}'",
                            description=f"Table on Slide {s_idx} does not have a designated header row.",
                            evidence="<a:tblPr> is missing firstRow='1' or firstRow is '0'",
                            fixable=True,
                            fix="Enable 'Header Row' in PowerPoint Table Design tools or set firstRow='1' in table properties.",
                        ))
        return findings


@register_rule
class TableMergedCellsRule(Rule):
    rule_id = "table-merged-cells"
    sc = "1.3.1"
    severity = "serious"
    title = "Table Contains Merged or Split Cells"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                if shape.has_table:
                    tbl = getattr(shape, "table", None)
                    if not tbl:
                        continue
                    has_merged = False
                    for row in tbl.rows:
                        for cell in row.cells:
                            tc = cell._tc
                            if (
                                int(tc.get("gridSpan", "1")) > 1
                                or int(tc.get("rowSpan", "1")) > 1
                                or tc.get("hMerge") in ("1", "true")
                                or tc.get("vMerge") in ("1", "true")
                            ):
                                has_merged = True
                                break
                        if has_merged:
                            break

                    if has_merged:
                        findings.append(Finding(
                            rule_id=self.rule_id,
                            sc=self.sc,
                            severity=self.severity,
                            location=f"Slide {s_idx}, Table '{shape.name}'",
                            description=f"Table '{shape.name}' on Slide {s_idx} contains merged or split cells.",
                            evidence=f"Table '{shape.name}' has cells with gridSpan > 1, rowSpan > 1, or hMerge/vMerge attributes",
                            fixable=False,
                            fix="Avoid merged cells in presentation tables; unmerge and split complex tables into simple uniform grids.",
                            why_unfixable="Automated unmerging could scramble tabular relationships and misalign data columns.",
                            manual_steps=[
                                f"Select table '{shape.name}' on Slide {s_idx}.",
                                "Go to 'Table Design' / 'Layout' on the PowerPoint ribbon.",
                                "Use 'Split Cells' to restore a regular grid where every data cell maps to exactly one column and row header.",
                                "If the table presents multiple distinct datasets, split it into two separate, simpler tables each with its own header row.",
                            ],
                        ))
        return findings


@register_rule
class ChartAltTextRule(Rule):
    rule_id = "chart-missing-alt"
    sc = "1.1.1"
    severity = "critical"
    title = "Chart or Embedded Object Missing Alt Text"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                is_chart_or_ole = False
                if getattr(shape, "has_chart", False):
                    is_chart_or_ole = True
                elif shape._element.xpath(".//*[local-name()='chart' or local-name()='oleObj']"):
                    is_chart_or_ole = True

                if not is_chart_or_ole:
                    continue

                cNvPr = shape._element.xpath(".//p:cNvPr")
                descr = ""
                title = ""
                if cNvPr:
                    descr = cNvPr[0].get("descr", "").strip()
                    title = cNvPr[0].get("title", "").strip()

                if not descr and not title:
                    findings.append(Finding(
                        rule_id=self.rule_id,
                        sc=self.sc,
                        severity=self.severity,
                        location=f"Slide {s_idx}, Chart '{shape.name}'",
                        description=f"Chart or embedded object '{shape.name}' on Slide {s_idx} has no alternative text.",
                        evidence=f"<p:cNvPr> for '{shape.name}' lacks both descr and title attributes",
                        fixable=False,
                        fix="Provide a concise text summary of the chart's data trend and include an accompanying data table.",
                        why_unfixable="Data visualization meaning requires human context or an accompanying data table.",
                        manual_steps=[
                            f"Select '{shape.name}' on Slide {s_idx}.",
                            "Right-click and select 'View Alt Text'.",
                            "Enter a 1-2 sentence description summarizing the main trends or data points.",
                            "Provide an accessible data table with explicit headers on the slide or in notes.",
                        ],
                    ))
        return findings


# ---------------------------------------------------------------------------
# Principle 3: Understandable
# ---------------------------------------------------------------------------

@register_rule
class LanguageRule(Rule):
    rule_id = "language-missing"
    sc = "3.1.1"
    severity = "serious"
    title = "Text Language Tag Missing"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            for shape in slide.shapes:
                tf = getattr(shape, "text_frame", None)
                if not tf:
                    continue
                for p in tf.paragraphs:
                    for run in p.runs:
                        if not run.text.strip():
                            continue
                        rPr = run._r.xpath(".//a:rPr")
                        lang = rPr[0].get("lang") if rPr else None
                        if not lang:
                            findings.append(Finding(
                                rule_id=self.rule_id,
                                sc=self.sc,
                                severity=self.severity,
                                location=f"Slide {s_idx}, Shape '{shape.name}'",
                                description=f"Text run on Slide {s_idx} does not declare a natural language tag.",
                                evidence="<a:rPr> element missing lang attribute",
                                fixable=True,
                                fix=f"Set language tag (e.g. lang='{ctx.default_lang}') on all text runs.",
                            ))
                            return findings  # Flag once per presentation to avoid noise
        return findings


# ---------------------------------------------------------------------------
# Principle 4: Robust / Meaningful Structure
# ---------------------------------------------------------------------------

@register_rule
class RestrictedAccessRule(Rule):
    rule_id = "document-restricted-access"
    sc = "4.1.2"
    severity = "critical"
    title = "Presentation Access Restricted or Encrypted"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        verifiers = prs._element.xpath(".//p:modifyVerifier | .//*[local-name()='modifyVerifier']")
        if verifiers:
            findings.append(Finding(
                rule_id=self.rule_id,
                sc=self.sc,
                severity=self.severity,
                location="Presentation Security (presentation.xml)",
                description="The presentation has modify verification or restricted access permissions enabled.",
                evidence="<p:modifyVerifier> found in presentation.xml",
                fixable=False,
                fix="Remove edit restrictions or DRM permissions so assistive technologies can read all content.",
                why_unfixable="Cryptographic permissions and DRM protection cannot be removed without owner credentials.",
                manual_steps=[
                    "Open the presentation in PowerPoint with authoring privileges.",
                    "Go to 'File' -> 'Info' -> 'Protect Presentation'.",
                    "Select 'Encrypt with Password' or 'Restrict Access' and clear all restrictions.",
                    "Save the presentation file.",
                ],
            ))
        return findings


@register_rule
class SemanticStructureRule(Rule):
    rule_id = "semantic-placeholders-missing"
    sc = "1.3.2"
    severity = "moderate"
    title = "Slide Lacks Structured Placeholders"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            has_placeholders = any(s.is_placeholder for s in slide.shapes)
            has_content = len(slide.shapes) > 0
            if has_content and not has_placeholders:
                findings.append(Finding(
                    rule_id=self.rule_id,
                    sc=self.sc,
                    severity=self.severity,
                    location=f"Slide {s_idx}",
                    description=f"Slide {s_idx} is built from arbitrary text boxes without structured slide placeholders.",
                    evidence=f"Slide {s_idx} has {len(slide.shapes)} shapes but 0 placeholders in p:spTree",
                    fixable=False,
                    fix="Rebuild the slide using standard PowerPoint slide layouts with title and content placeholders.",
                ))
        return findings


@register_rule
class ReadingOrderRule(Rule):
    rule_id = "reading-order-inverted"
    sc = "1.3.2"
    severity = "serious"
    title = "Slide Content Reading Order Inverted"

    def check(self, prs: Presentation, ctx: AuditContext) -> List[Finding]:
        findings = []
        for s_idx, slide in enumerate(prs.slides, start=1):
            if check_slide_reading_order(slide):
                findings.append(Finding(
                    rule_id=self.rule_id,
                    sc=self.sc,
                    severity=self.severity,
                    location=f"Slide {s_idx}",
                    description=f"Shapes on Slide {s_idx} appear visually out of order compared to their screen reader sequence.",
                    evidence=f"Slide {s_idx} shape tree coordinates show top-positioned elements placed after lower elements in p:spTree",
                    fixable=True,
                    fix="Reorder shapes in the PowerPoint Selection Pane / Reading Order pane to follow logical top-to-bottom sequence.",
                    why_unfixable="Automated spatial reordering may alter visual shape layering; review in Reading Order Pane.",
                    manual_steps=[
                        f"Navigate to Slide {s_idx} in PowerPoint.",
                        "On the ribbon, select 'Home' -> 'Arrange' -> 'Selection Pane' (or 'Review' -> 'Check Accessibility' -> 'Reading Order Pane').",
                        "Verify shapes are ordered top-to-bottom and left-to-right.",
                        "Drag shapes into the intended reading order.",
                    ],
                ))
        return findings
