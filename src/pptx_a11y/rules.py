"""Base classes and registry for PresentationML WCAG 2.2 audit rules."""
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import List, Type
from pptx.presentation import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from .findings import Finding, Severity


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
