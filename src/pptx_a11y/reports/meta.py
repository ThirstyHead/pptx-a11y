"""Canonical W3C WCAG 2.2 metadata, Understanding URLs, and POUR taxonomy."""

W3C_UNDERSTANDING_BASE = "https://www.w3.org/WAI/WCAG22/Understanding/"
W3C_QUICKREF = "https://www.w3.org/WAI/WCAG22/quickref/?currentsidebar=%23col_customize&levels=aaa"

# sc -> (Title, Level, POUR Principle, Canonical W3C URL)
SC_META = {
    "1.1.1": (
        "Non-text Content",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "non-text-content.html",
    ),
    "1.2.2": (
        "Captions (Prerecorded)",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "captions-prerecorded.html",
    ),
    "1.3.1": (
        "Info and Relationships",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "info-and-relationships.html",
    ),
    "1.3.2": (
        "Meaningful Sequence",
        "A",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "meaningful-sequence.html",
    ),
    "1.4.3": (
        "Contrast (Minimum)",
        "AA",
        "1 Perceivable",
        W3C_UNDERSTANDING_BASE + "contrast-minimum.html",
    ),
    "2.4.2": (
        "Page Titled",
        "A",
        "2 Operable",
        W3C_UNDERSTANDING_BASE + "page-titled.html",
    ),
    "2.4.4": (
        "Link Purpose (In Context)",
        "A",
        "2 Operable",
        W3C_UNDERSTANDING_BASE + "link-purpose-in-context.html",
    ),
    "3.1.1": (
        "Language of Page",
        "A",
        "3 Understandable",
        W3C_UNDERSTANDING_BASE + "language-of-page.html",
    ),
    "4.1.2": (
        "Name, Role, Value",
        "A",
        "4 Robust",
        W3C_UNDERSTANDING_BASE + "name-role-value.html",
    ),
}

PRINCIPLES = [
    ("1", "Perceivable"),
    ("2", "Operable"),
    ("3", "Understandable"),
    ("4", "Robust"),
]

POUR_INTROS = {
    "1": (
        "Perceivable content ensures that information and slide components can be received by everyone's "
        "senses. Visuals have text descriptions, colors provide high contrast, and structural information "
        "is explicitly coded so assistive technologies can read it."
    ),
    "2": (
        "Operable presentations let every reader navigate with ease: clear slide titles serve as bookmarks, "
        "meaningful reading order allows predictable keyboard and screen reader exploration, and hyperlinks "
        "clearly state where they lead."
    ),
    "3": (
        "Understandable material is clear and predictable: the presentation declares its natural language so "
        "speech synthesizers pronounce terms correctly, and layout choices avoid unexpected behavior."
    ),
    "4": (
        "Robust presentations use standard OpenXML slide placeholders and semantic shapes so that the content "
        "can be reliably interpreted across operating systems, mobile devices, and screen readers."
    ),
}
