"""Social model tone phrasing, who-benefits mappings, and language guards."""

# Approved "Who Benefits" statements grounded in the Social Model of Disability
WHO_MAP = {
    "1.1.1": (
        "People who are blind, have low vision, or process information through audio now receive the complete "
        "message through spoken screen reader narration or braille displays."
    ),
    "1.2.2": (
        "People who are deaf or hard of hearing can now access spoken dialogue and auditory information through "
        "synchronized closed captions."
    ),
    "1.3.1": (
        "People navigating by screen reader, or individuals who benefit from clean, predictable structure, "
        "can follow headers, table data relationships, and slide layouts without disorientation."
    ),
    "1.3.2": (
        "People who use screen readers or keyboard-only navigation experience content in the logical order "
        "intended by the author, rather than by accidental visual layering."
    ),
    "1.4.3": (
        "People with low vision, color vision differences, or anyone reading under bright ambient sunlight "
        "can comfortably distinguish the text from the slide background."
    ),
    "2.4.2": (
        "People who navigate presentations using slide titles as landmarks—such as screen reader users and "
        "audiences scanning outlines—can immediately locate and differentiate topics."
    ),
    "2.4.4": (
        "People listening to links in a list or scanning quickly can instantly understand where each destination "
        "leads without needing to decipher ambiguous phrases."
    ),
    "3.1.1": (
        "People relying on text-to-speech tools or automated translation software will hear the presentation read "
        "with accurate phonetics, accentuation, and grammatical inflection."
    ),
    "4.1.2": (
        "People using third-party assistive tools receive consistent, reliable names, roles, and values for every "
        "interactive and structural element on the slide."
    ),
}

# Approved barrier descriptions emphasizing document deficiencies
RULE_BARRIER_EXPLANATIONS = {
    "title-missing": (
        "The presentation file lacks an embedded title in its metadata properties, meaning file managers and screen "
        "readers identify it only by its raw filename."
    ),
    "slide-title-missing": (
        "This slide is missing a structured Title placeholder, making it invisible in navigation outlines and "
        "unlabeled for assistive technology users."
    ),
    "slide-title-duplicate": (
        "This slide shares an identical title with an earlier slide, creating confusion when navigating by headings."
    ),
    "image-alt-missing": (
        "This visual asset has neither alternative text nor a decorative flag, meaning screen reader users are left "
        "unaware of its presence or purpose."
    ),
    "table-header-missing": (
        "The table does not mark its top row as a header row, so screen readers cannot announce column headers as "
        "users move between data cells."
    ),
    "color-contrast": (
        "The text contrast ratio is below the required 4.5:1 threshold, creating a visual barrier under typical "
        "viewing conditions."
    ),
    "link-text-vague": (
        "The hyperlink uses generic anchor text (such as 'click here') rather than describing the destination."
    ),
    "language-missing": (
        "Text runs do not declare an explicit language tag, risking incorrect pronunciation by speech tools."
    ),
    "semantic-placeholders-missing": (
        "The slide relies on unparented text boxes rather than semantic layout placeholders, which degrades "
        "predictable screen reader reading order."
    ),
    "media-subtitles-missing": (
        "Audio or video content is presented without synchronized closed captions or text alternatives, "
        "preventing readers who are deaf or hard of hearing from accessing spoken dialogue."
    ),
    "table-merged-cells": (
        "The table contains merged or split cells, disrupting the two-dimensional grid and causing screen "
        "readers to lose track of column and row relationships."
    ),
    "reading-order-inverted": (
        "Slide shapes are visually positioned out of sequence compared to their underlying shape tree order, "
        "meaning screen readers announce content out of order."
    ),
    "section-name-default": (
        "The presentation group uses a generic default section name, offering no descriptive context for navigation."
    ),
    "section-name-duplicate": (
        "Multiple presentation sections share the same name, creating ambiguity in the slide outline."
    ),
    "document-restricted-access": (
        "The presentation has password encryption or Information Rights Management (IRM) enabled, blocking "
        "assistive technologies from accessing its structure."
    ),
    "chart-missing-alt": (
        "The chart or embedded object does not have alternative text or an accompanying data table, leaving its "
        "information inaccessible to screen reader users."
    ),
}

# Prohibited medical-model and condescending phrases
BANNED_PHRASES = [
    "suffer from",
    "suffers from",
    "handicapped",
    "normal users",
    "regular users",
    "afflicted with",
    "victim of",
    "confined to",
    "inaccessible to disabled",
    "broken presentation",
    "stupid mistake",
]


def assert_social_model_language(text: str) -> None:
    """Validate that text does not contain prohibited medical model or condescending terminology."""
    lower_text = text.lower()
    for phrase in BANNED_PHRASES:
        if phrase in lower_text:
            raise ValueError(f"Prohibited language detected: '{phrase}'. Adhere strictly to the Social Model.")
