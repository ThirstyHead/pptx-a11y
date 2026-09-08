"""Tests for HTML5 report generation, SMACSS themes, and WCAG contrast gates."""
import re
from pptx_a11y.contrast import calculate_contrast_ratio
from pptx_a11y.reports.html import render_html
from pptx_a11y.reports.theme import available_themes, theme_css


def _extract_color_from_css(css: str, token: str) -> str:
    m = re.search(rf"{token}\s*:\s*#([0-9a-fA-F]{{6}})", css)
    if not m:
        raise ValueError(f"Token {token} not found in css")
    return m.group(1)


def test_bundled_themes_contrast_gate():
    for manifest in available_themes():
        name = manifest["name"]
        if name == "print":
            continue  # Monochromatic print uses text decoration cues
        css = theme_css(name)
        bg = _extract_color_from_css(css, "--bg")
        fg = _extract_color_from_css(css, "--fg")
        ratio = calculate_contrast_ratio(fg, bg)
        assert ratio >= 4.5, f"Theme '{name}' fg contrast {ratio}:1 fails 4.5:1 gate"

        # Check severity tokens vs bg
        for sev in ("--sev-critical", "--sev-serious", "--sev-moderate"):
            sev_color = _extract_color_from_css(css, sev)
            s_ratio = calculate_contrast_ratio(sev_color, bg)
            assert s_ratio >= 4.5, f"Theme '{name}' {sev} contrast {s_ratio}:1 fails 4.5:1 gate"


def test_html_accessibility_landmarks():
    md = "# Sample Deck\n\n## Executive Summary\n\nDetails.\n\n## 1. Perceivable\n\nIntro."
    html_doc = render_html(md, theme="light")

    assert '<html lang="en">' in html_doc
    assert '<a class="skip" href="#main">Skip to main content</a>' in html_doc
    assert '<nav class="toc" aria-label="Table of contents">' in html_doc
    assert '<main id="main" class="report"' in html_doc
    assert '<div class="summary-banner">' in html_doc
