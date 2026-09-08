"""Tests for Markdown report engine adhering to WCAG 2.2 & Social Model."""
import pytest
from pptx_a11y.reports.md import render_md
from pptx_a11y.reports.tone import assert_social_model_language


def test_social_model_language_guard():
    with pytest.raises(ValueError, match="Prohibited language detected"):
        assert_social_model_language("Screen reader users suffer from missing titles.")

    # Compliant social model phrasing
    assert_social_model_language(
        "The presentation lacks slide titles, preventing screen reader users from navigating by landmark."
    )


def test_markdown_pour_structure():
    sample_audit = {
        "file": "deck.pptx",
        "audited_at": "2026-09-08T12:00:00Z",
        "tool": "pptx-a11y/0.1.0",
        "summary": {
            "total": 1,
            "blocking": 1,
            "pass": False,
            "by_severity": {"critical": 1, "serious": 0, "moderate": 0, "minor": 0},
        },
        "findings": [
            {
                "rule_id": "image-alt-missing",
                "sc": "1.1.1",
                "severity": "critical",
                "location": "Slide 1",
                "description": "Image has no alt text.",
                "evidence": "xml",
                "fixable": False,
                "fix": "Add alt text.",
            }
        ],
    }
    md = render_md(sample_audit)
    assert "## 1. Perceivable" in md
    assert "## 2. Operable" in md
    assert "## 3. Understandable" in md
    assert "## 4. Robust" in md
    assert "https://www.w3.org/WAI/WCAG22/Understanding/non-text-content.html" in md
    assert "Who Benefits:" in md
    assert "Barrier Detected:" in md


def test_markdown_remediation_progress():
    before = {
        "file": "deck.pptx",
        "summary": {"total": 4, "blocking": 2, "pass": False},
        "findings": [],
    }
    after = {
        "file": "deck.pptx",
        "summary": {"total": 2, "blocking": 0, "pass": True},
        "findings": [],
    }
    md = render_md(before, after_result=after)
    assert "Remediation Progress" in md
    assert "100.0% improvement" in md
    assert "PASS" in md
