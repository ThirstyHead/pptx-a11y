"""Validation tests for bundled example presentations."""
from pathlib import Path
from pptx_a11y.audit import audit_file

EXAMPLES_DIR = Path(__file__).parent.parent / "examples"


def test_clean_example_presentation():
    good_deck = EXAMPLES_DIR / "No Knead Bread.pptx"
    assert good_deck.exists(), "No Knead Bread.pptx must exist in examples/"

    res = audit_file(good_deck)
    assert res["summary"]["total"] == 0, f"Expected 0 findings in clean example, got: {res['findings']}"
    assert res["summary"]["pass"] is True


def test_barrier_test_example_presentation():
    bad_deck = EXAMPLES_DIR / "No Knead Bread-test.pptx"
    assert bad_deck.exists(), "No Knead Bread-test.pptx must exist in examples/"

    res = audit_file(bad_deck)
    assert res["summary"]["total"] > 0
    assert res["summary"]["pass"] is False

    rule_ids = {f["rule_id"] for f in res["findings"]}
    expected_rules = {
        "title-missing",
        "slide-title-missing",
        "slide-title-duplicate",
        "section-name-default",
        "section-name-duplicate",
        "link-text-vague",
        "image-alt-missing",
        "media-subtitles-missing",
        "table-header-missing",
        "table-merged-cells",
        "chart-missing-alt",
        "language-missing",
        "semantic-placeholders-missing",
        "reading-order-inverted",
        "document-restricted-access",
    }
    missing = expected_rules - rule_ids
    assert not missing, f"Test deck is missing expected barrier rules: {missing}"
