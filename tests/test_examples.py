"""Validation tests for bundled example presentations."""
from pathlib import Path
from pptx_a11y.audit import audit_file
from pptx_a11y.remediate import remediate_presentation

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
        "section-name-default",
        "section-name-duplicate",
        "link-text-vague",
        "image-alt-missing",
        "table-header-missing",
        "language-missing",
        "reading-order-inverted",
    }
    missing = expected_rules - rule_ids
    assert not missing, f"Test deck is missing expected barrier rules: {missing}"


def test_auto_remediation_of_test_presentation(tmp_path: Path):
    bad_deck = EXAMPLES_DIR / "No Knead Bread-test.pptx"
    remediated_p = tmp_path / "No Knead Bread-remediated.pptx"

    fixes = remediate_presentation(bad_deck, remediated_p)
    assert fixes["title_added"] > 0
    assert fixes["table_headers_set"] > 0
    assert fixes["language_tagged"] > 0
    assert fixes["sections_renamed"] > 0
    assert fixes["reading_order_fixed"] > 0
    assert fixes["alt_text_added"] > 0
    assert fixes["links_disambiguated"] > 0

    # Remediated deck must pass cleanly with 0 findings, exactly like No Knead Bread.pptx
    res = audit_file(remediated_p)
    assert res["summary"]["total"] == 0, f"Expected 0 findings after remediation, got: {res['findings']}"
    assert res["summary"]["pass"] is True
