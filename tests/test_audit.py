"""Integration tests running audit_file on clean and violation fixtures."""
from pathlib import Path
from pptx_a11y.audit import audit_file

FIXTURES = Path(__file__).parent / "fixtures"


def test_clean_deck_audit():
    result = audit_file(FIXTURES / "clean.pptx")
    # Clean deck should pass blocking violations
    assert result["summary"]["blocking"] == 0
    assert result["summary"]["pass"] is True


def test_violations_deck_audit():
    result = audit_file(FIXTURES / "violations.pptx")
    assert result["summary"]["blocking"] > 0
    assert result["summary"]["pass"] is False
    rule_ids = {f["rule_id"] for f in result["findings"]}
    assert "title-missing" in rule_ids
    assert "image-alt-missing" in rule_ids
    assert "table-header-missing" in rule_ids
