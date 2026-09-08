"""Tests for deterministic presentation remediation."""
from pathlib import Path
from pptx_a11y.audit import audit_file
from pptx_a11y.remediate import remediate_presentation

FIXTURES = Path(__file__).parent / "fixtures"


def test_remediation_reduces_blocking_violations(tmp_path: Path):
    in_file = FIXTURES / "violations.pptx"
    out_file = tmp_path / "violations_fixed.pptx"

    fixes = remediate_presentation(in_file, out_file)
    assert fixes["title_added"] > 0
    assert fixes["table_headers_set"] > 0

    # Audit the remediated deck
    res_before = audit_file(in_file)
    res_after = audit_file(out_file)

    assert res_after["summary"]["blocking"] < res_before["summary"]["blocking"]
    rule_ids_after = {f["rule_id"] for f in res_after["findings"]}
    assert "title-missing" not in rule_ids_after
    assert "table-header-missing" not in rule_ids_after


def test_original_file_immutability_and_sha256(tmp_path: Path):
    import hashlib
    import pytest

    in_file = FIXTURES / "violations.pptx"
    out_file = tmp_path / "violations_fixed.pptx"

    sha_before = hashlib.sha256(in_file.read_bytes()).hexdigest()
    fixes = remediate_presentation(in_file, out_file)
    sha_after = hashlib.sha256(in_file.read_bytes()).hexdigest()

    assert sha_before == sha_after
    assert fixes["original_sha256"] == sha_before
    assert fixes["original_file_immutable"] is True

    # In-place remediation must be rejected to prevent corrupting or altering the original file
    with pytest.raises(ValueError, match="Cannot remediate in place"):
        remediate_presentation(in_file, in_file)


def test_password_and_modify_verifier_stripped(tmp_path: Path):
    test_deck = Path(__file__).parents[1] / "examples" / "No Knead Bread-test.pptx"
    if not test_deck.exists():
        return

    out_file = tmp_path / "bread_fixed.pptx"
    fixes = remediate_presentation(test_deck, out_file)

    assert fixes["passwords_stripped"] >= 1
    assert fixes["restricted_access_removed"] >= 1

    # Verify remediated deck has no modify verifier in XML
    res = audit_file(out_file)
    rule_ids = {f["rule_id"] for f in res["findings"]}
    assert "document-restricted-access" not in rule_ids
