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
