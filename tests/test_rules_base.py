"""Tests for rule registration and base execution."""
from pptx import Presentation
from pptx_a11y.rules import Rule, register_rule, RULES, AuditContext
from pptx_a11y.findings import Finding


def test_rule_registration():
    @register_rule
    class DummyRule(Rule):
        rule_id = "dummy-rule"
        sc = "1.1.1"
        severity = "minor"
        title = "Dummy Test Rule"

        def check(self, prs, ctx):
            return [Finding(self.rule_id, self.sc, self.severity, "Slide 1", "Dummy", "xml", False, "None")]

    assert any(r.rule_id == "dummy-rule" for r in RULES)
    prs = Presentation()
    ctx = AuditContext(source_name="test.pptx")
    results = [r for r in RULES if r.rule_id == "dummy-rule"][0].check(prs, ctx)
    assert len(results) == 1
    assert results[0].rule_id == "dummy-rule"
