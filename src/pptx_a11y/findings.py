"""Finding data structures, sorting, and summary metrics."""
from dataclasses import asdict, dataclass
from typing import Any, Dict, List, Literal

Severity = Literal["critical", "serious", "moderate", "minor"]
SEVERITY_ORDER = {"critical": 0, "serious": 1, "moderate": 2, "minor": 3}


@dataclass
class Finding:
    rule_id: str
    sc: str  # WCAG Success Criterion (e.g. "1.1.1", "2.4.2")
    severity: Severity
    location: str  # e.g., "Slide 1, Shape 'Title 1'"
    description: str  # Clear, friendly description of the accessibility barrier
    evidence: str  # OpenXML element path or technical attribute detail
    fixable: bool  # Can this be deterministically remediated?
    fix: str  # Concrete guidance or description of remediation performed

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def findings_sorted(findings: List[Finding]) -> List[Finding]:
    """Sort findings by severity (critical first), then SC, then location."""
    return sorted(
        findings,
        key=lambda f: (
            SEVERITY_ORDER.get(f.severity, 99),
            f.sc,
            f.location,
            f.rule_id,
        ),
    )


def summarize(findings: List[Finding]) -> Dict[str, Any]:
    """Calculate summary statistics for a list of findings."""
    by_sev = {"critical": 0, "serious": 0, "moderate": 0, "minor": 0}
    for f in findings:
        if f.severity in by_sev:
            by_sev[f.severity] += 1
    total = len(findings)
    blocking = by_sev["critical"] + by_sev["serious"]
    return {
        "total": total,
        "by_severity": by_sev,
        "blocking": blocking,
        "pass": blocking == 0,
    }
