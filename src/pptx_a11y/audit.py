"""Audit engine: execute registered rules against .pptx and compile findings."""
import json
import time
from pathlib import Path
from typing import Any, Dict, List, Optional
from pptx import Presentation
from .findings import Finding, findings_sorted, summarize
from .rules import RULES, AuditContext


def audit_file(path: str | Path, ctx: Optional[AuditContext] = None) -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Presentation not found: {p}")
    prs = Presentation(str(p))
    if ctx is None:
        ctx = AuditContext(source_name=p.name)

    findings: List[Finding] = []
    for rule in RULES:
        try:
            findings.extend(rule.check(prs, ctx))
        except Exception as exc:
            findings.append(Finding(
                rule_id=f"{rule.rule_id}__error",
                sc=rule.sc,
                severity="moderate",
                location="internal",
                description=f"Rule {rule.rule_id} raised an internal error: {exc}",
                evidence=repr(exc),
                fixable=False,
                fix="Inspect slide oxml manually; treat as rule error.",
            ))

    sorted_f = findings_sorted(findings)
    return {
        "file": p.name,
        "audited_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "tool": "pptx-a11y/0.1.0",
        "findings": [f.to_dict() for f in sorted_f],
        "summary": summarize(sorted_f),
    }


def audit_result_to_json(result: Dict[str, Any]) -> str:
    return json.dumps(result, indent=2, sort_keys=True)
