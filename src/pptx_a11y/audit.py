"""Audit engine: execute registered rules against .pptx and compile findings."""
import json
import time
from pathlib import Path
from typing import Any, Dict, List, Optional
from pptx import Presentation
from .findings import Finding, findings_sorted, summarize
from .rules import RULES, AuditContext


def is_encrypted_package(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            header = f.read(1024)
            if b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" in header and b"EncryptedPackage" in header:
                return True
    except Exception:
        pass
    return False


def audit_file(path: str | Path, ctx: Optional[AuditContext] = None) -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Presentation not found: {p}")
    if ctx is None:
        ctx = AuditContext(source_name=p.name)

    if is_encrypted_package(p):
        finding = Finding(
            rule_id="document-restricted-access",
            sc="4.1.2",
            severity="critical",
            location="Package Security",
            description="The presentation file is password encrypted or protected with Information Rights Management (IRM).",
            evidence="EncryptedPackage OLE stream detected; contents inaccessible to assistive technology",
            fixable=False,
            fix="Remove encryption and password protection in PowerPoint settings.",
            why_unfixable="Cryptographic permissions and DRM protection cannot be removed without owner credentials.",
            manual_steps=[
                "Open the presentation in PowerPoint using your credentials.",
                "Go to 'File' -> 'Info' -> 'Protect Presentation'.",
                "Clear 'Encrypt with Password' or 'Restrict Access'.",
                "Save the presentation file.",
            ],
        )
        sorted_f = [finding]
        return {
            "file": p.name,
            "audited_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "tool": "pptx-a11y/0.2.0",
            "findings": [finding.to_dict()],
            "summary": summarize(sorted_f),
        }

    prs = Presentation(str(p))

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
