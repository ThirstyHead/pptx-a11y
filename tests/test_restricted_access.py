"""Tests for document restricted access rule (Section 508 / WCAG 4.1.2)."""
import zipfile
from pathlib import Path
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx_a11y.audit import audit_file
from pptx_a11y.rules import RestrictedAccessRule, AuditContext


def test_modify_verifier_detected_as_restricted_access():
    prs = Presentation()
    # inject modifyVerifier element into prs._element
    mv = parse_xml('<p:modifyVerifier xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" algorithmName="SHA-512" hashValue="abc"/>')
    prs._element.append(mv)

    rule = RestrictedAccessRule()
    findings = rule.check(prs, AuditContext("restricted.pptx"))
    assert len(findings) == 1
    f = findings[0]
    assert f.rule_id == "document-restricted-access"
    assert f.sc == "4.1.2"
    assert f.severity == "critical"
    assert f.why_unfixable is not None
    assert len(f.manual_steps) >= 3


def test_encrypted_package_audit_handled_gracefully(tmp_path: Path):
    enc_file = tmp_path / "encrypted.pptx"
    # Write OLE header with EncryptedPackage stream simulation
    enc_file.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"EncryptedPackage" * 20)

    res = audit_file(enc_file)
    assert res["summary"]["pass"] is False
    assert res["summary"]["blocking"] >= 1
    assert any(f["rule_id"] == "document-restricted-access" for f in res["findings"])
