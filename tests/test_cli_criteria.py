"""Test criteria configuration and what-if analysis via pptx-a11y CLI."""
from pathlib import Path
import pytest
from pptx_a11y.cli import main

FIXTURES = Path(__file__).parent / "fixtures"


def test_cli_init_criteria(tmp_path: Path, monkeypatch, capsys):
    checklist_path = tmp_path / "custom-criteria.txt"
    monkeypatch.setattr("sys.argv", ["pptx-a11y", "--init-criteria", str(checklist_path)])
    with pytest.raises(SystemExit) as exc:
        main()
    assert exc.value.code == 0
    assert checklist_path.exists()
    content = checklist_path.read_text(encoding="utf-8")
    assert "[x] 1.1.1 Non-text Content" in content
    assert "[x] 1.4.3 Contrast (Minimum)" in content
    assert "[x] 2.4.2 Page Titled" in content


def test_cli_criteria_what_if_exclusion(tmp_path: Path, monkeypatch):
    # Violations.pptx has violations including 1.4.3 Contrast (Minimum) and others.
    # We create a criteria file where 1.4.3 and all other violations are excluded, or test selective exclusion.
    checklist_path = tmp_path / "criteria.txt"
    # Write a checklist that excludes 1.4.3 Contrast (Minimum)
    checklist_path.write_text(
        "[ ] 1.4.3 Contrast (Minimum)\n[x] 1.1.1 Non-text Content\n",
        encoding="utf-8",
    )
    report_dir = tmp_path / "out"
    monkeypatch.setattr(
        "sys.argv",
        [
            "pptx-a11y",
            str(FIXTURES / "violations.pptx"),
            "--criteria",
            str(checklist_path),
            "--output-dir",
            str(report_dir),
            "--format",
            "md",
        ],
    )
    with pytest.raises(SystemExit):
        main()

    md_report = report_dir / "violations-a11y-report.md"
    assert md_report.exists()
    content = md_report.read_text(encoding="utf-8")
    assert "What-If Analysis Active" in content
    assert "[EXCLUDED FROM SUMMARY]" in content
