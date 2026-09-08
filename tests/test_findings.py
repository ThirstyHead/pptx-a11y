"""Unit tests for findings data model and summarizer."""
from pptx_a11y.findings import Finding, findings_sorted, summarize


def test_findings_sorting_and_summary():
    f1 = Finding("img-alt", "1.1.1", "minor", "Slide 1", "No alt", "xml", True, "Add alt")
    f2 = Finding("contrast", "1.4.3", "critical", "Slide 2", "Low contrast", "xml", False, "Fix color")
    f3 = Finding("title", "2.4.2", "serious", "Slide 1", "No title", "xml", True, "Add title")

    sorted_f = findings_sorted([f1, f2, f3])
    assert [f.severity for f in sorted_f] == ["critical", "serious", "minor"]

    summary = summarize(sorted_f)
    assert summary["total"] == 3
    assert summary["blocking"] == 2
    assert summary["pass"] is False
    assert summary["by_severity"]["critical"] == 1
    assert summary["by_severity"]["minor"] == 1
