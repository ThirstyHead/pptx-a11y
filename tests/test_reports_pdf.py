"""Tests for accessible PDF report generation from HTML."""
from pathlib import Path
import pikepdf
import pymupdf
from pptx_a11y.reports.html import render_html
from pptx_a11y.reports.pdf import render_pdf


def test_pdf_report_creation_and_metadata(tmp_path: Path):
    md = "# Accessible PowerPoint Audit\n\n## Executive Summary\n\nPresentation is compliant."
    html_doc = render_html(md, theme="light")
    pdf_path = tmp_path / "test_report.pdf"

    result_path = render_pdf(html_doc, out_path=pdf_path, title="Accessible PowerPoint Audit")
    assert result_path.exists()

    # Verify extractable text via pymupdf
    with pymupdf.open(result_path) as doc:
        text = "".join(str(page.get_text()) for page in doc)
        assert "Accessible PowerPoint Audit" in text
        assert "Executive Summary" in text

    # Verify accessibility metadata via pikepdf
    with pikepdf.open(result_path) as pdf:
        assert str(pdf.Root.Lang) == "en"
        assert pdf.Root.MarkInfo.Marked is True
        with pdf.open_metadata() as meta:
            assert meta["dc:title"] == "Accessible PowerPoint Audit"
