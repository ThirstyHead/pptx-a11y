"""pptx-a11y CLI entry point."""
import argparse
import sys
from pathlib import Path
from .audit import audit_file, audit_result_to_json
from .remediate import remediate_presentation
from .reports.html import render_html
from .reports.md import render_md
from .reports.pdf import render_pdf
from .reports.theme import available_themes
from .triage import run_interactive_triage


def main():
    parser = argparse.ArgumentParser(
        prog="pptx-a11y",
        description="Audit and remediate PowerPoint presentations against WCAG 2.2 AA standards.",
    )
    parser.add_argument("file", help="Path to PowerPoint .pptx file")
    parser.add_argument(
        "--format",
        default="md",
        help="Report formats (comma-separated): md, html, pdf, json (default: md)",
    )
    parser.add_argument(
        "--theme",
        default="light",
        help=f"SMACSS theme for HTML/PDF reports: {[t['name'] for t in available_themes()]}",
    )
    parser.add_argument("--output-dir", default=".", help="Directory to save generated reports")
    parser.add_argument("--fix", action="store_true", help="Perform deterministic remediation")
    parser.add_argument("--triage", action="store_true", help="Launch interactive human-in-the-loop triage session")
    parser.add_argument("--out-pptx", help="Output path for remediated .pptx file")

    args = parser.parse_args()
    input_path = Path(args.file)
    if not input_path.exists():
        print(f"Error: File '{input_path}' not found.", file=sys.stderr)
        sys.exit(1)

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = input_path.stem

    # Interactive triage mode
    if args.triage:
        triaged_pptx = Path(args.out_pptx) if args.out_pptx else out_dir / f"{stem}-triaged.pptx"
        run_interactive_triage(input_path, triaged_pptx)
        input_path = triaged_pptx
        stem = input_path.stem

    # 1. Audit original file
    audit_before = audit_file(input_path)
    audit_after = None

    # 2. Remediate if requested
    if args.fix:
        fixed_pptx = Path(args.out_pptx) if args.out_pptx else out_dir / f"{stem}-remediated.pptx"
        fixes = remediate_presentation(input_path, fixed_pptx)
        print(f"Remediation saved to: {fixed_pptx}")
        print(f"Fixes applied: {fixes}")
        audit_after = audit_file(fixed_pptx)

    # 3. Render canonical Markdown report (Source of Truth)
    formats = [f.strip().lower() for f in args.format.split(",")]
    md_text = render_md(audit_before, after_result=audit_after, source_path=str(input_path))

    if "md" in formats:
        md_file = out_dir / f"{stem}-a11y-report.md"
        md_file.write_text(md_text, encoding="utf-8")
        print(f"Markdown report: {md_file}")

    if "json" in formats:
        json_file = out_dir / f"{stem}-audit.json"
        json_file.write_text(audit_result_to_json(audit_before), encoding="utf-8")
        print(f"JSON audit: {json_file}")

    if "html" in formats:
        html_doc = render_html(md_text, theme=args.theme)
        html_file = out_dir / f"{stem}-a11y-report.html"
        html_file.write_text(html_doc, encoding="utf-8")
        print(f"Accessible HTML report: {html_file}")

    if "pdf" in formats:
        html_doc = render_html(md_text, theme=args.theme)
        pdf_file = out_dir / f"{stem}-a11y-report.pdf"
        render_pdf(html_doc, out_path=pdf_file)
        print(f"Accessible PDF report: {pdf_file}")

    # Exit code reflects compliance status (0 = compliant, 1 = action required)
    final_summary = (audit_after or audit_before)["summary"]
    sys.exit(0 if final_summary["pass"] else 1)


if __name__ == "__main__":
    main()
