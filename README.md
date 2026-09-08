# pptx-a11y

Audit and remediate Microsoft PowerPoint `.pptx` presentations against **WCAG 2.2 AA** standards — standalone, deterministic, and headless.

`pptx-a11y` evaluates PresentationML structures to uncover digital accessibility barriers across all four POUR principles (**Perceivable**, **Operable**, **Understandable**, **Robust**). It applies non-destructive, deterministic remediations and compiles comprehensive audit reports in **Markdown** (the single source of truth), **accessible HTML5** (styled via SMACSS), and **accessible tagged PDF**.

Part of the document accessibility trio alongside [docx-a11y](https://github.com/ThirstyHead/docx-a11y) and [pdf-a11y](https://github.com/ThirstyHead/pdf-a11y).

---

## Key Features

- **POUR Structure (WCAG 2.2 AA):** Every finding and recommendation is categorized under Perceivable, Operable, Understandable, or Robust, linking to canonical W3C Understanding documentation.
- **Social Model of Disability:** Language focuses on document deficiencies and barriers rather than personal limitations. Unit-tested language guards prevent ableist or medical-model phrasing.
- **Single Source of Truth:** Reports are authored as CommonMark Markdown, then converted into accessible HTML5 and accessible tagged PDF with zero text discrepancies.
- **SMACSS Theme Engine:** Build-free, text-based CSS theme system with 6 bundled palettes (`light`, `dark`, `ocean`, `forest`, `high-contrast`, `print`). Supports custom user themes in `~/.config/pptx-a11y/themes/`.
- **Contrast Rigor (WCAG 1.4.3):** Every bundled theme and evaluated slide component enforces a minimum 4.5:1 text contrast ratio (3:1 for large text).
- **Interactive & Batch GUI:** Native cross-platform desktop interface for bulk directory remediation, file drag-and-drop, real-time progress, and visual triage.
- **Tagged Accessible PDF:** Exports multi-page PDFs carrying `/Lang`, `/Title`, `/MarkInfo /Marked true`, and validated tag trees.
- **Deterministic Remediation:** Auto-remediates presentation titles, table header rows, text run natural language tags, and section duplicates.

---

## Installation (macOS & Cross-Platform)

Requires **Python >= 3.10**.

### Recommended: Using `pipx` (Isolated CLI)

```bash
# Install directly from local repository or source checkout
pipx install /Users/scott/code/local/pptx-a11y

# Or editable install for active development:
pipx install --editable /Users/scott/code/local/pptx-a11y
```

### Using Python Virtual Environment (`venv`)

```bash
# 1. Clone or navigate to the repository
cd /Users/scott/code/local/pptx-a11y

# 2. Create virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 3. Install in editable mode with development dependencies
pip install -e ".[dev]"
```

Verify installation:
```bash
pptx-a11y --help
```

---

## Quickstart

### Example Test Presentations

Two sample presentations based on an artisanal bread baking recipe are provided in `examples/` for evaluating both the CLI and GUI:

- **`examples/No Knead Bread.pptx` (Clean Reference):**
  A fully accessible 5-slide presentation featuring structured slide placeholders, accessible table headers, descriptive image alt text, distinct slide titles, and natural language tags. Audits cleanly with **0 findings (100% WCAG 2.2 AA compliance)**.

- **`examples/No Knead Bread-test.pptx` (Accessibility Barriers):**
  Intentionally includes every digital accessibility barrier `pptx-a11y` checks: missing metadata title, missing/duplicate slide titles, missing image alt text, merged table cells, undeclared table headers, inverted reading order, uncaptioned media, missing language tags, default/duplicate section names, and vague link text.

```bash
# Test the clean deck (passes cleanly, exits 0):
pptx-a11y "examples/No Knead Bread.pptx" --format md,html,pdf,json --output-dir ./reports-clean

# Test the barrier deck (detects barriers, exits 1):
pptx-a11y "examples/No Knead Bread-test.pptx" --format md,html,pdf,json --output-dir ./reports-barriers
```

Or drag both files directly into the desktop GUI (`pptx-a11y --gui` or `pptx-a11y-gui`) to test batch analysis, progress tracking, and interactive triage!

---

### 1. Audit a Presentation

Audit a slide deck and generate all report formats (`.md`, `.html`, `.pdf`, `.json`):

```bash
pptx-a11y presentation.pptx --format md,html,pdf,json --output-dir ./reports
```

### 2. Remediate Violations

Apply deterministic fixes (adds missing title metadata, marks table header rows, tags run languages):

```bash
pptx-a11y presentation.pptx --fix --output-dir ./remediated
```
This saves `presentation-remediated.pptx` alongside before-and-after audit reports highlighting remediation progress.

### 3. Choose a Theme

Select one of the built-in accessible SMACSS themes for HTML and PDF output:

```bash
pptx-a11y presentation.pptx --format html,pdf --theme ocean --output-dir ./reports
```

Available bundled themes:
- `light`: Clean corporate theme (default).
- `dark`: High-legibility charcoal dark mode.
- `ocean`: Marine blue palette.
- `forest`: Botanical green palette.
- `high-contrast`: Maximum contrast pure black and yellow/cyan accents.
- `print`: Theme-independent black-on-white stylesheet relying on text-weight cues.

---

## Custom Branding & User Themes

You can define custom, human-editable themes without touching Python code. Place them in `~/.config/pptx-a11y/themes/<theme-name>/`:

```
~/.config/pptx-a11y/themes/brand/
├── theme.json
└── tokens.css
```

### `theme.json`
```json
{
  "name": "brand",
  "label": "Acme Brand Theme",
  "mode": "light",
  "default": false
}
```

### `tokens.css` (10 required CSS variables)
```css
:root {
  --bg: #ffffff;
  --fg: #1a1a1a;
  --muted: #cccccc;
  --accent: #0055aa;
  --link: #004488;
  --code-bg: #f5f5f5;
  --sev-critical: #b30000;
  --sev-serious: #cc5500;
  --sev-moderate: #886600;
  --sev-minor: #444444;
}
```

User themes automatically take precedence over bundled themes with matching names.

---

## WCAG 2.2 Coverage

| Rule ID | WCAG 2.2 SC | Level | Principle | Description | Auto-Fixable? |
|---|---|---|---|---|---|
| `image-alt-missing` | **1.1.1** Non-text Content | A | Perceivable | Visual asset lacks alt text or decorative flag | Manual |
| `table-header-missing` | **1.3.1** Info & Relationships | A | Perceivable | Table does not declare header row (`firstRow="1"`) | Yes |
| `semantic-placeholders-missing` | **1.3.2** Meaningful Sequence | A | Robust | Slide lacks structured layout placeholders | Manual |
| `color-contrast` | **1.4.3** Contrast (Minimum) | AA | Perceivable | Text fails 4.5:1 contrast against slide background | Manual |
| `title-missing` | **2.4.2** Page Titled | A | Operable | Presentation core metadata missing title | Yes |
| `slide-title-missing` | **2.4.2** Page Titled | A | Operable | Slide lacks structured title placeholder | Yes |
| `slide-title-duplicate` | **2.4.2** Page Titled | A | Operable | Slide title duplicated across slides | Manual |
| `link-text-vague` | **2.4.4** Link Purpose | A | Operable | Non-descriptive anchor text ("click here", "more") | Manual |
| `language-missing` | **3.1.1** Language of Page | A | Understandable | Text runs lack natural language tags | Yes |

---

## Graphical Desktop Application (GUI)

For content creators, educators, and teams remediating bulk presentations without writing terminal commands:

```bash
# Launch via CLI flag:
pptx-a11y --gui

# Or standalone launcher:
pptx-a11y-gui
```

### GUI Features:
- **Bulk Directory Ingestion:** Point to any folder to scan and remediate dozens of decks simultaneously.
- **Drag and Drop:** Drag `.pptx` presentations or entire directories directly into the window.
- **Visual Barrier Triage:** Click to resolve missing alt text or slide titles interactively with full context.
- **Custom Themes & Formats:** Output Markdown, HTML, tagged PDF, and audit JSON in any of the 6 SMACSS themes.
- **Native Packaging & Trusted Distribution:** See [docs/distribution.md](docs/distribution.md) for macOS Notarization, Windows Authenticode, and Linux AppImage guides.

---

## CLI Reference

```
usage: pptx-a11y [-h] [--format FORMAT] [--theme THEME] [--output-dir OUTPUT_DIR] [--fix] [--out-pptx OUT_PPTX] file

Audit and remediate PowerPoint presentations against WCAG 2.2 AA standards.

positional arguments:
  file                  Path to PowerPoint .pptx file

options:
  -h, --help            show this help message and exit
  --format FORMAT       Report formats (comma-separated): md, html, pdf, json (default: md)
  --theme THEME         SMACSS theme for HTML/PDF reports: ['light', 'dark', 'ocean', 'forest', 'high-contrast', 'print']
  --output-dir OUTPUT_DIR
                        Directory to save generated reports (default: current directory)
  --fix                 Perform deterministic remediation
  --out-pptx OUT_PPTX   Output path for remediated .pptx file
```

### Exit Codes
- `0`: Presentation passes all blocking WCAG 2.2 AA checks.
- `1`: Presentation contains unresolved blocking accessibility barriers.
- `2`: Input file error / invalid argument.

---

## Running Tests

```bash
pytest --cov=pptx_a11y tests/
```

---

## License

MIT License. See [LICENSE](LICENSE) for details.
