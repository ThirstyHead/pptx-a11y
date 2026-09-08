"""Visual interactive remediation dialog for author-intent accessibility barriers."""
from pathlib import Path
from typing import Any, Dict, List
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)
from pptx import Presentation
from ..triage import _mark_shape_decorative, _set_shape_alt_text, _set_slide_title


class TriageDialog(QDialog):
    def __init__(self, pptx_path: Path | str, findings: List[Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.pptx_path = Path(pptx_path)
        self.prs = Presentation(str(self.pptx_path))

        # Filter triageable findings
        self.findings = [
            f
            for f in findings
            if f.get("rule_id")
            in ("image-alt-missing", "chart-missing-alt", "slide-title-missing", "slide-title-duplicate")
        ]
        self.current_idx = 0
        self.items_modified = 0

        self.setWindowTitle("Interactive Remediation Triage")
        self.resize(600, 320)
        self._init_ui()
        self._load_current()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        self.lbl_progress = QLabel("Barrier 1 of 1")
        self.lbl_progress.setStyleSheet("font-weight: 600; color: #2563eb;")
        layout.addWidget(self.lbl_progress)

        self.lbl_location = QLabel("Location:")
        self.lbl_location.setStyleSheet("font-weight: 600; color: #334155;")
        layout.addWidget(self.lbl_location)

        self.lbl_desc = QLabel("Description:")
        self.lbl_desc.setWordWrap(True)
        layout.addWidget(self.lbl_desc)

        # Input row
        self.lbl_input_prompt = QLabel("Descriptive Alternative Text:")
        layout.addWidget(self.lbl_input_prompt)

        self.txt_input = QLineEdit()
        layout.addWidget(self.txt_input)

        self.chk_decorative = QCheckBox("Mark as Decorative (Purely aesthetic, ignore in screen readers)")
        self.chk_decorative.toggled.connect(self._on_decorative_toggled)
        layout.addWidget(self.chk_decorative)

        layout.addStretch()

        # Action Buttons
        btn_layout = QHBoxLayout()
        self.btn_skip = QPushButton("Skip")
        self.btn_skip.clicked.connect(self.skip_current)

        self.btn_apply = QPushButton("Apply & Next")
        self.btn_apply.setStyleSheet("background-color: #2563eb; color: white; font-weight: 600;")
        self.btn_apply.clicked.connect(self.apply_current)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_skip)
        btn_layout.addWidget(self.btn_apply)
        layout.addLayout(btn_layout)

    def _on_decorative_toggled(self, checked: bool):
        self.txt_input.setEnabled(not checked)
        if checked:
            self.txt_input.clear()

    def _load_current(self):
        if not self.findings or self.current_idx >= len(self.findings):
            self._finish()
            return

        f = self.findings[self.current_idx]
        tot = len(self.findings)
        self.lbl_progress.setText(f"Barrier {self.current_idx + 1} of {tot}")
        self.lbl_location.setText(f"Location: {f.get('location', '')}")
        self.lbl_desc.setText(f.get("description", ""))

        rule_id = f.get("rule_id", "")
        self.chk_decorative.setChecked(False)
        self.txt_input.clear()
        self.txt_input.setEnabled(True)

        if rule_id in ("image-alt-missing", "chart-missing-alt"):
            self.lbl_input_prompt.setText("Descriptive Alternative Text:")
            self.chk_decorative.setVisible(True)
        elif rule_id == "slide-title-missing":
            self.lbl_input_prompt.setText("New Slide Title:")
            self.chk_decorative.setVisible(False)
        elif rule_id == "slide-title-duplicate":
            self.lbl_input_prompt.setText("Distinct Slide Title:")
            self.chk_decorative.setVisible(False)

    def apply_current(self):
        if self.current_idx >= len(self.findings):
            return

        f = self.findings[self.current_idx]
        rule_id = f.get("rule_id", "")
        loc = f.get("location", "")

        if self.chk_decorative.isChecked():
            _mark_shape_decorative(self.prs, loc)
            self.items_modified += 1
        else:
            text = self.txt_input.text().strip()
            if text:
                if rule_id in ("image-alt-missing", "chart-missing-alt"):
                    _set_shape_alt_text(self.prs, loc, text)
                    self.items_modified += 1
                elif rule_id in ("slide-title-missing", "slide-title-duplicate"):
                    _set_slide_title(self.prs, loc, text)
                    self.items_modified += 1

        # Advance
        self.current_idx += 1
        if self.current_idx < len(self.findings):
            self._load_current()
        else:
            self._finish()

    def skip_current(self):
        self.current_idx += 1
        if self.current_idx < len(self.findings):
            self._load_current()
        else:
            self._finish()

    def _finish(self):
        if self.items_modified > 0:
            self.prs.save(str(self.pptx_path))
        self.accept()
