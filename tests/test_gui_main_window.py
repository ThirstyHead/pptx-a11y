"""Tests for GUI MainWindow layout, table binding, and controls."""
from pathlib import Path
from PySide6.QtCore import Qt
from pptx_a11y.gui.main_window import MainWindow

FIXTURES = Path(__file__).parent / "fixtures"


def test_main_window_initial_state(qtbot):
    win = MainWindow()
    qtbot.addWidget(win)

    assert win.windowTitle() == "pptx-a11y: PowerPoint WCAG Accessibility Remediation"
    assert win.table.columnCount() == 5
    assert win.table.rowCount() == 0
    assert win.btn_start.isEnabled() is False
    assert win.btn_stop.isEnabled() is False


def test_main_window_add_files_and_clear(qtbot):
    win = MainWindow()
    qtbot.addWidget(win)

    win.add_file_paths([FIXTURES / "clean.pptx", FIXTURES / "violations.pptx"])
    assert win.table.rowCount() == 2
    assert win.btn_start.isEnabled() is True

    # Clear table
    win.clear_files()
    assert win.table.rowCount() == 0
    assert win.btn_start.isEnabled() is False
