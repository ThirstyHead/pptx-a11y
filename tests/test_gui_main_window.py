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


def test_main_window_worker_callbacks(qtbot, tmp_path: Path):
    win = MainWindow()
    qtbot.addWidget(win)

    win.add_file_paths([FIXTURES / "clean.pptx"])
    win._on_item_started(0, 1, "clean.pptx")
    assert "clean.pptx" in win.lbl_status.text()

    win._on_item_progress(0, "Analyzing shapes...")
    assert "Analyzing shapes" in win.lbl_status.text()

    win._on_item_finished(0, "Completed", 0, 100.0)
    assert win.progress_bar.value() == 1

    win._on_all_completed(1, 0)
    assert "Completed" in win.lbl_status.text()
