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


def test_main_window_two_pane_story_components(qtbot):
    win = MainWindow()
    qtbot.addWidget(win)

    # Before pane
    assert hasattr(win, "pane_before")
    assert win.pane_before.objectName() == "pane_before"
    assert win.table_before == win.table

    # Center bridge
    assert hasattr(win, "center_bridge")
    assert win.center_bridge.objectName() == "center_bridge"
    assert hasattr(win, "btn_remediate_bridge")
    assert win.btn_start == win.btn_remediate_bridge

    # After pane
    assert hasattr(win, "pane_after")
    assert win.pane_after.objectName() == "pane_after"
    assert hasattr(win, "tree_after")
    assert hasattr(win, "btn_open_output_folder")
    assert hasattr(win, "btn_view_report")
    assert hasattr(win, "guide_banner")


def test_main_window_after_tree_population_and_view(qtbot, tmp_path: Path):
    win = MainWindow()
    qtbot.addWidget(win)

    win.txt_out_dir.setText(str(tmp_path))

    # Create dummy output files
    fixed_pptx = tmp_path / "clean.fixed.pptx"
    fixed_pptx.write_text("dummy fixed pptx content", encoding="utf-8")
    md_report = tmp_path / "clean-a11y-report.md"
    md_report.write_text("# Clean Report\nScore: 100%", encoding="utf-8")

    from pptx_a11y.gui.models import BatchItem
    item = BatchItem(path=FIXTURES / "clean.pptx")
    item.score = 100.0
    item.status = "Remediated"

    win._populate_after_tree_for_item(item)

    assert win.tree_after.topLevelItemCount() == 1
    doc_node = win.tree_after.topLevelItem(0)
    assert doc_node is not None
    assert "clean.pptx" in doc_node.text(0)

    # Children: fixed pptx and md report
    assert doc_node.childCount() == 2
    child_texts = [c.text(0) for i in range(doc_node.childCount()) if (c := doc_node.child(i)) is not None]
    assert any("Fixed PPTX" in t for t in child_texts)
    assert any("Audit Report" in t for t in child_texts)

