"""Tests for BatchWorker thread execution and Qt signals."""
from pathlib import Path
from pptx_a11y.gui.models import BatchItem
from pptx_a11y.gui.worker import BatchWorker

FIXTURES = Path(__file__).parent / "fixtures"


def test_batch_worker_execution(qtbot, tmp_path: Path):
    fixture_path = FIXTURES / "clean.pptx"
    item = BatchItem(path=fixture_path)

    worker = BatchWorker(
        items=[item],
        out_dir=tmp_path,
        formats=["md", "json"],
        theme="ocean",
        auto_fix=False,
    )

    started_signals = []
    finished_signals = []
    completed_signals = []

    worker.item_started.connect(lambda idx, tot, name: started_signals.append((idx, tot, name)))
    worker.item_finished.connect(lambda idx, st, cnt, sc: finished_signals.append((idx, st, cnt, sc)))
    worker.all_completed.connect(lambda tot, err: completed_signals.append((tot, err)))

    with qtbot.waitSignal(worker.all_completed, timeout=5000):
        worker.start()

    assert len(started_signals) == 1
    assert started_signals[0][2] == "clean.pptx"

    assert len(finished_signals) == 1
    assert item.error_message is None
    assert finished_signals[0][1] == "Completed"

    assert len(completed_signals) == 1
    assert completed_signals[0] == (1, 0)

    # Check reports were written
    assert (tmp_path / "clean-a11y-report.md").exists()
    assert (tmp_path / "clean-audit.json").exists()
