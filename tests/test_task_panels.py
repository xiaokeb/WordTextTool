from threading import Event
from PyQt6.QtWidgets import QMessageBox
from word_text_tool.ui.main_window import MainWindow
from word_text_tool.models import FileResult, FileStatus
from word_text_tool.task_worker import TaskWorker


def test_three_tabs_and_global_lock_released_only_after_worker(qtbot, tmp_path):
    window = MainWindow()
    qtbot.addWidget(window)
    assert window.tabs.count() == 3
    wait = Event()
    path = tmp_path / "source.txt"
    def operation(item, cancelled):
        wait.wait(2)
        return FileResult(path, FileStatus.SUCCESS, "ok")
    results = []
    window.rename_panel.start_task([path], operation, lambda *args: results.append(args))
    try:
        assert not window.start_button.isEnabled()
        assert window.number_panel.locked
        assert window.rename_panel.worker is not None
    finally:
        wait.set()
    qtbot.waitUntil(lambda: window.rename_panel.worker is None)
    assert window.start_button.isEnabled()
    assert not window.number_panel.locked
    assert results


def test_worker_cancel_preserves_current_file_and_skips_rest(qtbot, tmp_path):
    entered, release = Event(), Event()
    path = tmp_path / "file"
    calls = []
    def operation(item, cancelled):
        calls.append(item)
        entered.set()
        release.wait(2)
        return FileResult(path, FileStatus.SUCCESS, "done")
    worker = TaskWorker([1, 2, 3], operation)
    results = []
    worker.completed.connect(results.append)
    worker.start()
    assert entered.wait(2)
    worker.request_cancel()
    release.set()
    qtbot.waitUntil(lambda: bool(results))
    worker.wait()
    assert calls == [1]
    assert results[0][1]


def test_number_preview_invalidation(qtbot):
    window = MainWindow()
    qtbot.addWidget(window)
    panel = window.number_panel
    panel.previews = ("stale",)
    panel.template.setText("合同-yy")
    assert panel.previews == ()
    panel.mode.setCurrentIndex(1)
    panel.fixed_text.setText("固定编号")
    assert panel.settings().fixed


def test_close_during_preview_does_not_launch_followup(qtbot, tmp_path, monkeypatch):
    window = MainWindow()
    qtbot.addWidget(window)
    window.show()
    release = Event()
    followups = []
    def operation(item, cancelled):
        release.wait(2)
        return item
    window.number_panel.start_task([tmp_path], operation, lambda *args: followups.append(args))
    def answer(*args):
        release.set()
        qtbot.waitUntil(lambda: window.number_panel.worker is None)
        return QMessageBox.StandardButton.Yes
    monkeypatch.setattr(QMessageBox, "question", answer)
    window.close()
    assert not followups
    assert not window.isVisible()
