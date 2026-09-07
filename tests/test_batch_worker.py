from pathlib import Path
from threading import Event

from PyQt6.QtTest import QSignalSpy

from word_text_tool.batch_worker import BatchWorker
from word_text_tool.models import FileResult, FileStatus, ReplacementRule


class RecordingProcessor:
    def __init__(self, *, fail_on: Path | None = None, started=None, release=None):
        self.fail_on = fail_on
        self.started = started
        self.release = release
        self.calls = []

    def process(self, source, rules, output_dir):
        self.calls.append((source, rules, output_dir))
        if self.started is not None and len(self.calls) == 1:
            self.started.set()
            self.release.wait(timeout=3)
        if source == self.fail_on:
            raise RuntimeError("sensitive implementation detail")
        return FileResult(
            source,
            FileStatus.SUCCESS,
            "完成",
            source.with_suffix(".out.docx"),
            1,
        )


def _wait_for_worker(qtbot, worker, spy):
    qtbot.waitUntil(lambda: len(spy) == 1, timeout=3000)
    qtbot.waitUntil(lambda: not worker.isRunning(), timeout=3000)
    assert worker.wait(3000)
    return spy[0][0]


def test_worker_processes_in_frozen_order_and_emits_progress(qtbot):
    processor = RecordingProcessor()
    files = [Path("a.docx"), Path("b.docx")]
    rules = [ReplacementRule("旧", "新")]
    worker = BatchWorker(processor, files, rules, None)
    progress = QSignalSpy(worker.progress)
    completed = QSignalSpy(worker.batch_finished)

    files.append(Path("late.docx"))
    rules.append(ReplacementRule("后加", "不会使用"))
    try:
        worker.start()
        summary = _wait_for_worker(qtbot, worker, completed)
    finally:
        worker.request_cancel()
        worker.wait(3000)

    assert [call[0] for call in processor.calls] == [Path("a.docx"), Path("b.docx")]
    assert all(call[1] == (ReplacementRule("旧", "新"),) for call in processor.calls)
    assert [tuple(item) for item in progress] == [
        (1, 2, "a.docx", 50),
        (2, 2, "b.docx", 100),
    ]
    assert [result.source for result in summary.results] == [
        Path("a.docx"),
        Path("b.docx"),
    ]
    assert summary.cancelled is False


def test_worker_converts_one_unexpected_error_and_continues(qtbot):
    failed = Path("bad.docx")
    processor = RecordingProcessor(fail_on=failed)
    worker = BatchWorker(
        processor,
        (failed, Path("good.docx")),
        (ReplacementRule("旧", "新"),),
        None,
    )
    completed = QSignalSpy(worker.batch_finished)
    try:
        worker.start()
        summary = _wait_for_worker(qtbot, worker, completed)
    finally:
        worker.request_cancel()
        worker.wait(3000)

    assert [call[0] for call in processor.calls] == [failed, Path("good.docx")]
    assert [result.status for result in summary.results] == [
        FileStatus.FAILED,
        FileStatus.SUCCESS,
    ]
    assert "sensitive implementation detail" not in summary.results[0].message


def test_cancel_waits_for_current_file_and_skips_remaining(qtbot):
    started = Event()
    release = Event()
    processor = RecordingProcessor(started=started, release=release)
    worker = BatchWorker(
        processor,
        (Path("a.docx"), Path("b.docx")),
        (ReplacementRule("旧", "新"),),
        None,
    )
    completed = QSignalSpy(worker.batch_finished)
    try:
        worker.start()
        assert started.wait(timeout=2)
        worker.request_cancel()
        assert worker.isRunning()
        release.set()
        summary = _wait_for_worker(qtbot, worker, completed)
    finally:
        release.set()
        worker.request_cancel()
        worker.wait(3000)

    assert [call[0] for call in processor.calls] == [Path("a.docx")]
    assert len(summary.results) == 1
    assert summary.cancelled is True


def test_cancel_before_start_processes_nothing(qtbot):
    processor = RecordingProcessor()
    worker = BatchWorker(
        processor,
        (Path("a.docx"),),
        (ReplacementRule("旧", "新"),),
        None,
    )
    completed = QSignalSpy(worker.batch_finished)
    try:
        worker.request_cancel()
        worker.start()
        summary = _wait_for_worker(qtbot, worker, completed)
    finally:
        worker.request_cancel()
        worker.wait(3000)

    assert processor.calls == []
    assert summary.results == ()
    assert summary.cancelled is True


def test_empty_queue_emits_uncancelled_empty_summary(qtbot):
    processor = RecordingProcessor()
    worker = BatchWorker(processor, (), (), None)
    completed = QSignalSpy(worker.batch_finished)
    try:
        worker.start()
        summary = _wait_for_worker(qtbot, worker, completed)
    finally:
        worker.request_cancel()
        worker.wait(3000)

    assert processor.calls == []
    assert summary.results == ()
    assert summary.cancelled is False
