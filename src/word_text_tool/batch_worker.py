"""Run an immutable document batch sequentially in a Qt worker thread."""

from pathlib import Path
from threading import Event

from PyQt6.QtCore import QThread, pyqtSignal

from word_text_tool.document_processor import DocumentProcessor
from word_text_tool.models import (
    BatchSummary,
    FileResult,
    FileStatus,
    ReplacementRule,
)


class BatchWorker(QThread):
    """Process files in order and stop only at safe file boundaries."""

    progress = pyqtSignal(int, int, str, int)
    batch_finished = pyqtSignal(object)

    def __init__(
        self,
        processor: DocumentProcessor,
        files: tuple[Path, ...],
        rules: tuple[ReplacementRule, ...],
        custom_output_dir: Path | None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._processor = processor
        self._files = tuple(files)
        self._rules = tuple(rules)
        self._custom_output_dir = custom_output_dir
        self._cancel_requested = Event()

    def request_cancel(self) -> None:
        """Request cancellation after the current file finishes safely."""
        self._cancel_requested.set()

    def run(self) -> None:
        results: list[FileResult] = []
        total = len(self._files)

        for index, source in enumerate(self._files, start=1):
            if self._cancel_requested.is_set():
                break

            try:
                result = self._processor.process(
                    source,
                    self._rules,
                    self._custom_output_dir,
                )
            except Exception:
                result = FileResult(
                    source,
                    FileStatus.FAILED,
                    "处理文件时发生异常，已跳过该文件",
                )

            results.append(result)
            percent = round(index * 100 / total)
            self.progress.emit(index, total, source.name, percent)

        self.batch_finished.emit(
            BatchSummary(tuple(results), self._cancel_requested.is_set())
        )
