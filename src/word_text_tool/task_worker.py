"""A sequential worker for immutable previews and operation results."""
from threading import Event
from PyQt6.QtCore import QThread, pyqtSignal
from word_text_tool.models import BatchSummary, FileResult, FileStatus
from word_text_tool.safe_output import error_message


class TaskWorker(QThread):
    progress = pyqtSignal(int, int)
    completed = pyqtSignal(object)

    def __init__(self, items, operation, parent=None):
        super().__init__(parent)
        self.items = tuple(items)
        self.operation = operation
        self.cancelled = Event()

    def request_cancel(self):
        self.cancelled.set()

    def run(self):
        results = []
        for index, item in enumerate(self.items):
            if self.cancelled.is_set():
                break
            try:
                results.append(self.operation(item, self.cancelled))
            except Exception as exc:
                source = getattr(item, "source", item)
                if isinstance(item, tuple):
                    source = item[0]
                results.append(FileResult(source, FileStatus.FAILED, error_message(exc)))
            self.progress.emit(index + 1, len(self.items))
        self.completed.emit((tuple(results), self.cancelled.is_set()))
