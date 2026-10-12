"""Shared batch controls; previews run in a worker, never on the GUI thread."""
from pathlib import Path
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QTableWidget,
    QTableWidgetItem, QAbstractItemView, QHeaderView, QLabel, QLineEdit,
    QProgressBar, QMessageBox, QDialog, QPlainTextEdit, QDialogButtonBox,
)
from word_text_tool.task_worker import TaskWorker
from word_text_tool.models import BatchSummary, FileResult, FileStatus
from word_text_tool.safe_output import fingerprint, error_message
from word_text_tool.ui.result_dialog import ResultDialog


class TaskPanel(QWidget):
    busy_changed = pyqtSignal(bool)
    idle = pyqtSignal()

    def __init__(self, file_filter, suffixes=None, parent=None):
        super().__init__(parent)
        self.file_filter = file_filter
        self.suffixes = suffixes
        self.files = []
        self.worker = None
        self._pending = None
        self._finish_callback = None
        self.locked = False
        self.dialogs = set()
        self.layout = QVBoxLayout(self)
        self.controls = QWidget(self)
        self.form = QVBoxLayout(self.controls)
        self.layout.addWidget(self.controls)
        row = QHBoxLayout()
        for text, slot in [("添加文件", self.browse_files), ("浏览文件夹", self.browse_folder),
                           ("删除文件", self.remove_selected), ("上移", lambda: self.move(-1)),
                           ("下移", lambda: self.move(1)), ("清空列表", self.clear_files)]:
            button = QPushButton(text)
            button.clicked.connect(slot)
            row.addWidget(button)
        self.form.addLayout(row)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["文件", "预览 / 状态"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setMinimumHeight(190)
        self.table.cellDoubleClicked.connect(self.show_detail)
        self.form.addWidget(self.table)
        self.status = QLabel("准备就绪；双击行可查看完整预览")
        self.status.setWordWrap(True)
        self.progress = QProgressBar()
        self.cancel_button = QPushButton("取消处理")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel)
        self.layout.addWidget(self.status)
        self.layout.addWidget(self.progress)
        self.layout.addWidget(self.cancel_button)

    def add_output(self):
        row = QHBoxLayout()
        self.output = QLineEdit()
        self.output.setReadOnly(True)
        self.output.setPlaceholderText("默认保存到每个源文件同级的“修改版”目录")
        select = QPushButton("保存目录")
        select.clicked.connect(self.browse_output)
        reset = QPushButton("恢复默认")
        reset.clicked.connect(self.output.clear)
        self.output.textChanged.connect(self.invalidate)
        row.addWidget(self.output)
        row.addWidget(select)
        row.addWidget(reset)
        self.form.addLayout(row)

    def browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择保存目录")
        if path:
            self.output.setText(path)
            self.output.setToolTip(path)

    def output_dir(self):
        return Path(self.output.text()) if self.output.text() else None

    def browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "添加文件", "", self.file_filter)
        self.add_paths(paths)

    def browse_folder(self):
        directory = QFileDialog.getExistingDirectory(self, "扫描当前层文件")
        if directory:
            try:
                paths = sorted(Path(directory).iterdir(), key=lambda p: p.name.casefold())
                self.add_paths([p for p in paths if p.is_file() and
                                (self.suffixes is None or p.suffix.lower() in self.suffixes)])
            except OSError:
                QMessageBox.warning(self, "无法扫描", "文件夹无法读取")

    def add_paths(self, paths):
        if self.locked or self.worker is not None:
            return
        skipped = []
        for raw in paths:
            path = Path(raw).absolute()
            try:
                if path.is_symlink():
                    raise ValueError("不处理符号链接")
                path = path.resolve()
                if path in self.files:
                    raise ValueError("重复文件")
                if self.suffixes is not None and path.suffix.lower() not in self.suffixes:
                    raise ValueError("仅支持 .xlsx 文件")
                if not path.is_file() or path.is_symlink() or path.stat().st_size > 50 * 1024 * 1024:
                    raise ValueError("不是普通文件、符号链接或超过 50 MiB")
                self.files.append(path)
            except (ValueError, OSError) as exc:
                skipped.append(f"{path.name}：{str(exc) if isinstance(exc, ValueError) else '无法读取'}")
        self.refresh()
        if skipped:
            self.show_text("已跳过", "\n".join(skipped))

    def refresh(self):
        self.table.setRowCount(len(self.files))
        for i, path in enumerate(self.files):
            item = QTableWidgetItem(path.name)
            item.setToolTip(str(path))
            self.table.setItem(i, 0, item)
        self.invalidate()

    def invalidate(self, *args):
        for row in range(self.table.rowCount()):
            self.table.setItem(row, 1, QTableWidgetItem("待预览"))

    def remove_selected(self):
        row = self.table.currentRow()
        if row >= 0:
            self.files.pop(row)
            self.refresh()

    def clear_files(self):
        if self.files and QMessageBox.question(self, "清空列表", "确认清空本页文件列表？") != QMessageBox.StandardButton.Yes:
            return
        self.files.clear()
        self.refresh()

    def move(self, offset):
        row = self.table.currentRow()
        target = row + offset
        if 0 <= row < len(self.files) and 0 <= target < len(self.files):
            self.files[row], self.files[target] = self.files[target], self.files[row]
            self.refresh()
            self.table.selectRow(target)

    def set_locked(self, locked):
        self.locked = locked
        self.controls.setEnabled(not locked)

    def start_task(self, items, operation, finished):
        if self.worker is not None or self.locked:
            return
        self._pending = None
        self._finish_callback = finished
        worker = TaskWorker(items, operation, self)
        self.worker = worker
        worker.completed.connect(lambda value, w=worker: self.receive(w, value))
        worker.progress.connect(lambda current, total, w=worker: self.update_progress(w, current, total))
        worker.finished.connect(lambda w=worker: QTimer.singleShot(0, lambda: self.finished(w)))
        self.set_locked(True)
        self.cancel_button.setEnabled(True)
        self.progress.setValue(0)
        self.status.setText("正在处理，请稍候…")
        self.busy_changed.emit(True)
        try:
            worker.start()
        except Exception:
            self.worker = None
            worker.deleteLater()
            self.set_locked(False)
            self.cancel_button.setEnabled(False)
            self.busy_changed.emit(False)
            self.status.setText("后台任务启动失败，请重试")
            self.idle.emit()

    def update_progress(self, worker, current, total):
        if self.worker is worker and total:
            self.progress.setValue(round(current * 100 / total))

    def closing(self):
        window = self.window()
        return bool(getattr(window, "_close_after_extra", False) or getattr(window, "_close_prompt_active", False))

    def receive(self, worker, value):
        if self.worker is worker:
            self._pending = value

    def finished(self, worker):
        if self.worker is not worker:
            return
        callback = self._finish_callback
        pending = self._pending
        self.worker = None
        worker.deleteLater()
        self.set_locked(False)
        self.cancel_button.setEnabled(False)
        self.status.setText("已取消" if pending and pending[1] else "处理完成")
        self.busy_changed.emit(False)
        try:
            if pending is not None and not self.closing():
                callback(*pending)
            elif pending is not None and getattr(self.window(), "_close_prompt_active", False):
                self._deferred_finish = (callback, pending)
        except Exception as exc:
            self.status.setText(error_message(exc))
        finally:
            self.idle.emit()

    def flush_deferred(self, show):
        deferred = getattr(self, "_deferred_finish", None)
        self._deferred_finish = None
        if show and deferred is not None:
            callback, pending = deferred
            QTimer.singleShot(0, lambda: callback(*pending) if not self.closing() else None)

    def cancel(self):
        if self.worker is not None:
            self.worker.request_cancel()
            self.cancel_button.setEnabled(False)
            self.status.setText("正在安全取消，等待当前文件完成（办公操作超时上限 120 秒）…")

    def show_results(self, results, cancelled):
        dialog = ResultDialog(BatchSummary(tuple(results), cancelled), self)
        self.keep_dialog(dialog)

    def keep_dialog(self, dialog):
        self.dialogs.add(dialog)
        dialog.finished.connect(lambda _: (self.dialogs.discard(dialog), dialog.deleteLater()))
        dialog.open()

    def show_text(self, title, text):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        editor = QPlainTextEdit(text)
        editor.setReadOnly(True)
        layout.addWidget(editor)
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        layout.addWidget(close)
        dialog.resize(780, 450)
        self.keep_dialog(dialog)

    def show_detail(self, row, column):
        text = self.table.item(row, 1)
        if text:
            self.show_text("完整预览", str(self.files[row]) + "\n" + (text.toolTip() or text.text()))
