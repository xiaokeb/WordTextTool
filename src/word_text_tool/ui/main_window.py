"""单页响应式主窗口与安全批处理线程生命周期。"""

from pathlib import Path

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QCloseEvent, QFont
from PyQt6.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from word_text_tool.batch_worker import BatchWorker
from word_text_tool.document_processor import DocumentProcessor
from word_text_tool.file_service import OutputDirectoryError, prepare_output_directory
from word_text_tool.models import BatchSummary
from word_text_tool.ui.file_panel import FilePanel
from word_text_tool.ui.result_dialog import ResultDialog
from word_text_tool.ui.rule_panel import RulePanel


_DEFAULT_OUTPUT_TEXT = "未选择，将保存到源文件所在目录的‘修改版’文件夹"


class MainWindow(QMainWindow):
    def __init__(self, worker_factory=BatchWorker, parent=None) -> None:
        super().__init__(parent)
        self.worker_factory = worker_factory
        self.worker = None
        self.custom_output_dir: Path | None = None
        self._worker_summary: BatchSummary | None = None
        self._skipped_snapshot = ()
        self._close_after_worker = False
        self._result_dialogs: set[ResultDialog] = set()

        self.setWindowTitle("Word 文本处理工具")
        self.setFont(QFont("Microsoft YaHei UI"))
        self.resize(1000, 900)
        self.setMinimumSize(820, 700)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        content = QWidget(scroll_area)
        content.setMinimumWidth(760)
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(18, 16, 18, 18)
        content_layout.setSpacing(12)

        title = QLabel("Word 文本处理工具", content)
        title_font = QFont("Microsoft YaHei UI", 16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        introduction = QLabel(
            "批量选择 DOCX 文件，逐条设置精确替换规则，并将修改版安全保存到新文件。",
            content,
        )
        introduction.setWordWrap(True)
        introduction.setAlignment(Qt.AlignmentFlag.AlignCenter)
        content_layout.addWidget(title)
        content_layout.addWidget(introduction)

        self.file_panel = FilePanel(content)
        file_group = QGroupBox("选择文件", content)
        file_layout = QVBoxLayout(file_group)
        file_layout.addWidget(self.file_panel)
        content_layout.addWidget(file_group, 2)

        self.rule_panel = RulePanel(content)
        rule_group = QGroupBox("替换规则", content)
        rule_layout = QVBoxLayout(rule_group)
        rule_layout.addWidget(self.rule_panel)
        content_layout.addWidget(rule_group, 3)

        self.output_path_line_edit = QLineEdit(content)
        self.output_path_line_edit.setReadOnly(True)
        self.output_path_line_edit.setText(_DEFAULT_OUTPUT_TEXT)
        self.output_path_line_edit.setToolTip(_DEFAULT_OUTPUT_TEXT)
        self.output_browse_button = QPushButton("浏览", content)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_path_line_edit, 1)
        output_row.addWidget(self.output_browse_button)
        output_group = QGroupBox("保存位置", content)
        output_group.setLayout(output_row)
        content_layout.addWidget(output_group)

        self.progress_bar = QProgressBar(content)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_label = QLabel("准备就绪", content)
        self.status_label.setWordWrap(True)
        progress_layout = QVBoxLayout()
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        progress_group = QGroupBox("处理进度", content)
        progress_group.setLayout(progress_layout)
        content_layout.addWidget(progress_group)

        self.start_button = QPushButton("开始处理", content)
        self.cancel_button = QPushButton("取消处理", content)
        self.clear_button = QPushButton("清空所有", content)
        action_row = QHBoxLayout()
        action_row.addStretch(1)
        action_row.addWidget(self.start_button)
        action_row.addWidget(self.cancel_button)
        action_row.addWidget(self.clear_button)
        action_row.addStretch(1)
        content_layout.addLayout(action_row)

        scroll_area.setWidget(content)
        self.setCentralWidget(scroll_area)

        for button in self.findChildren(QPushButton):
            button.setMinimumWidth(110)
            button.setFixedHeight(36)

        self.output_browse_button.clicked.connect(self._browse_output_directory)
        self.start_button.clicked.connect(self._start_processing)
        self.cancel_button.clicked.connect(self._cancel_processing)
        self.clear_button.clicked.connect(self._clear_all)
        self._set_processing_state(False)

    def _browse_output_directory(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择保存目录")
        if not directory:
            return
        self.custom_output_dir = Path(directory).resolve(strict=False)
        display = str(self.custom_output_dir)
        self.output_path_line_edit.setText(display)
        self.output_path_line_edit.setToolTip(display)

    def _start_processing(self) -> None:
        if self.worker is not None:
            return

        files = self.file_panel.files()
        rules = self.rule_panel.rules()
        if not files:
            QMessageBox.warning(
                self, "无法开始", "请至少选择一个有效的 DOCX 文件"
            )
            return
        if not rules:
            QMessageBox.warning(self, "无法开始", "请至少添加一条替换规则")
            return

        if self.custom_output_dir is not None:
            try:
                prepare_output_directory(files[0], self.custom_output_dir)
            except OutputDirectoryError as exc:
                QMessageBox.warning(self, "保存路径不可用", str(exc))
                return

        files_snapshot = tuple(files)
        rules_snapshot = tuple(rules)
        skipped_snapshot = tuple(self.file_panel.skipped_results())
        candidate = None
        self._set_processing_state(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在准备处理…")
        self._worker_summary = None
        self._skipped_snapshot = skipped_snapshot

        try:
            candidate = self.worker_factory(
                DocumentProcessor(),
                files_snapshot,
                rules_snapshot,
                self.custom_output_dir,
                self,
            )
            self.worker = candidate
            candidate.progress.connect(
                lambda current, total, filename, percent, source=candidate: (
                    self._dispatch_progress(
                        source, current, total, filename, percent
                    )
                )
            )
            candidate.batch_finished.connect(
                lambda summary, source=candidate: self._on_batch_finished(
                    summary, source
                )
            )
            candidate.finished.connect(
                lambda source=candidate: self._on_worker_finished(source)
            )
            candidate.start()
        except Exception:
            self.worker = None
            self._worker_summary = None
            self._skipped_snapshot = ()
            self._set_processing_state(False)
            self.status_label.setText("启动失败")
            if candidate is not None and hasattr(candidate, "deleteLater"):
                candidate.deleteLater()
            QMessageBox.warning(
                self, "无法开始", "后台处理线程启动失败，请重试"
            )

    def _dispatch_progress(
        self,
        worker,
        current: int,
        total: int,
        filename: str,
        percent: int,
    ) -> None:
        if self.worker is worker:
            self._on_progress(current, total, filename, percent)

    def _on_progress(
        self, current: int, total: int, filename: str, percent: int
    ) -> None:
        self.progress_bar.setValue(max(0, min(100, percent)))
        self.status_label.setText(
            f"正在处理：{filename} ({current}/{total})"
        )

    def _cancel_processing(self) -> None:
        worker = self.worker
        if worker is not None and worker.isRunning():
            self.status_label.setText("正在安全结束当前文件…")
            worker.request_cancel()
            self.cancel_button.setEnabled(False)

    def _on_batch_finished(self, worker_summary: BatchSummary, worker=None) -> None:
        source = self.worker if worker is None else worker
        if source is None or self.worker is not source:
            return
        results = self._skipped_snapshot + tuple(worker_summary.results)
        self._worker_summary = BatchSummary(results, worker_summary.cancelled)
        self.status_label.setText("正在结束后台线程…")

    def _on_worker_finished(self, worker) -> None:
        if self.worker is not worker:
            return
        # batch_finished 由 run() 内发出，finished 由 QThread 在 run()
        # 返回后发出。延后一个事件循环以容纳跨线程队列中的汇总信号。
        QTimer.singleShot(0, lambda source=worker: self._finalize_worker(source))

    def _finalize_worker(self, worker) -> None:
        if self.worker is not worker:
            return

        summary = self._worker_summary
        if summary is None:
            summary = BatchSummary(tuple(self._skipped_snapshot))

        self.worker = None
        self._worker_summary = None
        self._skipped_snapshot = ()
        self._set_processing_state(False)
        self.status_label.setText("处理已取消" if summary.cancelled else "处理完成")
        if hasattr(worker, "deleteLater"):
            worker.deleteLater()

        if self._close_after_worker:
            self._close_after_worker = False
            QTimer.singleShot(0, self.close)
            return
        self._show_result_dialog(summary)

    def _show_result_dialog(self, summary: BatchSummary) -> None:
        dialog = ResultDialog(summary, self)
        dialog.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self._result_dialogs.add(dialog)
        dialog.finished.connect(
            lambda _result, item=dialog: self._result_dialogs.discard(item)
        )
        dialog.open()

    def _set_processing_state(self, processing: bool) -> None:
        self.file_panel.set_locked(processing)
        self.rule_panel.set_locked(processing)
        self.output_browse_button.setEnabled(not processing)
        self.start_button.setEnabled(not processing)
        self.clear_button.setEnabled(not processing)
        self.cancel_button.setEnabled(processing)

    def _clear_all(self) -> None:
        has_content = bool(self.file_panel.files() or self.rule_panel.rules())
        if has_content and QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有文件、规则和保存设置吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return

        self.file_panel.clear_files()
        self.rule_panel.clear_rules()
        self.custom_output_dir = None
        self.output_path_line_edit.setText(_DEFAULT_OUTPUT_TEXT)
        self.output_path_line_edit.setToolTip(_DEFAULT_OUTPUT_TEXT)
        self.progress_bar.setValue(0)
        self.status_label.setText("准备就绪")

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 - Qt API
        worker = self.worker
        if worker is None:
            self._close_after_worker = False
            event.accept()
            return

        if self._close_after_worker:
            event.ignore()
            return

        if worker.isRunning():
            answer = QMessageBox.question(
                self,
                "确认退出",
                "文件仍在处理。退出前将安全完成当前文件，是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            # The modal question runs its own event loop. finished/finalize may
            # already have released the worker while the user was answering.
            if self.worker is None:
                self._close_after_worker = False
                event.accept()
                return
            if self.worker is not worker:
                event.ignore()
                return
            self._close_after_worker = True
            self.status_label.setText("正在安全结束当前文件…")
            if worker.isRunning():
                worker.request_cancel()
            self.cancel_button.setEnabled(False)
            event.ignore()
            return

        # run() 已返回但 finished 还在主线程事件队列中时，
        # 仍由 finished 的收尾逻辑释放对象并二次关闭。
        self._close_after_worker = True
        event.ignore()
