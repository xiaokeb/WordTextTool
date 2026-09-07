"""按四种固定状态展示可选择、可复制的批处理结果。"""

from PyQt6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QPlainTextEdit,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from word_text_tool.models import BatchSummary, FileResult, FileStatus


_STATUS_TABS = (
    (FileStatus.SUCCESS, "处理成功"),
    (FileStatus.UNMATCHED, "未匹配"),
    (FileStatus.SKIPPED, "已跳过"),
    (FileStatus.FAILED, "处理失败"),
)


def _result_line(result: FileResult) -> str:
    parts = [result.source.name, result.message]
    if result.output is not None:
        parts.append(str(result.output))
    return " | ".join(parts)


class ResultDialog(QDialog):
    def __init__(self, summary: BatchSummary, parent=None) -> None:
        super().__init__(parent)
        self.summary = summary
        self.setWindowTitle("处理结果")

        self.tabs = QTabWidget(self)
        self.editors: dict[FileStatus, QPlainTextEdit] = {}
        for status, title in _STATUS_TABS:
            results = tuple(
                result for result in summary.results if result.status is status
            )
            page = QWidget(self.tabs)
            page_layout = QVBoxLayout(page)
            editor = QPlainTextEdit(page)
            editor.setReadOnly(True)
            editor.setPlainText("\n".join(_result_line(result) for result in results))
            page_layout.addWidget(editor)
            self.editors[status] = editor
            self.tabs.addTab(page, f"{title} ({len(results)})")

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(self.tabs)
        layout.addWidget(buttons)
        self.resize(780, 520)
