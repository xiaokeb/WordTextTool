"""Review the exact target names before changing any file path."""
from PyQt6.QtWidgets import QLineEdit, QCheckBox, QHBoxLayout, QPushButton, QLabel, QMessageBox, QTableWidgetItem
from word_text_tool.rename_service import preview_renames, execute_rename
from word_text_tool.models import FileResult, FileStatus
from word_text_tool.ui.task_panel import TaskPanel


class RenamePanel(TaskPanel):
    def __init__(self, parent=None):
        self.previews = ()
        super().__init__("所有文件 (*)", parent=parent)
        self.form.addWidget(QLabel("直接修改所选文件名，保留扩展名和文件内容；执行前请核对预览。"))
        row = QHBoxLayout()
        row.addWidget(QLabel("删除内容（空格精确匹配）"))
        self.suffix = QLineEdit(" - 修改版")
        row.addWidget(self.suffix)
        self.form.addLayout(row)
        self.all_matches = QCheckBox("移除主文件名中全部匹配串（默认只移除末尾一次）")
        self.form.addWidget(self.all_matches)
        buttons = QHBoxLayout()
        self.preview_button = QPushButton("预览名称")
        self.execute_button = QPushButton("执行重命名")
        self.preview_button.clicked.connect(self.preview)
        self.execute_button.clicked.connect(self.execute)
        buttons.addWidget(self.preview_button)
        buttons.addWidget(self.execute_button)
        self.form.addLayout(buttons)
        self.suffix.textChanged.connect(self.invalidate)
        self.all_matches.toggled.connect(self.invalidate)

    def invalidate(self, *args):
        self.previews = ()
        super().invalidate()

    def preview(self):
        if not self.files or not self.suffix.text():
            QMessageBox.warning(self, "无法预览", "请添加文件并输入待删除内容")
            return
        files, suffix, all_matches = tuple(self.files), self.suffix.text(), self.all_matches.isChecked()
        self.start_task([files], lambda _, cancelled: preview_renames(files, suffix, all_matches), self.preview_finished)

    def preview_finished(self, results, cancelled):
        if cancelled or not results or isinstance(results[0], FileResult):
            self.previews = ()
            return
        self.previews = results[0]
        for row, result in enumerate(self.previews):
            text = f"{result.target.name if result.target else '不能处理'} | {result.message}"
            item = QTableWidgetItem(text)
            item.setToolTip(str(result.target or result.source) + "\n" + result.message)
            self.table.setItem(row, 1, item)

    def execute(self):
        if self.locked or self.worker is not None:
            return
        if not self.previews:
            self.preview()
            return
        changed = [p for p in self.previews if p.target and p.target != p.source]
        if not changed:
            QMessageBox.information(self, "无需执行", "预览中没有可以重命名的文件")
            return
        message = "\n".join(f"{p.source.name} → {p.target.name}" for p in changed)
        confirmation = QMessageBox(QMessageBox.Icon.Question, "确认重命名", "将直接修改以下文件名：\n" + message[:1800],
                                  QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, self)
        confirmation.setDefaultButton(QMessageBox.StandardButton.No)
        confirmation.setDetailedText(message)
        if confirmation.exec() != QMessageBox.StandardButton.Yes:
            return
        self.start_task(self.previews, lambda preview, cancelled: execute_rename(preview), self.execution_finished)

    def execution_finished(self, results, cancelled):
        renamed = {r.source: r.output for r in results if r.status == FileStatus.SUCCESS}
        self.files = [renamed.get(path, path) for path in self.files]
        self.refresh()
        self.show_results(results, cancelled)
