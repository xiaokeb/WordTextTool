"""Numbering preview and explicit three-action Office workflow."""
from PyQt6.QtWidgets import QComboBox, QLineEdit, QSpinBox, QFormLayout, QHBoxLayout, QPushButton, QLabel, QMessageBox, QTableWidgetItem
from word_text_tool.numbering_service import NumberSettings, NumberPreview, preview_number
from word_text_tool.excel_processor import process_excel
from word_text_tool.models import FileResult, FileStatus
from word_text_tool.safe_output import available_path
from word_text_tool.ui.task_panel import TaskPanel


class NumberPanel(TaskPanel):
    def __init__(self, parent=None):
        self.previews = ()
        self.preview_action = None
        super().__init__("Excel 工作簿 (*.xlsx)", {".xlsx"}, parent)
        self.form.addWidget(QLabel("按第一张到最后一张工作表编号；所有工作表须可见且有效。"))
        self.form.addWidget(QLabel("固定模式先插入新首行。编号区域合并、居中、加粗，字体 Times New Roman。"))
        form = QFormLayout()
        self.mode = QComboBox()
        self.mode.addItems(["自动编号", "固定编号 + 递增编号（先插入首行）"])
        self.template = QLineEdit("编号-yy")
        self.fixed_text = QLineEdit()
        self.start = QSpinBox()
        self.start.setRange(1, 999999999)
        self.start.setValue(1)
        self.engine = QComboBox()
        self.engine.addItems(["自动检测", "Microsoft Excel", "WPS 表格"])
        for title, widget in [("模式", self.mode), ("编号模板（一个 yy）", self.template),
                              ("固定编号", self.fixed_text), ("起始序号", self.start), ("办公软件", self.engine)]:
            form.addRow(title, widget)
        self.form.addLayout(form)
        self.add_output()
        buttons = QHBoxLayout()
        for title, callback in [("预览编号", lambda: self.preview("number")),
                                ("写入编号", lambda: self.execute("number")),
                                ("导出 PDF", lambda: self.execute("pdf")),
                                ("一键转换", lambda: self.execute("convert")),
                                ("重置编号", lambda: self.start.setValue(1))]:
            button = QPushButton(title)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        self.form.addLayout(buttons)
        self.template.textChanged.connect(self.invalidate)
        self.fixed_text.textChanged.connect(self.invalidate)
        self.mode.currentIndexChanged.connect(self.invalidate)
        self.engine.currentIndexChanged.connect(self.invalidate)
        self.start.valueChanged.connect(self.invalidate)

    def settings(self):
        return NumberSettings(self.template.text(), self.start.value(), self.mode.currentIndex() == 1,
                              self.fixed_text.text(), ("auto", "excel", "wps")[self.engine.currentIndex()])

    def invalidate(self, *args):
        self.previews = ()
        self.preview_action = None
        super().invalidate()

    def preview(self, action, then_execute=False):
        if not self.files:
            QMessageBox.warning(self, "无法预览", "请先添加 .xlsx 文件")
            return
        settings = self.settings()
        if action != "pdf":
            try:
                settings.validate()
            except ValueError as exc:
                QMessageBox.warning(self, "编号设置无效", str(exc))
                return
        items = tuple((path, index) for index, path in enumerate(self.files))
        operation = lambda item, cancelled: preview_number(item[0], settings, item[1], pdf_only=action == "pdf")
        self.start_task(items, operation, lambda results, cancelled: self.preview_finished(results, cancelled, action, then_execute))

    def preview_finished(self, results, cancelled, action, then_execute):
        self.previews = results if not cancelled else ()
        self.preview_action = "pdf" if action == "pdf" else "number"
        reserved = []
        for row, result in enumerate(results):
            if isinstance(result, NumberPreview):
                directory = self.output_dir() or result.source.parent / "修改版"
                extension = ".pdf" if action == "pdf" else ".xlsx"
                stem = result.source.stem if action == "pdf" else result.source.stem + " - 编号版"
                path = available_path(directory, stem, extension, reserved)
                reserved.append(path)
                text = (f"编号：{result.number or '仅导出 PDF'}\n工作表：" + "、".join(result.sheets)
                        + f"\n预计输出：{path}\n" + "\n".join(result.overwrites))
                if self.settings().fixed and action != "pdf":
                    text = f"插入首行；固定编号：{self.settings().fixed_text}\n" + text
            else:
                text = result.message
            item = QTableWidgetItem(text.replace("\n", " | ")[:180])
            item.setToolTip(text)
            self.table.setItem(row, 1, item)
        if then_execute and not cancelled:
            self.execute(action)

    def execute(self, action):
        if self.locked or self.worker is not None or self.closing():
            return
        needed = "pdf" if action == "pdf" else "number"
        if not self.previews or self.preview_action != needed:
            self.preview(action, then_execute=True)
            return
        valid = [p for p in self.previews if isinstance(p, NumberPreview)]
        if not valid:
            self.show_text("无法处理", "\n".join(p.message for p in self.previews))
            return
        changes = [f"{p.source.name}\n" + "\n".join(p.overwrites) for p in valid if p.overwrites]
        message = f"按当前预览处理 {len(valid)} 个工作簿，源文件不覆盖。"
        if action != "pdf" and self.settings().fixed:
            message += "\n固定模式将在每张工作表最上方插入一行。"
        if changes and action != "pdf":
            message += "\n将替换以下已有内容：\n" + "\n".join(changes)
        confirmation = QMessageBox(QMessageBox.Icon.Question, "确认执行", message[:1800],
                                  QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, self)
        confirmation.setDefaultButton(QMessageBox.StandardButton.No)
        if len(message) > 1800:
            confirmation.setDetailedText(message)
        if confirmation.exec() != QMessageBox.StandardButton.Yes:
            return
        settings = self.settings()
        directory = self.output_dir()
        invalid = tuple(p for p in self.previews if isinstance(p, FileResult))
        self.start_task(valid, lambda p, cancelled: process_excel(p, settings, directory, action, cancelled),
                        lambda results, cancelled: self.execution_finished(invalid + results, cancelled))

    def execution_finished(self, results, cancelled):
        self.previews = ()
        self.show_results(results, cancelled)
