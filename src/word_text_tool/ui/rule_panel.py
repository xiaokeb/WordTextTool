"""只读替换规则列表、输入校验与完整内容查看。"""

from PyQt6.QtCore import QAbstractTableModel, QModelIndex, QSize, Qt, pyqtSignal
from PyQt6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStyledItemDelegate,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from word_text_tool.models import ReplacementRule
from word_text_tool.replacement_engine import RuleValidationError, validate_rule


_PREVIEW_LIMIT = 500


def _preview_text(value: str) -> str:
    if len(value) <= _PREVIEW_LIMIT:
        return value
    return f"{value[:_PREVIEW_LIMIT]}…（双击查看完整内容）"


class RuleTableModel(QAbstractTableModel):
    """以完整字符串保存规则，显示预览不参与替换输入。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._rules: list[ReplacementRule] = []

    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802 - Qt API
        return 0 if parent.isValid() else len(self._rules)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802 - Qt API
        return 0 if parent.isValid() else 2

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or not 0 <= index.row() < len(self._rules):
            return None
        rule = self._rules[index.row()]
        value = rule.before if index.column() == 0 else rule.after
        if role == Qt.ItemDataRole.DisplayRole:
            return _preview_text(value)
        if role == Qt.ItemDataRole.ToolTipRole:
            return value
        return None

    def headerData(  # noqa: N802 - Qt API
        self,
        section,
        orientation,
        role=Qt.ItemDataRole.DisplayRole,
    ):
        if (
            role == Qt.ItemDataRole.DisplayRole
            and orientation == Qt.Orientation.Horizontal
            and 0 <= section < 2
        ):
            return ("替换前", "替换后")[section]
        return super().headerData(section, orientation, role)

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable

    def rules(self) -> tuple[ReplacementRule, ...]:
        return tuple(self._rules)

    def append_rule(self, rule: ReplacementRule) -> None:
        row = len(self._rules)
        self.beginInsertRows(QModelIndex(), row, row)
        self._rules.append(rule)
        self.endInsertRows()

    def remove_rule(self, row: int) -> None:
        if not 0 <= row < len(self._rules):
            return
        self.beginRemoveRows(QModelIndex(), row, row)
        del self._rules[row]
        self.endRemoveRows()

    def clear(self) -> None:
        self.beginResetModel()
        self._rules.clear()
        self.endResetModel()


class CappedRowDelegate(QStyledItemDelegate):
    """允许多行预览，同时避免单条规则撑开整张表。"""

    def sizeHint(self, option, index):  # noqa: N802 - Qt API
        base = super().sizeHint(option, index)
        return QSize(base.width(), min(base.height(), 120))


class RulePanel(QWidget):
    rules_changed = pyqtSignal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self.model = RuleTableModel(self)
        self.table = QTableView(self)
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.table.setWordWrap(True)
        self.table.setItemDelegate(CappedRowDelegate(self.table))
        self.table.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setMinimumHeight(180)

        self.before_input = QPlainTextEdit(self)
        self.before_input.setPlaceholderText("输入要查找的完整文本")
        self.after_input = QPlainTextEdit(self)
        self.after_input.setPlaceholderText("输入替换后文本；留空表示删除")
        self.before_input.setMinimumHeight(100)
        self.after_input.setMinimumHeight(100)

        input_row = QHBoxLayout()
        for label_text, editor in (
            ("替换前：", self.before_input),
            ("替换后：", self.after_input),
        ):
            column = QVBoxLayout()
            column.addWidget(QLabel(label_text, self))
            column.addWidget(editor)
            input_row.addLayout(column, 1)

        self.add_button = QPushButton("添加规则", self)
        self.delete_button = QPushButton("删除规则", self)
        button_row = QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(self.add_button)
        button_row.addWidget(self.delete_button)
        button_row.addStretch(1)

        layout = QVBoxLayout(self)
        layout.addWidget(self.table, 3)
        layout.addLayout(input_row, 2)
        layout.addLayout(button_row)

        self.add_button.clicked.connect(self._add_rule)
        self.delete_button.clicked.connect(self._delete_selected)
        self.table.doubleClicked.connect(self._show_details)

    def rules(self) -> tuple[ReplacementRule, ...]:
        return self.model.rules()

    def clear_rules(self) -> None:
        self.model.clear()
        self.before_input.clear()
        self.after_input.clear()
        self.rules_changed.emit()

    def set_locked(self, locked: bool) -> None:
        for widget in (
            self.before_input,
            self.after_input,
            self.add_button,
            self.delete_button,
        ):
            widget.setEnabled(not locked)

    def _add_rule(self) -> None:
        before = self.before_input.toPlainText()
        after = self.after_input.toPlainText()
        try:
            rule = validate_rule(before, after, self.rules())
        except RuleValidationError as exc:
            QMessageBox.warning(self, "规则无效", str(exc))
            return

        if after == "" and QMessageBox.question(
            self,
            "确认删除内容",
            "替换后为空将删除所有匹配内容，是否继续？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) != QMessageBox.StandardButton.Yes:
            return

        self.model.append_rule(rule)
        self.before_input.clear()
        self.after_input.clear()
        self.before_input.setFocus()
        self.rules_changed.emit()

    def _delete_selected(self) -> None:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "请选择规则", "请先选中一条规则")
            return
        self.model.remove_rule(rows[0].row())
        self.rules_changed.emit()

    def _create_details_dialog(self, rule: ReplacementRule) -> QDialog:
        dialog = QDialog(self)
        dialog.setWindowTitle("规则完整内容")
        layout = QVBoxLayout(dialog)
        for label_text, value in (
            ("替换前：", rule.before),
            ("替换后：", rule.after),
        ):
            layout.addWidget(QLabel(label_text, dialog))
            editor = QPlainTextEdit(value, dialog)
            editor.setReadOnly(True)
            layout.addWidget(editor, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, dialog)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.resize(700, 500)
        return dialog

    def _show_details(self, index: QModelIndex) -> None:
        if index.isValid() and 0 <= index.row() < len(self.rules()):
            dialog = self._create_details_dialog(self.rules()[index.row()])
            try:
                dialog.exec()
            finally:
                dialog.deleteLater()
