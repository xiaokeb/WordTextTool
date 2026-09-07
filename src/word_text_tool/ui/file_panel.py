"""DOCX 文件选择、去重、跳过提示与列表管理。"""

from pathlib import Path

from PyQt6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from word_text_tool.file_service import discover_files
from word_text_tool.models import FileResult


class FileTableModel(QAbstractTableModel):
    """单列只读模型；文件名用于显示，规范路径用于提示。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._files: tuple[Path, ...] = ()

    def rowCount(self, parent=QModelIndex()) -> int:  # noqa: N802 - Qt API
        return 0 if parent.isValid() else len(self._files)

    def columnCount(self, parent=QModelIndex()) -> int:  # noqa: N802 - Qt API
        return 0 if parent.isValid() else 1

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or not 0 <= index.row() < len(self._files):
            return None
        path = self._files[index.row()]
        if role == Qt.ItemDataRole.DisplayRole:
            return path.name
        if role == Qt.ItemDataRole.ToolTipRole:
            return str(path)
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
            and section == 0
        ):
            return "文件名"
        return super().headerData(section, orientation, role)

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable

    def replace_files(self, files: tuple[Path, ...]) -> None:
        self.beginResetModel()
        self._files = files
        self.endResetModel()


class FilePanel(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._files: list[Path] = []
        self._skipped: list[FileResult] = []

        self.path_line_edit = QLineEdit(self)
        self.path_line_edit.setReadOnly(True)
        self.path_line_edit.setText("未选择路径")
        self.path_line_edit.setToolTip("未选择路径")

        self.browse_files_button = QPushButton("浏览文件", self)
        self.browse_folder_button = QPushButton("浏览文件夹", self)
        path_row = QHBoxLayout()
        path_row.addWidget(self.path_line_edit, 1)
        path_row.addWidget(self.browse_files_button)
        path_row.addWidget(self.browse_folder_button)

        self.count_label = QLabel("已选择 0 个文件", self)
        self.notice_label = QLabel("", self)
        self.notice_label.setWordWrap(True)
        self.notice_label.setStyleSheet("color: #8a5a00;")
        self.notice_label.hide()

        self.model = FileTableModel(self)
        self.table = QTableView(self)
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setMinimumHeight(140)

        self.add_file_button = QPushButton("添加文件", self)
        self.delete_button = QPushButton("删除文件", self)
        button_row = QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(self.add_file_button)
        button_row.addWidget(self.delete_button)
        button_row.addStretch(1)

        layout = QVBoxLayout(self)
        layout.addLayout(path_row)
        layout.addWidget(self.count_label)
        layout.addWidget(self.notice_label)
        layout.addWidget(self.table, 1)
        layout.addLayout(button_row)

        self.browse_files_button.clicked.connect(self._browse_files)
        self.browse_folder_button.clicked.connect(self._browse_folder)
        self.add_file_button.clicked.connect(self._browse_files)
        self.delete_button.clicked.connect(self._delete_selected)

    def files(self) -> tuple[Path, ...]:
        return tuple(self._files)

    def skipped_results(self) -> tuple[FileResult, ...]:
        return tuple(self._skipped)

    def add_paths(self, paths) -> None:
        discovery = discover_files((Path(path) for path in paths), self._files)
        self._files.extend(discovery.accepted)
        self._skipped.extend(discovery.skipped)
        self._refresh_model_and_status()
        if discovery.skipped:
            reasons = "；".join(
                f"{result.source.name}：{result.message}"
                for result in discovery.skipped[:3]
            )
            remaining = len(discovery.skipped) - 3
            suffix = f"；另有 {remaining} 项" if remaining > 0 else ""
            self.notice_label.setText(
                f"本次已跳过 {len(discovery.skipped)} 项。{reasons}{suffix}"
            )
            self.notice_label.show()

    def clear_files(self) -> None:
        self._files.clear()
        self._skipped.clear()
        self.notice_label.clear()
        self.notice_label.hide()
        self._refresh_model_and_status()

    def set_locked(self, locked: bool) -> None:
        for button in (
            self.browse_files_button,
            self.browse_folder_button,
            self.add_file_button,
            self.delete_button,
        ):
            button.setEnabled(not locked)

    def _browse_files(self) -> None:
        filenames, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 Word 文件",
            "",
            "Word 文件 (*.docx *.doc)",
        )
        if filenames:
            self.add_paths(filenames)

    def _browse_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if directory:
            self.add_paths((directory,))

    def _delete_selected(self) -> None:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "请选择文件", "请先选中一个文件")
            return
        row = rows[0].row()
        if 0 <= row < len(self._files):
            del self._files[row]
            self._refresh_model_and_status()

    def _refresh_model_and_status(self) -> None:
        files = tuple(self._files)
        self.model.replace_files(files)
        self.count_label.setText(f"已选择 {len(files)} 个文件")

        if not files:
            status = "未选择路径"
            tooltip = status
        else:
            parents = tuple(dict.fromkeys(path.parent for path in files))
            if len(parents) == 1:
                status = str(parents[0])
            else:
                status = f"路径：{len(parents)} 个不同目录"
            tooltip = "\n".join(str(parent) for parent in parents)

        self.path_line_edit.setText(status)
        self.path_line_edit.setToolTip(tooltip)
