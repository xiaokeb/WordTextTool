from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMessageBox

from word_text_tool.models import BatchSummary, FileResult, FileStatus
from word_text_tool.ui.file_panel import FilePanel
from word_text_tool.ui.result_dialog import ResultDialog


def test_file_panel_adds_once_and_deletes_only_the_selected_row(
    qtbot, tmp_path, monkeypatch
):
    first = tmp_path / "第一份.docx"
    second = tmp_path / "第二份.docx"
    first.write_bytes(b"x")
    second.write_bytes(b"x")
    panel = FilePanel()
    qtbot.addWidget(panel)
    information = []
    monkeypatch.setattr(
        QMessageBox, "information", lambda *args: information.append(args)
    )

    panel.add_paths((first, first, second))
    assert panel.files() == (first.resolve(), second.resolve())
    assert len(panel.skipped_results()) == 1
    assert "重复" in panel.skipped_results()[0].message
    assert panel.model.headerData(0, Qt.Orientation.Horizontal) == "文件名"
    assert not (
        panel.model.flags(panel.model.index(0, 0)) & Qt.ItemFlag.ItemIsEditable
    )

    panel.delete_button.click()
    assert information
    assert panel.files() == (first.resolve(), second.resolve())

    panel.table.selectRow(0)
    panel.delete_button.click()
    assert panel.files() == (second.resolve(),)


def test_legacy_doc_shows_immediate_nonblocking_notice_and_clear_resets_it(
    qtbot, tmp_path, monkeypatch
):
    legacy = tmp_path / "旧文档.doc"
    legacy.write_bytes(b"legacy")
    panel = FilePanel()
    qtbot.addWidget(panel)
    # Guard against accidental modal UI regressions hanging this test.
    monkeypatch.setattr(QMessageBox, "information", lambda *args: None)
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: None)

    panel.add_paths((legacy,))

    assert panel.files() == ()
    assert panel.notice_label.isVisibleTo(panel)
    assert "旧文档.doc" in panel.notice_label.text()
    assert ".doc" in panel.skipped_results()[0].message
    assert panel.skipped_results()[0].source == legacy.resolve()

    panel.clear_files()
    assert panel.skipped_results() == ()
    assert panel.notice_label.text() == ""
    assert panel.path_line_edit.text() == "未选择路径"
    assert panel.count_label.text() == "已选择 0 个文件"


def test_path_status_tracks_one_and_multiple_directories_and_locking(
    qtbot, tmp_path
):
    first_dir = tmp_path / "目录一"
    second_dir = tmp_path / "目录二"
    first_dir.mkdir()
    second_dir.mkdir()
    first = first_dir / "one.docx"
    sibling = first_dir / "two.docx"
    other = second_dir / "three.docx"
    for path in (first, sibling, other):
        path.write_bytes(b"x")
    panel = FilePanel()
    qtbot.addWidget(panel)

    panel.add_paths((first, sibling))
    assert panel.path_line_edit.isReadOnly()
    assert panel.path_line_edit.text() == str(first_dir.resolve())
    assert panel.path_line_edit.toolTip() == str(first_dir.resolve())

    panel.add_paths((other,))
    assert panel.path_line_edit.text() == "路径：2 个不同目录"
    assert panel.path_line_edit.toolTip().splitlines() == [
        str(first_dir.resolve()),
        str(second_dir.resolve()),
    ]

    panel.set_locked(True)
    assert panel.table.isEnabled()
    assert all(
        not button.isEnabled()
        for button in (
            panel.browse_files_button,
            panel.browse_folder_button,
            panel.add_file_button,
            panel.delete_button,
        )
    )


def test_result_dialog_keeps_fixed_groups_counts_and_complete_read_only_details(
    qtbot
):
    output = Path("输出") / "成功.docx"
    results = (
        FileResult(Path("成功.docx"), FileStatus.SUCCESS, "修改完成", output, 3),
        FileResult(Path("未匹配.docx"), FileStatus.UNMATCHED, "未找到匹配内容"),
        FileResult(Path("已跳过.doc"), FileStatus.SKIPPED, "旧版格式不支持"),
        FileResult(Path("失败.docx"), FileStatus.FAILED, "文档损坏"),
    )
    dialog = ResultDialog(BatchSummary(results))
    qtbot.addWidget(dialog)

    assert [dialog.tabs.tabText(index) for index in range(4)] == [
        "处理成功 (1)",
        "未匹配 (1)",
        "已跳过 (1)",
        "处理失败 (1)",
    ]
    assert dialog.editors[FileStatus.SUCCESS].toPlainText() == (
        f"成功.docx | 修改完成 | {output}"
    )
    assert dialog.editors[FileStatus.UNMATCHED].toPlainText() == (
        "未匹配.docx | 未找到匹配内容"
    )
    assert dialog.editors[FileStatus.SKIPPED].toPlainText() == (
        "已跳过.doc | 旧版格式不支持"
    )
    assert dialog.editors[FileStatus.FAILED].toPlainText() == (
        "失败.docx | 文档损坏"
    )
    assert all(editor.isReadOnly() for editor in dialog.editors.values())
