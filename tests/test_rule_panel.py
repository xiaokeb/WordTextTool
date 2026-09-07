import pytest
from PyQt6 import sip
from PyQt6.QtCore import QCoreApplication, QEvent, Qt, QTimer
from PyQt6.QtWidgets import QDialog, QMessageBox, QPlainTextEdit

from word_text_tool.ui.rule_panel import RulePanel


def _enter_rule(panel, before, after):
    panel.before_input.setPlainText(before)
    panel.after_input.setPlainText(after)
    panel.add_button.click()


def test_rule_table_is_two_column_read_only_and_keeps_full_10000_characters(qtbot):
    panel = RulePanel()
    qtbot.addWidget(panel)
    before = "甲" * 10_000

    _enter_rule(panel, before, "乙")

    model = panel.table.model()
    first = model.index(0, 0)
    assert model.headerData(0, Qt.Orientation.Horizontal) == "替换前"
    assert model.headerData(1, Qt.Orientation.Horizontal) == "替换后"
    assert not (model.flags(first) & Qt.ItemFlag.ItemIsEditable)
    assert panel.rules()[0].before == before
    assert model.data(first, Qt.ItemDataRole.ToolTipRole) == before
    assert len(model.data(first, Qt.ItemDataRole.DisplayRole)) < len(before)
    assert panel.before_input.toPlainText() == ""
    assert panel.after_input.toPlainText() == ""


def test_rule_details_show_both_complete_values_in_read_only_editors(qtbot):
    panel = RulePanel()
    qtbot.addWidget(panel)
    before = "查" * 10_000
    after = "换" * 10_000
    _enter_rule(panel, before, after)

    dialog = panel._create_details_dialog(panel.rules()[0])
    qtbot.addWidget(dialog)
    editors = dialog.findChildren(QPlainTextEdit)

    assert [editor.toPlainText() for editor in editors] == [before, after]
    assert all(editor.isReadOnly() for editor in editors)


def test_repeated_rule_details_release_closed_dialogs_and_editors(qtbot):
    panel = RulePanel()
    qtbot.addWidget(panel)
    _enter_rule(panel, "查" * 10_000, "换" * 10_000)
    opened = []
    for _ in range(3):
        def close_details():
            for dialog in panel.findChildren(QDialog):
                if dialog.isVisible():
                    opened.append(dialog)
                    dialog.reject()

        QTimer.singleShot(0, close_details)
        panel._show_details(panel.model.index(0, 0))
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        assert panel.findChildren(QDialog) == []
        assert all(sip.isdeleted(dialog) for dialog in opened)
        assert len(panel.findChildren(QPlainTextEdit)) == 2


def test_delete_requires_a_selected_row_then_removes_only_that_row(
    qtbot, monkeypatch
):
    panel = RulePanel()
    qtbot.addWidget(panel)
    information = []
    monkeypatch.setattr(
        QMessageBox, "information", lambda *args: information.append(args)
    )
    _enter_rule(panel, "第一条", "一")
    _enter_rule(panel, "第二条", "二")

    panel.delete_button.click()
    assert information
    assert [rule.before for rule in panel.rules()] == ["第一条", "第二条"]

    panel.table.selectRow(0)
    panel.delete_button.click()
    assert [rule.before for rule in panel.rules()] == ["第二条"]


@pytest.mark.parametrize(
    ("answer", "expected_count"),
    [
        (QMessageBox.StandardButton.No, 0),
        (QMessageBox.StandardButton.Yes, 1),
    ],
)
def test_empty_replacement_obeys_explicit_confirmation(
    qtbot, monkeypatch, answer, expected_count
):
    panel = RulePanel()
    qtbot.addWidget(panel)
    questions = []
    monkeypatch.setattr(
        QMessageBox,
        "question",
        lambda *args: questions.append(args) or answer,
    )

    _enter_rule(panel, "删除内容", "")

    assert len(questions) == 1
    assert len(panel.rules()) == expected_count


def test_duplicate_before_is_rejected_and_lock_disables_mutation_controls(
    qtbot, monkeypatch
):
    panel = RulePanel()
    qtbot.addWidget(panel)
    warnings = []
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: warnings.append(args))
    _enter_rule(panel, "旧值", "第一次")
    _enter_rule(panel, "旧值", "第二次")

    assert len(warnings) == 1
    assert panel.rules() == (panel.rules()[0],)
    assert panel.rules()[0].after == "第一次"

    panel.set_locked(True)
    assert panel.table.isEnabled()
    assert all(
        not widget.isEnabled()
        for widget in (
            panel.before_input,
            panel.after_input,
            panel.add_button,
            panel.delete_button,
        )
    )
