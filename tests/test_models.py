from pathlib import Path

from word_text_tool.models import BatchSummary, FileResult, FileStatus, ReplacementRule


def test_batch_summary_counts_each_status():
    results = (
        FileResult(Path("a.docx"), FileStatus.SUCCESS, "完成", Path("out.docx"), 2),
        FileResult(Path("b.docx"), FileStatus.UNMATCHED, "未匹配"),
        FileResult(Path("c.doc"), FileStatus.SKIPPED, "暂不支持 .doc"),
        FileResult(Path("d.docx"), FileStatus.FAILED, "文档损坏"),
    )
    summary = BatchSummary(results=results, cancelled=True)
    assert summary.count(FileStatus.SUCCESS) == 1
    assert summary.count(FileStatus.UNMATCHED) == 1
    assert summary.count(FileStatus.SKIPPED) == 1
    assert summary.count(FileStatus.FAILED) == 1
    assert summary.cancelled is True


def test_replacement_rule_is_immutable():
    rule = ReplacementRule("旧编号", "新编号")
    assert (rule.before, rule.after) == ("旧编号", "新编号")
