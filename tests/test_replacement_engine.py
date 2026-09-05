import pytest

from word_text_tool.models import ReplacementRule
from word_text_tool.replacement_engine import RuleValidationError, replace_text, validate_rule


def test_replacement_is_non_cascading_and_not_duplicated():
    rules = (ReplacementRule("xxx公司", "yyy公司"), ReplacementRule("yyy公司", "zzz公司"))
    outcome = replace_text("甲方：xxx公司", rules)
    assert outcome.text == "甲方：yyy公司"
    assert outcome.count == 1


def test_leftmost_match_uses_longest_rule_at_same_position():
    rules = (ReplacementRule("项目", "A"), ReplacementRule("项目编号", "B"))
    assert replace_text("项目编号：01", rules).text == "B：01"


def test_soft_line_break_is_matched_exactly():
    rule = ReplacementRule("第一行\n第二行", "合并")
    assert replace_text("第一行\n第二行", (rule,)).text == "合并"


def test_empty_replacement_deletes_match():
    assert replace_text("甲删除乙", (ReplacementRule("删除", ""),)).text == "甲乙"


@pytest.mark.parametrize("before", ["", "   ", "\t\n"])
def test_rule_rejects_empty_or_whitespace_only_before(before):
    with pytest.raises(RuleValidationError, match="不能为空或纯空白"):
        validate_rule(before, "新值", ())


def test_rule_preserves_significant_surrounding_whitespace():
    rule = validate_rule(" 项目编号 ", " 新编号 ", ())
    assert rule.before == " 项目编号 "
    assert rule.after == " 新编号 "


def test_rule_rejects_duplicate_before_text():
    existing = (ReplacementRule("项目编号", "A"),)
    with pytest.raises(RuleValidationError, match="已存在"):
        validate_rule("项目编号", "B", existing)


def test_rule_rejects_text_longer_than_10000_characters():
    with pytest.raises(RuleValidationError, match="10000"):
        validate_rule("甲" * 10001, "乙", ())

