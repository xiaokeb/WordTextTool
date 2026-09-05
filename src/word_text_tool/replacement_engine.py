from dataclasses import dataclass
from typing import Sequence

from word_text_tool.models import ReplacementRule


MAX_RULE_TEXT_LENGTH = 10_000


class RuleValidationError(ValueError):
    """用户输入的替换规则无效。"""


@dataclass(frozen=True, slots=True)
class TextMatch:
    start: int
    end: int
    replacement: str


@dataclass(frozen=True, slots=True)
class ReplacementOutcome:
    text: str
    count: int


def validate_rule(
    before: str,
    after: str,
    existing_rules: Sequence[ReplacementRule],
) -> ReplacementRule:
    if not before or before.isspace():
        raise RuleValidationError("替换前不能为空或纯空白")
    if len(before) > MAX_RULE_TEXT_LENGTH or len(after) > MAX_RULE_TEXT_LENGTH:
        raise RuleValidationError("替换前和替换后均不能超过 10000 个字符")
    if any(rule.before == before for rule in existing_rules):
        raise RuleValidationError("相同的替换前内容已存在")
    return ReplacementRule(before=before, after=after)


def find_matches(text: str, rules: Sequence[ReplacementRule]) -> tuple[TextMatch, ...]:
    ordered = sorted(
        ((index, rule) for index, rule in enumerate(rules) if rule.before),
        key=lambda item: (-len(item[1].before), item[0]),
    )
    matches: list[TextMatch] = []
    position = 0
    while position < len(text):
        selected = next(
            (rule for _, rule in ordered if text.startswith(rule.before, position)),
            None,
        )
        if selected is None:
            position += 1
            continue
        end = position + len(selected.before)
        matches.append(TextMatch(position, end, selected.after))
        position = end
    return tuple(matches)


def replace_text(text: str, rules: Sequence[ReplacementRule]) -> ReplacementOutcome:
    matches = find_matches(text, rules)
    if not matches:
        return ReplacementOutcome(text=text, count=0)

    parts: list[str] = []
    cursor = 0
    for match in matches:
        parts.append(text[cursor : match.start])
        parts.append(match.replacement)
        cursor = match.end
    parts.append(text[cursor:])
    return ReplacementOutcome(text="".join(parts), count=len(matches))
