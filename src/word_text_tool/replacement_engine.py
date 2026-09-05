from dataclasses import dataclass
from typing import Sequence

from docx.text.run import Run

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


def _locate_start(lengths: Sequence[int], offset: int) -> tuple[int, int]:
    cursor = 0
    for index, length in enumerate(lengths):
        if offset < cursor + length:
            return index, offset - cursor
        cursor += length
    return len(lengths) - 1, lengths[-1]


def _locate_end(lengths: Sequence[int], offset: int) -> tuple[int, int]:
    cursor = 0
    for index, length in enumerate(lengths):
        cursor += length
        if offset <= cursor:
            return index, offset - (cursor - length)
    return len(lengths) - 1, lengths[-1]


def replace_runs(runs: Sequence[Run], rules: Sequence[ReplacementRule]) -> int:
    if not runs:
        return 0

    original_texts = [run.text for run in runs]
    lengths = [len(text) for text in original_texts]
    matches = find_matches("".join(original_texts), rules)
    if not matches:
        return 0

    rewritten = original_texts.copy()
    for match in reversed(matches):
        start_index, start_offset = _locate_start(lengths, match.start)
        end_index, end_offset = _locate_end(lengths, match.end)
        if start_index == end_index:
            current = rewritten[start_index]
            rewritten[start_index] = (
                current[:start_offset] + match.replacement + current[end_offset:]
            )
            continue

        rewritten[start_index] = rewritten[start_index][:start_offset] + match.replacement
        for index in range(start_index + 1, end_index):
            rewritten[index] = ""
        rewritten[end_index] = rewritten[end_index][end_offset:]

    for run, old_text, new_text in zip(runs, original_texts, rewritten, strict=True):
        if old_text != new_text:
            run.text = new_text
    return len(matches)
