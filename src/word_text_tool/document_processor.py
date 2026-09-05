"""Replace text in supported DOCX stories without rewriting fields or rich runs."""

from collections.abc import Iterator, Sequence
from dataclasses import dataclass

from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from word_text_tool.models import ReplacementRule
from word_text_tool.replacement_engine import replace_runs


_RUN = qn("w:r")
_HYPERLINK = qn("w:hyperlink")
_FIELD_CHAR = qn("w:fldChar")
_FIELD_TYPE = qn("w:fldCharType")
_BREAK = qn("w:br")
_BREAK_TYPE = qn("w:type")
_SAFE_CHILDREN = {qn(name) for name in ("w:rPr", "w:t", "w:tab", "w:br", "w:cr")}
_STORY_NAMES = (
    "header", "footer", "first_page_header", "first_page_footer",
    "even_page_header", "even_page_footer",
)


@dataclass
class _FieldState:
    # Fields may span paragraphs, so state belongs to the story, not a paragraph.
    depth: int = 0

    def observe(self, run_element) -> bool:
        markers = list(run_element.iterchildren(tag=_FIELD_CHAR))
        for marker in markers:
            kind = marker.get(_FIELD_TYPE)
            if kind == "begin":
                self.depth += 1
            elif kind == "end":
                self.depth = max(0, self.depth - 1)
        return bool(markers)


def _iter_container_paragraphs(container, seen_cells: set) -> Iterator[Paragraph]:
    """Keep XML document order, including tables between field paragraphs."""
    for block in container.iter_inner_content():
        if isinstance(block, Paragraph):
            yield block
        else:
            for row in block.rows:
                for cell in row.cells:
                    # Keep elements alive in the set; Python id values can be reused.
                    if cell._tc in seen_cells:
                        continue
                    seen_cells.add(cell._tc)
                    yield from _iter_container_paragraphs(cell, seen_cells)


def _is_safe_text_run(element) -> bool:
    for child in element:
        if child.tag not in _SAFE_CHILDREN:
            return False
        if child.tag == _BREAK:
            # Run.text normalizes breaks. Page/column and clear=... carry layout
            # semantics that must survive, so the entire run forms a boundary.
            if child.get(_BREAK_TYPE) not in (None, "textWrapping"):
                return False
            if any(attribute != _BREAK_TYPE for attribute in child.attrib):
                return False
    return True


def _iter_run_groups(container, paragraph: Paragraph, fields: _FieldState):
    group: list[Run] = []
    for child in container:
        if child.tag == _RUN:
            has_markers = fields.observe(child)
            if not has_markers and fields.depth == 0 and _is_safe_text_run(child):
                group.append(Run(child, paragraph))
                continue

        if group:
            yield tuple(group)
            group.clear()
        if child.tag == _HYPERLINK:
            # Every link is an independent region. Share field state so code and
            # results stay protected even when markers occur inside hyperlinks.
            yield from _iter_run_groups(child, paragraph, fields)
        # fldSimple and all other unsupported paragraph children are boundaries.
        # Do not descend into drawings, text boxes, or automatic simple fields.
    if group:
        yield tuple(group)


def _replace_story(container, rules: Sequence[ReplacementRule]) -> int:
    fields = _FieldState()
    total = 0
    for paragraph in _iter_container_paragraphs(container, set()):
        for group in _iter_run_groups(paragraph._p, paragraph, fields):
            total += replace_runs(group, rules)
    return total


def replace_document(document: DocumentObject, rules: Sequence[ReplacementRule]) -> int:
    """Return match count for body and existing unique header/footer parts.

    Matching never crosses paragraphs, hyperlinks, automatic fields, or runs
    whose non-text XML would be lost by the replacement engine's text setter.
    This function only edits the loaded document; it does not save files.
    """
    total = _replace_story(document, rules)
    seen_parts: set[str] = set()
    for section in document.sections:
        for name in _STORY_NAMES:
            story = getattr(section, name)
            # An inherited story is handled by its defining section. Accessing
            # .part on an undefined story would create an unwanted empty part.
            if story.is_linked_to_previous:
                continue
            part_name = str(story.part.partname)
            if part_name in seen_parts:
                continue
            seen_parts.add(part_name)
            total += _replace_story(story, rules)
    return total
