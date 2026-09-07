"""Replace text in supported DOCX stories without rewriting fields or rich runs."""

import os
import tempfile
from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile

from docx import Document
from docx.document import Document as DocumentObject
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from word_text_tool.file_service import (
    OutputDirectoryError,
    next_output_path,
    prepare_output_directory,
)
from word_text_tool.models import FileResult, FileStatus, ReplacementRule
from word_text_tool.replacement_engine import replace_runs


_RUN = qn("w:r")
_HYPERLINK = qn("w:hyperlink")
_FIELD_CHAR = qn("w:fldChar")
_FIELD_TYPE = qn("w:fldCharType")
_BREAK = qn("w:br")
_BREAK_TYPE = qn("w:type")
_SAFE_CHILDREN = {qn(name) for name in ("w:rPr", "w:t", "w:tab", "w:br", "w:cr")}
_FIELD_WRAPPERS = {
    qn(name) for name in (
        "w:hyperlink", "w:ins", "w:del", "w:moveFrom", "w:moveTo",
        "w:smartTag", "w:customXml", "w:sdt", "w:sdtContent", "w:dir", "w:bdo",
    )
}
_FIELD_BLOCKS = {qn(name) for name in ("w:p", "w:tbl", "w:tr", "w:tc")}
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


def _observe_wrapped_fields(element, fields: _FieldState) -> None:
    """Track wrapped field markers without making wrapped text editable."""
    for child in element:
        if child.tag == _RUN:
            fields.observe(child)
        elif child.tag in _FIELD_WRAPPERS or child.tag in _FIELD_BLOCKS:
            _observe_wrapped_fields(child, fields)


def _iter_container_paragraphs(
    container, seen_cells: set, fields: _FieldState
) -> Iterator[Paragraph]:
    """Keep XML document order, including tables between field paragraphs."""
    element = container._element
    if isinstance(container, DocumentObject):
        element = element.body
    for child in element:
        if child.tag == qn("w:p"):
            yield Paragraph(child, container)
        elif child.tag == qn("w:tbl"):
            block = Table(child, container)
            for row in block.rows:
                for cell in row.cells:
                    # Keep elements alive in the set; Python id values can be reused.
                    if cell._tc in seen_cells:
                        continue
                    seen_cells.add(cell._tc)
                    yield from _iter_container_paragraphs(cell, seen_cells, fields)
        elif child.tag in _FIELD_WRAPPERS:
            # Block content controls/revisions are protected, but their markers
            # still participate in this story. Never descend into drawings or
            # textboxes: those carry a separate field state.
            _observe_wrapped_fields(child, fields)


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
        elif child.tag in _FIELD_WRAPPERS:
            # Revision, smart-tag, and content-control text stays unsupported,
            # but field markers inside those wrappers still affect later runs.
            _observe_wrapped_fields(child, fields)
        # fldSimple and all unsupported paragraph children remain boundaries.
    if group:
        yield tuple(group)


def _replace_story(container, rules: Sequence[ReplacementRule]) -> int:
    fields = _FieldState()
    total = 0
    for paragraph in _iter_container_paragraphs(container, set(), fields):
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


class DocumentProcessor:
    """Process one DOCX and publish a verified copy without changing its source."""

    def process(
        self,
        source: Path,
        rules: Sequence[ReplacementRule],
        custom_output_dir: Path | None,
    ) -> FileResult:
        original_source = Path(source)
        result_source = original_source
        temporary_path: Path | None = None
        try:
            result_source = original_source.resolve(strict=False)
            document = Document(result_source)
            replacement_count = replace_document(document, rules)
            if replacement_count == 0:
                return FileResult(
                    result_source,
                    FileStatus.UNMATCHED,
                    "未找到匹配内容",
                )

            output_dir = prepare_output_directory(
                result_source, custom_output_dir
            )
            with tempfile.NamedTemporaryFile(
                dir=output_dir,
                prefix=".word-text-tool-",
                suffix=".docx",
                delete=False,
            ) as temporary:
                temporary_path = Path(temporary.name)

            document.save(temporary_path)
            Document(temporary_path)

            while True:
                output_path = next_output_path(result_source, output_dir)
                try:
                    os.rename(temporary_path, output_path)
                except FileExistsError:
                    continue
                temporary_path = None
                return FileResult(
                    result_source,
                    FileStatus.SUCCESS,
                    "修改完成，已保存",
                    output_path,
                    replacement_count,
                )
        except (PackageNotFoundError, BadZipFile):
            return FileResult(
                result_source,
                FileStatus.FAILED,
                "文档损坏、已加密或不是有效的 DOCX 文件",
            )
        except OutputDirectoryError as exc:
            return FileResult(result_source, FileStatus.FAILED, str(exc))
        except PermissionError:
            return FileResult(
                result_source,
                FileStatus.FAILED,
                "文件被占用或没有访问权限",
            )
        except OSError:
            return FileResult(
                result_source,
                FileStatus.FAILED,
                "文件处理失败，请检查文件占用状态和保存目录",
            )
        except Exception:
            return FileResult(
                result_source,
                FileStatus.FAILED,
                "文档处理失败，请确认文件有效且未加密",
            )
        finally:
            if temporary_path is not None:
                try:
                    temporary_path.unlink(missing_ok=True)
                except OSError:
                    # Cleanup is best-effort and must not hide the file result.
                    pass
