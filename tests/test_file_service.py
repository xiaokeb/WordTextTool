import tempfile
from pathlib import Path

import pytest

from word_text_tool.file_service import (
    MAX_FILE_SIZE_BYTES,
    OutputDirectoryError,
    discover_files,
    next_output_path,
    prepare_output_directory,
)
from word_text_tool.models import FileStatus


def test_discover_filters_doc_duplicates_and_nested_files(tmp_path):
    current = tmp_path / "current.docx"
    legacy = tmp_path / "legacy.doc"
    nested_dir = tmp_path / "nested"
    nested_dir.mkdir()
    nested = nested_dir / "nested.docx"
    current.write_bytes(b"x")
    legacy.write_bytes(b"x")
    nested.write_bytes(b"x")

    result = discover_files((tmp_path, current), ())

    assert result.accepted == (current.resolve(),)
    assert {item.source.name for item in result.skipped} == {
        "legacy.doc",
        "current.docx",
    }
    assert all(item.status is FileStatus.SKIPPED for item in result.skipped)
    assert nested.resolve() not in result.accepted


def test_oversized_file_is_skipped_but_exact_limit_is_accepted(tmp_path):
    exact = tmp_path / "exact.docx"
    oversized = tmp_path / "oversized.docx"
    for path, size in (
        (exact, MAX_FILE_SIZE_BYTES),
        (oversized, MAX_FILE_SIZE_BYTES + 1),
    ):
        with path.open("wb") as stream:
            stream.seek(size - 1)
            stream.write(b"x")

    result = discover_files((exact, oversized), ())

    assert result.accepted == (exact.resolve(),)
    assert tuple(item.source for item in result.skipped) == (oversized.resolve(),)
    assert "50MB" in result.skipped[0].message


def test_missing_file_is_skipped(tmp_path):
    missing = tmp_path / "missing.docx"

    result = discover_files((missing,), ())

    assert result.accepted == ()
    assert result.skipped[0].source == missing.resolve()
    assert "不存在或不是普通文件" in result.skipped[0].message


def test_existing_paths_are_normalized_before_duplicate_detection(tmp_path):
    current = tmp_path / "current.docx"
    current.write_bytes(b"x")
    alternate_spelling = current.parent / "." / current.name

    result = discover_files((alternate_spelling,), (current.resolve(),))

    assert result.accepted == ()
    assert result.skipped[0].message == "重复文件已跳过"


def test_default_output_and_collision_numbering(tmp_path):
    source = tmp_path / "合同.docx"
    source.write_bytes(b"x")

    output_dir = prepare_output_directory(source, None)
    first = next_output_path(source, output_dir)

    assert output_dir == tmp_path / "修改版"
    assert first.name == "合同 - 修改版.docx"
    first.write_bytes(b"existing")
    assert next_output_path(source, output_dir).name == "合同 - 修改版 (1).docx"


def test_custom_output_directory_is_used_without_creating_a_subdirectory(tmp_path):
    output = tmp_path / "output"
    output.mkdir()

    result = prepare_output_directory(tmp_path / "source.docx", output)

    assert result == output.resolve()
    assert not (output / "修改版").exists()


def test_missing_custom_output_directory_is_rejected(tmp_path):
    with pytest.raises(OutputDirectoryError, match="自定义保存目录不存在"):
        prepare_output_directory(tmp_path / "source.docx", tmp_path / "missing")


def test_collision_counter_advances_to_two(tmp_path):
    source = tmp_path / "合同.docx"
    (tmp_path / "合同 - 修改版.docx").write_bytes(b"x")
    (tmp_path / "合同 - 修改版 (1).docx").write_bytes(b"x")

    assert next_output_path(source, tmp_path).name == "合同 - 修改版 (2).docx"


def test_unreadable_file_is_skipped_without_losing_other_files(tmp_path, monkeypatch):
    blocked = tmp_path / "blocked.docx"
    readable = tmp_path / "readable.docx"
    blocked.write_bytes(b"x")
    readable.write_bytes(b"x")
    original_open = Path.open

    def controlled_open(path, *args, **kwargs):
        if path == blocked:
            raise PermissionError("denied")
        return original_open(path, *args, **kwargs)

    monkeypatch.setattr(Path, "open", controlled_open)

    result = discover_files((blocked, readable), ())

    assert result.accepted == (readable.resolve(),)
    assert result.skipped[0].source == blocked.resolve()
    assert result.skipped[0].message == "文件不可读取"


def test_directory_iteration_error_keeps_files_yielded_and_later_candidates(
    tmp_path, monkeypatch
):
    troubled = tmp_path / "troubled"
    troubled.mkdir()
    yielded = troubled / "yielded.docx"
    later = tmp_path / "later.docx"
    yielded.write_bytes(b"x")
    later.write_bytes(b"x")
    original_iterdir = Path.iterdir

    def controlled_iterdir(path):
        if path != troubled:
            return original_iterdir(path)

        def fail_after_one_file():
            yield yielded
            raise OSError("scan failed")

        return fail_after_one_file()

    monkeypatch.setattr(Path, "iterdir", controlled_iterdir)

    result = discover_files((troubled, later), ())

    assert result.accepted == (yielded.resolve(), later.resolve())
    assert len(result.skipped) == 1
    assert result.skipped[0].source == troubled.resolve()
    assert result.skipped[0].status is FileStatus.SKIPPED
    assert result.skipped[0].message == "文件夹无法扫描，已跳过"


def test_stat_error_is_skipped_without_losing_other_files(tmp_path, monkeypatch):
    blocked = tmp_path / "blocked.docx"
    readable = tmp_path / "readable.docx"
    blocked.write_bytes(b"x")
    readable.write_bytes(b"x")
    original_stat = Path.stat

    def controlled_stat(path, *args, **kwargs):
        if path == blocked:
            raise OSError("stat failed")
        return original_stat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "stat", controlled_stat)

    result = discover_files((blocked, readable), ())

    assert result.accepted == (readable.resolve(),)
    assert result.skipped[0].source == blocked.resolve()
    assert result.skipped[0].message == "无法读取路径信息，已跳过"


def test_resolve_error_is_skipped_without_losing_other_files(tmp_path, monkeypatch):
    blocked = tmp_path / "blocked.docx"
    readable = tmp_path / "readable.docx"
    blocked.write_bytes(b"x")
    readable.write_bytes(b"x")
    original_resolve = Path.resolve

    def controlled_resolve(path, *args, **kwargs):
        if path == blocked:
            raise OSError("resolve failed")
        return original_resolve(path, *args, **kwargs)

    monkeypatch.setattr(Path, "resolve", controlled_resolve)

    result = discover_files((blocked, readable), ())

    assert result.accepted == (original_resolve(readable),)
    assert result.skipped[0].source == blocked
    assert result.skipped[0].message == "路径无法解析，已跳过"


def test_unwritable_custom_output_is_rejected(tmp_path, monkeypatch):
    output = tmp_path / "output"
    output.mkdir()

    def deny_temporary_file(**kwargs):
        raise PermissionError("denied")

    monkeypatch.setattr(tempfile, "NamedTemporaryFile", deny_temporary_file)

    with pytest.raises(OutputDirectoryError, match="保存目录不可写"):
        prepare_output_directory(tmp_path / "source.docx", output)


def test_custom_output_stat_error_is_reported(tmp_path, monkeypatch):
    output = tmp_path / "output"
    output.mkdir()
    original_stat = Path.stat

    def controlled_stat(path, *args, **kwargs):
        if path == output:
            raise OSError("stat failed")
        return original_stat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "stat", controlled_stat)

    with pytest.raises(OutputDirectoryError, match="无法检查自定义保存目录"):
        prepare_output_directory(tmp_path / "source.docx", output)


def test_default_output_source_resolve_error_is_reported(tmp_path, monkeypatch):
    source = tmp_path / "source.docx"
    original_resolve = Path.resolve

    def controlled_resolve(path, *args, **kwargs):
        if path == source:
            raise OSError("resolve failed")
        return original_resolve(path, *args, **kwargs)

    monkeypatch.setattr(Path, "resolve", controlled_resolve)

    with pytest.raises(OutputDirectoryError, match="无法解析源文件路径"):
        prepare_output_directory(source, None)


def test_output_collision_lstat_error_is_reported(tmp_path, monkeypatch):
    source = tmp_path / "source.docx"
    candidate = tmp_path / "source - 修改版.docx"
    original_lstat = Path.lstat

    def controlled_lstat(path, *args, **kwargs):
        if path == candidate:
            raise OSError("stat failed")
        return original_lstat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "lstat", controlled_lstat)

    with pytest.raises(OutputDirectoryError, match="无法检查输出文件是否存在"):
        next_output_path(source, tmp_path)


def test_broken_link_placeholder_is_treated_as_an_output_collision(
    tmp_path, monkeypatch
):
    source = tmp_path / "source.docx"
    occupied = tmp_path / "source - 修改版.docx"
    available = tmp_path / "source - 修改版 (1).docx"
    original_lstat = Path.lstat

    def controlled_lstat(path, *args, **kwargs):
        if path == occupied:
            return original_lstat(tmp_path, *args, **kwargs)
        if path == available:
            raise FileNotFoundError
        return original_lstat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "lstat", controlled_lstat)

    assert next_output_path(source, tmp_path) == available
