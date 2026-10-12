import pytest
from word_text_tool.rename_service import preview_renames, execute_rename, valid_name
from word_text_tool.models import FileStatus
from word_text_tool.safe_output import fingerprint, OperationError


def test_conflicts_preview_and_bytes_preserved(tmp_path):
    source = tmp_path / "报告 - 修改版.docx"
    source.write_bytes(b"original")
    (tmp_path / "报告.docx").write_bytes(b"existing")
    before = fingerprint(source)
    preview, = preview_renames([source])
    assert preview.target.name == "报告 (1).docx"
    result = execute_rename(preview)
    assert result.status == FileStatus.SUCCESS
    assert fingerprint(result.output).digest == before.digest
    assert (tmp_path / "报告.docx").read_bytes() == b"existing"


def test_batch_targets_reserved_and_existing_source_never_overwritten(tmp_path):
    files = [tmp_path / "A - 修改版.txt", tmp_path / "A - 修改版 - 修改版.txt", tmp_path / "A.txt"]
    for index, path in enumerate(files):
        path.write_text(str(index))
    previews = preview_renames(files, all_matches=True)
    assert len({p.target for p in previews}) == 3
    for p in previews:
        assert execute_rename(p).status in (FileStatus.SUCCESS, FileStatus.UNMATCHED)
    assert files[-1].read_text() == "2"


def test_external_target_race_blocks_rename(tmp_path):
    source = tmp_path / "A - 修改版.txt"
    source.write_text("source")
    preview, = preview_renames([source])
    preview.target.write_text("external")
    result = execute_rename(preview)
    assert result.status == FileStatus.FAILED
    assert source.read_text() == "source"
    assert preview.target.read_text() == "external"


def test_changed_source_blocks_rename(tmp_path):
    source = tmp_path / "A - 修改版.txt"
    source.write_text("source")
    preview, = preview_renames([source])
    source.write_text("changed")
    assert execute_rename(preview).status == FileStatus.FAILED
    assert source.exists()


@pytest.mark.parametrize("name", ["", "CON.txt", "LPT1.pdf", "COM¹.txt", "A.", "A ", "../x", "A:b", "a" * 256])
def test_invalid_windows_names(name):
    with pytest.raises(OperationError):
        valid_name(name)


def test_exact_suffix_and_extension_preserved(tmp_path):
    source = tmp_path / "前 - 修改版中 - 修改版.docx"
    source.write_text("x")
    last, = preview_renames([source])
    all_, = preview_renames([source], all_matches=True)
    assert last.target.name == "前 - 修改版中.docx"
    assert all_.target.name == "前中.docx"
