"""Preview-first renaming, preserving bytes and protecting every source path."""
import os
from dataclasses import dataclass
from pathlib import Path

from word_text_tool.models import FileResult, FileStatus
from word_text_tool.safe_output import Fingerprint, OperationError, available_path, error_message, fingerprint, require_unchanged


@dataclass(frozen=True)
class RenamePreview:
    source: Path
    target: Path | None
    fingerprint: Fingerprint | None
    message: str = ""


def valid_name(name):
    stem = name.split(".")[0].upper()
    reserved = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}
    reserved |= {f"{prefix}{n}" for prefix in ("COM", "LPT") for n in list("123456789") + list("¹²³")}
    if (not name or name in (".", "..") or name[-1] in " ." or stem in reserved
            or any(c in '<>:"/\\|?*' or ord(c) < 32 for c in name)
            or len(name.encode("utf-16-le")) // 2 > 255):
        raise OperationError("新文件名为空、非法或过长")


def preview_renames(files, suffix=" - 修改版", all_matches=False):
    if not suffix:
        raise OperationError("待删除内容不能为空")
    reserved = [Path(p).absolute() for p in files]
    result = []
    for source in reserved.copy():
        try:
            snapshot = fingerprint(source)
            stem = source.stem
            renamed = stem.replace(suffix, "") if all_matches else stem.removesuffix(suffix)
            if stem == renamed:
                result.append(RenamePreview(source, source, snapshot, "无匹配，名称不变"))
                continue
            if not renamed or renamed[-1] in " .":
                raise OperationError("主文件名为空或结尾为空格/点")
            valid_name(renamed + source.suffix)
            target = available_path(source.parent, renamed, source.suffix, reserved)
            valid_name(target.name)
            reserved.append(target)
            result.append(RenamePreview(source, target, snapshot, "可重命名"))
        except Exception as exc:
            result.append(RenamePreview(source, None, None, error_message(exc)))
    return tuple(result)


def execute_rename(preview):
    try:
        if preview.target is None:
            return FileResult(preview.source, FileStatus.SKIPPED, preview.message)
        require_unchanged(preview.source, preview.fingerprint)
        if preview.target == preview.source:
            return FileResult(preview.source, FileStatus.UNMATCHED, "无匹配，名称不变")
        if os.name != "nt" or preview.target.parent != preview.source.parent:
            raise OperationError("重命名仅支持 Windows 同目录操作")
        if os.path.lexists(preview.target):
            raise OperationError("预览目标已被占用，请重新预览")
        os.rename(preview.source, preview.target)
        return FileResult(preview.source, FileStatus.SUCCESS, "重命名完成，文件内容未修改", preview.target)
    except Exception as exc:
        return FileResult(preview.source, FileStatus.FAILED, error_message(exc))
