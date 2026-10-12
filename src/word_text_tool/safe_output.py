"""Bounded source snapshots and non-overwriting Windows output publication."""
import hashlib
import os
import stat
from dataclasses import dataclass
from pathlib import Path

from word_text_tool.file_service import MAX_FILE_SIZE_BYTES


class OperationError(ValueError):
    """A user-facing validation or processing error."""


@dataclass(frozen=True)
class Fingerprint:
    size: int
    modified: int
    device: int
    inode: int
    digest: str


def fingerprint(path: Path) -> Fingerprint:
    path = Path(path)
    if path.is_symlink():
        raise OperationError("不处理符号链接，请选择实际文件")
    with path.open("rb") as stream:
        before = os.fstat(stream.fileno())
        if not stat.S_ISREG(before.st_mode) or before.st_size > MAX_FILE_SIZE_BYTES:
            raise OperationError("文件不是普通文件或超过 50 MiB")
        digest = hashlib.sha256()
        total = 0
        while block := stream.read(1024 * 1024):
            total += len(block)
            if total > MAX_FILE_SIZE_BYTES:
                raise OperationError("读取期间文件超过 50 MiB")
            digest.update(block)
        after = os.fstat(stream.fileno())
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise OperationError("文件在读取期间变化，请重新预览")
    return Fingerprint(after.st_size, after.st_mtime_ns, after.st_dev, after.st_ino, digest.hexdigest())


def require_unchanged(path: Path, expected: Fingerprint) -> None:
    if fingerprint(path) != expected:
        raise OperationError("文件自预览后已变化，请重新预览")


def available_path(directory: Path, stem: str, suffix: str, reserved=()) -> Path:
    occupied = {os.path.normcase(str(p)) for p in reserved}
    for index in range(10000):
        extra = f" ({index})" if index else ""
        path = directory / f"{stem}{extra}{suffix}"
        if not os.path.lexists(path) and os.path.normcase(str(path)) not in occupied:
            return path
    raise OperationError("同名文件过多，无法分配安全输出名称")


def publish(temporary: Path, directory: Path, stem: str, suffix: str) -> Path:
    if os.name != "nt":
        raise OperationError("安全文件发布仅支持 Windows")
    for _ in range(10000):
        destination = available_path(directory, stem, suffix)
        try:
            os.rename(temporary, destination)
            return destination
        except FileExistsError:
            continue
    raise OperationError("输出名称持续被占用，请重试")


def error_message(exc: Exception) -> str:
    if isinstance(exc, OperationError):
        return str(exc)
    if isinstance(exc, PermissionError):
        return "文件被占用或没有访问权限"
    if isinstance(exc, FileNotFoundError):
        return "文件或目录不存在"
    return "操作失败，请检查文件有效性、占用状态和保存权限"
