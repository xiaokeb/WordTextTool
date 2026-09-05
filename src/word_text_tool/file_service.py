import os
import stat
import tempfile
from collections.abc import Iterable
from pathlib import Path

from word_text_tool.models import DiscoveryResult, FileResult, FileStatus


MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024
_SUPPORTED_DISCOVERY_SUFFIXES = {".doc", ".docx"}


class OutputDirectoryError(OSError):
    """输出目录无法安全使用。"""


def _path_key(resolved: Path) -> str:
    return os.path.normcase(str(resolved))


def _skipped(path: Path, message: str) -> FileResult:
    return FileResult(path, FileStatus.SKIPPED, message)


def _resolve_for_discovery(path: Path, message: str) -> tuple[Path | None, FileResult | None]:
    try:
        return path.resolve(strict=False), None
    except OSError:
        return None, _skipped(path, message)


def _scan_directory(directory: Path, skipped: list[FileResult]) -> list[Path]:
    files: list[Path] = []
    try:
        iterator = directory.iterdir()
        while True:
            try:
                child = next(iterator)
            except StopIteration:
                break
            except OSError:
                skipped.append(_skipped(directory, "文件夹无法扫描，已跳过"))
                break

            if child.suffix.lower() not in _SUPPORTED_DISCOVERY_SUFFIXES:
                continue
            resolved, error = _resolve_for_discovery(child, "文件路径无法解析，已跳过")
            if error is not None:
                skipped.append(error)
                continue
            try:
                child_stat = resolved.stat()
            except FileNotFoundError:
                skipped.append(_skipped(resolved, "文件不存在或不是普通文件"))
                continue
            except OSError:
                skipped.append(_skipped(resolved, "无法读取文件信息，已跳过"))
                continue
            if stat.S_ISREG(child_stat.st_mode):
                files.append(resolved)
    except OSError:
        skipped.append(_skipped(directory, "文件夹无法扫描，已跳过"))
    return sorted(files, key=lambda child: child.name.casefold())


def _inspect_file(
    path: Path,
    seen: set[str],
    accepted: list[Path],
    skipped: list[FileResult],
) -> None:
    resolved, error = _resolve_for_discovery(path, "文件路径无法解析，已跳过")
    if error is not None:
        skipped.append(error)
        return

    try:
        file_stat = resolved.stat()
    except FileNotFoundError:
        skipped.append(_skipped(resolved, "文件不存在或不是普通文件"))
        return
    except OSError:
        skipped.append(_skipped(resolved, "无法读取文件信息，已跳过"))
        return

    if not stat.S_ISREG(file_stat.st_mode):
        skipped.append(_skipped(resolved, "文件不存在或不是普通文件"))
        return
    suffix = resolved.suffix.lower()
    if suffix == ".doc":
        skipped.append(_skipped(resolved, "旧版 .doc 暂不支持，将在后续版本处理"))
        return
    if suffix != ".docx":
        skipped.append(_skipped(resolved, "不支持的文件格式"))
        return

    key = _path_key(resolved)
    if key in seen:
        skipped.append(_skipped(resolved, "重复文件已跳过"))
        return
    if file_stat.st_size > MAX_FILE_SIZE_BYTES:
        skipped.append(_skipped(resolved, "文件超过 50MB，已跳过"))
        return
    try:
        with resolved.open("rb") as stream:
            stream.read(1)
    except OSError:
        skipped.append(_skipped(resolved, "文件不可读取"))
        return

    seen.add(key)
    accepted.append(resolved)


def discover_files(
    candidates: Iterable[Path], existing: Iterable[Path] = ()
) -> DiscoveryResult:
    accepted: list[Path] = []
    skipped: list[FileResult] = []
    seen: set[str] = set()
    for existing_path in existing:
        path = Path(existing_path)
        resolved, error = _resolve_for_discovery(
            path, "已有文件路径无法解析，已跳过"
        )
        if error is not None:
            skipped.append(error)
        else:
            seen.add(_path_key(resolved))

    expanded: list[Path] = []
    for candidate in candidates:
        path = Path(candidate)
        resolved, error = _resolve_for_discovery(path, "路径无法解析，已跳过")
        if error is not None:
            skipped.append(error)
            continue
        try:
            candidate_stat = resolved.stat()
        except FileNotFoundError:
            expanded.append(resolved)
            continue
        except OSError:
            skipped.append(_skipped(resolved, "无法读取路径信息，已跳过"))
            continue
        if stat.S_ISDIR(candidate_stat.st_mode):
            expanded.extend(_scan_directory(resolved, skipped))
        else:
            expanded.append(resolved)

    for path in expanded:
        _inspect_file(path, seen, accepted, skipped)
    return DiscoveryResult(tuple(accepted), tuple(skipped))


def _assert_writable(directory: Path) -> None:
    try:
        with tempfile.NamedTemporaryFile(
            dir=directory, prefix=".write-test-", delete=True
        ):
            pass
    except OSError as exc:
        raise OutputDirectoryError("保存目录不可写") from exc


def prepare_output_directory(
    source: Path, custom_output_dir: Path | None
) -> Path:
    if custom_output_dir is None:
        try:
            directory = Path(source).resolve(strict=False).parent / "修改版"
        except OSError as exc:
            raise OutputDirectoryError("无法解析源文件路径") from exc
        try:
            directory.mkdir(parents=False, exist_ok=True)
        except OSError as exc:
            raise OutputDirectoryError("无法创建默认保存目录") from exc
    else:
        try:
            directory = Path(custom_output_dir).resolve(strict=False)
        except OSError as exc:
            raise OutputDirectoryError("无法解析自定义保存目录") from exc
        try:
            directory_stat = directory.stat()
        except FileNotFoundError as exc:
            raise OutputDirectoryError("自定义保存目录不存在") from exc
        except OSError as exc:
            raise OutputDirectoryError("无法检查自定义保存目录") from exc
        if not stat.S_ISDIR(directory_stat.st_mode):
            raise OutputDirectoryError("自定义保存目录不存在")
    _assert_writable(directory)
    return directory


def next_output_path(source: Path, output_dir: Path) -> Path:
    source = Path(source)
    output_dir = Path(output_dir)
    base = f"{source.stem} - 修改版"
    index = 0
    while True:
        suffix = "" if index == 0 else f" ({index})"
        candidate = output_dir / f"{base}{suffix}.docx"
        try:
            candidate.lstat()
        except FileNotFoundError:
            return candidate
        except OSError as exc:
            raise OutputDirectoryError("无法检查输出文件是否存在") from exc
        index += 1
