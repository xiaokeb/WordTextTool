from dataclasses import dataclass
from enum import Enum
from pathlib import Path


class FileStatus(str, Enum):
    SUCCESS = "success"
    UNMATCHED = "unmatched"
    SKIPPED = "skipped"
    FAILED = "failed"


@dataclass(frozen=True, slots=True)
class ReplacementRule:
    before: str
    after: str


@dataclass(frozen=True, slots=True)
class FileResult:
    source: Path
    status: FileStatus
    message: str
    output: Path | None = None
    replacements: int = 0
    outputs: tuple[Path, ...] = ()


@dataclass(frozen=True, slots=True)
class DiscoveryResult:
    accepted: tuple[Path, ...] = ()
    skipped: tuple[FileResult, ...] = ()


@dataclass(frozen=True, slots=True)
class BatchSummary:
    results: tuple[FileResult, ...] = ()
    cancelled: bool = False

    def count(self, status: FileStatus) -> int:
        return sum(result.status is status for result in self.results)
