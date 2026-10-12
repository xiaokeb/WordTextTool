"""Common reproducible inputs for candidate provenance checks."""
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / "dist" / "V2.0-candidate"


def file_digest(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
    return digest.hexdigest().upper()


def source_digest():
    digest = hashlib.sha256()
    inputs = sorted((ROOT / "src").rglob("*.py")) + [ROOT / name for name in
        ("build.spec", "requirements.txt", "requirements-lock.txt")]
    for path in inputs:
        digest.update(path.relative_to(ROOT).as_posix().encode("utf-8") + b"\0" + path.read_bytes() + b"\0")
    return digest.hexdigest()
