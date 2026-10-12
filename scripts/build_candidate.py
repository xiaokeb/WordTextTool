"""Build V2 separately from V1 and bind provenance to the successful build."""
from datetime import datetime, timezone
import json
import subprocess
import sys
from candidate_files import ROOT, DEST, file_digest, source_digest


def main():
    receipt = ROOT / "build" / "candidate-build.json"
    receipt.parent.mkdir(exist_ok=True)
    # Invalidate a previous successful receipt before beginning a new build.
    receipt.write_text(json.dumps({"ok": False}), encoding="utf-8")
    subprocess.run([sys.executable, "-m", "pip", "check"], cwd=ROOT, check=True)
    subprocess.run([sys.executable, str(ROOT / "scripts" / "check_environment.py")], cwd=ROOT, check=True)
    before = source_digest()
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
                    "--distpath", str(DEST), "--workpath", str(ROOT / "build" / "V2.0"),
                    str(ROOT / "build.spec")], cwd=ROOT, check=True)
    if before != source_digest():
        raise RuntimeError("Source changed during the build; rebuild before packaging")
    exe = DEST / "Word 文本处理工具.exe"
    receipt.write_text(json.dumps({"ok": True, "source_and_build_sha256": before,
                                  "exe_sha256": file_digest(exe), "size": exe.stat().st_size,
                                  "built_at_utc": datetime.now(timezone.utc).isoformat()}, indent=2), encoding="utf-8")
    subprocess.run([sys.executable, str(ROOT / "scripts" / "package_candidate.py")], cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
