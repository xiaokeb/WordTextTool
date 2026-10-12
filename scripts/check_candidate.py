"""Validate candidate contents, source provenance and recorded EXE acceptance."""
import argparse
import json
from pathlib import Path
from candidate_files import ROOT, DEST, file_digest, source_digest


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--office-report", type=Path)
    parser.add_argument("--automatic-report", type=Path)
    parser.add_argument("--window-report", type=Path)
    args = parser.parse_args()
    manifest = json.loads((DEST / "manifest.json").read_text(encoding="utf-8"))
    errors = []
    if manifest["source_and_build_sha256"] != source_digest():
        errors.append("Candidate source/build digest is stale")
    for name, expected in manifest["files"].items():
        path = DEST / name
        if path.parent != DEST or not path.is_file() or file_digest(path) != expected["sha256"] or path.stat().st_size != expected["size"]:
            errors.append(f"Candidate content mismatch: {name}")
    if file_digest(DEST / manifest["exe"]) != manifest["sha256"]:
        errors.append("EXE digest mismatch")
    verified = {}
    for label, path in (("fixed", args.office_report), ("automatic", args.automatic_report)):
        if path:
            report = json.loads(path.read_text(encoding="utf-8"))
            if (report.get("status") != "success" or not report.get("source_unchanged")
                    or not report.get("all_sheets_verified") or report.get("mode") != label
                    or report.get("candidate_sha256") != manifest["sha256"]
                    or any(call.get("engine") != "wps" for call in report.get("calls", []))):
                errors.append(f"{label} EXE/WPS acceptance is missing or stale")
            verified[label] = str(path)
    if args.window_report:
        report = json.loads(args.window_report.read_text(encoding="utf-8"))
        if not report.get("main_window_created") or report.get("normal_close_exit_code") != 0 or report.get("candidate_sha256") != manifest["sha256"]:
            errors.append("EXE window acceptance is missing or stale")
        verified["window"] = str(args.window_report)
    result = {"ok": not errors, "errors": errors, "exe_sha256": manifest["sha256"], "verified": verified}
    (ROOT / "build" / "candidate-verification.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
