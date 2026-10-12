"""Verify the exact WordMod dependency lock; optionally query OSV advisories."""
import argparse
from importlib.metadata import distributions
import json
from pathlib import Path
import platform
import sys
import urllib.request
import hashlib
from datetime import datetime, timezone

ROOT = Path(__file__).resolve().parents[1]


def canonical(name):
    return name.lower().replace("_", "-").replace(".", "-")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit", action="store_true")
    args = parser.parse_args()
    installed = {canonical(d.metadata["Name"]): d.version for d in distributions()}
    lock = {}
    for line in (ROOT / "requirements-lock.txt").read_text(encoding="utf-8").splitlines():
        if "==" in line and not line.startswith("#"):
            name, version = line.split("==")
            lock[canonical(name)] = version
    mismatch = {name: {"expected": version, "actual": installed.get(name)}
                for name, version in lock.items() if installed.get(name) != version}
    extra = {name: version for name, version in installed.items() if name not in lock}
    report = {"python": platform.python_version(), "executable": sys.executable,
              "locked_packages": len(lock), "mismatches": mismatch, "extra_packages": extra,
              "checked_at_utc": datetime.now(timezone.utc).isoformat(),
              "lock_sha256": hashlib.sha256((ROOT / "requirements-lock.txt").read_bytes()).hexdigest().upper()}
    if args.audit:
        queries = [{"package": {"name": n, "ecosystem": "PyPI"}, "version": v} for n, v in lock.items()]
        request = urllib.request.Request("https://api.osv.dev/v1/querybatch",
                                         json.dumps({"queries": queries}).encode(),
                                         {"Content-Type": "application/json"})
        with urllib.request.urlopen(request, timeout=30) as response:
            results = json.load(response)["results"]
        report["advisories"] = {name: result["vulns"] for name, result in zip(lock, results) if result.get("vulns")}
        report["audit_source"] = "https://api.osv.dev/v1/querybatch"
    destination = ROOT / "build" / ("environment-audit.json" if args.audit else "environment-verification.json")
    destination.parent.mkdir(exist_ok=True)
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 1 if mismatch or extra or report.get("advisories") else 0


if __name__ == "__main__":
    raise SystemExit(main())
