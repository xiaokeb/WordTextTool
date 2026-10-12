"""Publish verified V2.0 assets using the configured GitHub Git credential.

Credentials stay in memory and are sent only to github.com's API/upload hosts.
Existing tags/releases/assets are never deleted or overwritten.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import urllib.error
import urllib.parse
import urllib.request

from candidate_files import ROOT, DEST, file_digest, source_digest

REPOSITORY = "xiaokeb/WordTextTool"
TAG = "V2.0"
API = "https://api.github.com"
RELEASE_DIR = ROOT / "dist" / TAG


def git(*args):
    return subprocess.check_output(["git", *args], cwd=ROOT, text=True).strip()


def credential():
    environment = dict(os.environ, GCM_INTERACTIVE="Never", GIT_TERMINAL_PROMPT="0")
    result = subprocess.run(["git", "credential", "fill"],
                            input="url=https://github.com/" + REPOSITORY + ".git\n\n",
                            cwd=ROOT, text=True, capture_output=True, env=environment, timeout=20)
    values = dict(line.split("=", 1) for line in result.stdout.splitlines() if "=" in line)
    if result.returncode or not values.get("password"):
        raise RuntimeError("No usable GitHub publishing credential in the configured Git credential manager")
    return values["password"]


def request(token, path, *, method="GET", data=None, content_type="application/json", raw=False):
    url = path if path.startswith("https://") else API + path
    if urllib.parse.urlparse(url).hostname not in ("api.github.com", "uploads.github.com"):
        raise RuntimeError("Unexpected GitHub API host")
    body = data if raw or data is None else json.dumps(data, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, method=method, headers={
        "Authorization": "Bearer " + token, "Accept": "application/vnd.github+json",
        "Content-Type": content_type, "User-Agent": "WordTextTool-release",
        "X-GitHub-Api-Version": "2022-11-28"})
    with urllib.request.urlopen(req, timeout=90) as response:
        return json.load(response)


def prepare():
    verified = json.loads((ROOT / "build" / "candidate-verification.json").read_text(encoding="utf-8"))
    manifest = json.loads((DEST / "manifest.json").read_text(encoding="utf-8"))
    audit = json.loads((ROOT / "build" / "environment-audit.json").read_text(encoding="utf-8"))
    if (not verified.get("ok") or set(verified.get("verified", {})) != {"fixed", "automatic", "window"}
            or verified.get("exe_sha256") != manifest["sha256"]
            or source_digest() != manifest["source_and_build_sha256"]
            or audit.get("advisories") is None or audit["advisories"]
            or audit.get("mismatches") or audit.get("extra_packages")
            or audit.get("lock_sha256") != file_digest(ROOT / "requirements-lock.txt")):
        raise RuntimeError("Current source, dependency audit and EXE acceptance must all pass before release preparation")
    if file_digest(DEST / manifest["exe"]) != manifest["sha256"]:
        raise RuntimeError("Validated EXE has changed")
    for name, item in manifest["files"].items():
        if Path(name).name != name or file_digest(DEST / name) != item["sha256"]:
            raise RuntimeError("Candidate attachment has changed")
    commit = git("rev-parse", TAG + "^{commit}")
    if git("status", "--porcelain", "--untracked-files=no"):
        raise RuntimeError("Commit the release inputs before preparing official assets")
    RELEASE_DIR.mkdir(exist_ok=True)
    shutil.copyfile(DEST / manifest["exe"], RELEASE_DIR / "WordTextTool-V2.0.exe")
    shutil.copyfile(ROOT / "docs" / "V2.0发布说明.md", RELEASE_DIR / "V2.0-ReleaseNotes.md")
    shutil.copyfile(DEST / "三表验收样本.xlsx", RELEASE_DIR / "AcceptanceSample.xlsx")
    (RELEASE_DIR / "README.md").write_text("""# Word 文本处理工具 V2.0

双击 WordTextTool-V2.0.exe，无需另装 Python。
提供 Word 文本替换、Excel/WPS 编号及 PDF、文件名后缀管理三个页签。
编号和 PDF 需要本地 WPS 表格或 Microsoft Excel；本版在 Windows 11/WPS 下完成真实验收。
Microsoft Excel 接口已核对身份及调用逻辑，没有真实 Excel 真机验收记录。

编号模板包含一个 yy，按列表顺序给文件编号，所有工作表必须可见有效。
自动模式写 E1:F1；固定模式先插入首行，固定编号写 E1:F1，递增编号写 E2:F2。
两种模式均合并、居中、加粗并设置 Times New Roman。先预览，再写入或一键转换。
Word/Excel 源文件不覆盖；后缀管理会直接修改所选文件名，请核对预览再确认执行。
查看 V2.0-ReleaseNotes.md 了解验证与限制，AcceptanceSample.xlsx 可用于自测。
SHA256SUMS.txt 用于核对附件，manifest.json 记录版本、提交、源码和产物摘要。
""", encoding="utf-8")
    release = {"version": "2.0", "tag": TAG, "release_commit": commit,
               "source_and_build_sha256": manifest["source_and_build_sha256"],
               "built_at_utc": manifest["built_at_utc"], "exe_sha256": manifest["sha256"],
               "validation_scope": "Windows 11/WPS real-engine; Microsoft Excel interface contracts only"}
    assets = ("WordTextTool-V2.0.exe", "README.md", "V2.0-ReleaseNotes.md", "AcceptanceSample.xlsx")
    release["files"] = {name: {"size": (RELEASE_DIR / name).stat().st_size, "sha256": file_digest(RELEASE_DIR / name)} for name in assets}
    (RELEASE_DIR / "manifest.json").write_text(json.dumps(release, ensure_ascii=False, indent=2), encoding="utf-8")
    sums = [(name, item["sha256"]) for name, item in release["files"].items()]
    sums.append(("manifest.json", file_digest(RELEASE_DIR / "manifest.json")))
    (RELEASE_DIR / "SHA256SUMS.txt").write_text("".join(digest + "  " + name + "\n" for name, digest in sums), encoding="utf-8")
    return release, (*assets, "manifest.json", "SHA256SUMS.txt")


def publish(token, release, names):
    # Confirm the remote tag already resolves to this release commit.
    remote_tag = request(token, "/repos/" + REPOSITORY + "/git/ref/tags/" + TAG)["object"]
    while remote_tag["type"] == "tag":
        remote_tag = request(token, "/repos/" + REPOSITORY + "/git/tags/" + remote_tag["sha"])["object"]
    if remote_tag["sha"] != release["release_commit"]:
        raise RuntimeError("Remote tag points to a different commit")
    try:
        result = request(token, "/repos/" + REPOSITORY + "/releases/tags/" + TAG)
    except urllib.error.HTTPError as exc:
        if exc.code != 404:
            raise
        result = request(token, "/repos/" + REPOSITORY + "/releases", method="POST", data={
            "tag_name": TAG, "target_commitish": release["release_commit"],
            "name": "V2.0：办公扩展、全表编号与安全文件管理", "draft": True, "prerelease": False,
            "body": (ROOT / "docs" / "V2.0发布说明.md").read_text(encoding="utf-8")})
    existing = {asset["name"]: asset for asset in result["assets"]}
    for name in names:
        path = RELEASE_DIR / name
        digest = "sha256:" + file_digest(path).lower()
        if name in existing:
            if existing[name].get("digest") != digest:
                raise RuntimeError("An existing Release asset differs: " + name)
            continue
        upload = result["upload_url"].split("{")[0] + "?" + urllib.parse.urlencode({"name": name})
        asset = request(token, upload, method="POST", data=path.read_bytes(), content_type="application/octet-stream", raw=True)
        if asset["size"] != path.stat().st_size or asset.get("digest") != digest:
            raise RuntimeError("Uploaded asset checksum mismatch: " + name)
    result = request(token, "/repos/" + REPOSITORY + "/releases/" + str(result["id"]))
    remote_assets = {item["name"]: item for item in result["assets"]}
    for name in names:
        if remote_assets.get(name, {}).get("digest") != "sha256:" + file_digest(RELEASE_DIR / name).lower():
            raise RuntimeError("Remote attachment verification failed: " + name)
    if result["draft"]:
        result = request(token, "/repos/" + REPOSITORY + "/releases/" + str(result["id"]), method="PATCH", data={"draft": False})
    receipt = {"url": result["html_url"], "tag": TAG, "commit": release["release_commit"],
               "draft": result["draft"], "assets": list(names)}
    (ROOT / "build" / "release-receipt.json").write_text(json.dumps(receipt, indent=2), encoding="utf-8")
    return receipt


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check-auth", action="store_true")
    parser.add_argument("--prepare-only", action="store_true")
    args = parser.parse_args()
    try:
        if args.check_auth:
            repo = request(credential(), "/repos/" + REPOSITORY)
            print(json.dumps({"repository": repo["full_name"], "permissions": repo.get("permissions", {})}))
            return 0
        release, names = prepare()
        result = {"prepared": str(RELEASE_DIR), "tag": TAG, "commit": release["release_commit"]} if args.prepare_only else publish(credential(), release, names)
        print(json.dumps(result, ensure_ascii=True, indent=2))
        return 0
    except urllib.error.HTTPError as exc:
        print("GitHub API returned HTTP " + str(exc.code) + "; publication not confirmed")
    except (OSError, RuntimeError, subprocess.SubprocessError, ValueError) as exc:
        print(str(exc))
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
