"""Download the pinned EmbeddingGemma 2 files and verify every file.

Usage: python fetch_model.py <dest_dir> [--all]
Without --all, the script downloads only the files that the worker needs.
The script uses only the standard library. It never sends mail data.
"""
import hashlib, json, os, sys, urllib.request
from pathlib import Path

LOCK = json.loads((Path(__file__).with_name("model.lock.json")).read_text(encoding="utf-8"))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def ok(path, f):
    return path.is_file() and path.stat().st_size == f["size"] and sha256(path) == f["sha256"]


def main():
    dest = Path(sys.argv[1]).resolve()
    want_all = "--all" in sys.argv[2:]
    for f in LOCK["files"]:
        if not (f["runtime"] or want_all):
            continue
        out = dest / f["path"]
        if ok(out, f):
            print("ok", f["path"]); continue
        out.parent.mkdir(parents=True, exist_ok=True)
        part = out.with_name(out.name + ".part")
        url = f"https://huggingface.co/{LOCK['repo']}/resolve/{LOCK['revision']}/{f['path']}"
        have = part.stat().st_size if part.exists() else 0
        req = urllib.request.Request(url, headers={"Range": f"bytes={have}-"} if have else {})
        # The scheme is the fixed https:// above; repo, revision and path come from the committed lock file.
        with urllib.request.urlopen(req, timeout=60) as r, open(part, "ab" if have and r.status == 206 else "wb") as fh:  # nosemgrep: dynamic-urllib-use-detected
            while block := r.read(1 << 20):
                fh.write(block)
        if not ok(part, f):
            part.unlink()
            sys.exit(f"hash mismatch: {f['path']}")
        os.replace(part, out)
        print("downloaded", f["path"])
    print("model files verified at revision", LOCK["revision"])


if __name__ == "__main__":
    main()
