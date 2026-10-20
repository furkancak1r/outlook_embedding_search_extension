"""Derive the text-only model directory from the pinned official EmbeddingGemma 2 files.

The output model.safetensors holds only the `language_model.*` tensors (the 270M text component).
Tensor bytes are copied unchanged. Vision and audio tensors are not copied, so the add-in never ships or reads them.
Usage: python make_text_only.py <official_dir> <out_dir>
"""
import hashlib
import json
import shutil
import struct
import sys
from pathlib import Path

LOCK = json.loads(Path(__file__).with_name("model.lock.json").read_text(encoding="utf-8"))
COPY = ["config.json", "tokenizer.json", "tokenizer_config.json", "README.md"]


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def read_header(p):
    with open(p, "rb") as f:
        n = struct.unpack("<Q", f.read(8))[0]
        return json.loads(f.read(n)), 8 + n


def main(src, dst):
    src, dst = Path(src), Path(dst)
    want = {f["path"]: f for f in LOCK["files"]}
    if sha256(src / "model.safetensors") != want["model.safetensors"]["sha256"]:
        sys.exit("official model.safetensors hash mismatch")
    header, base = read_header(src / "model.safetensors")
    meta = header.pop("__metadata__", {})
    names = sorted(k for k in header if k.startswith("language_model."))
    out_header, off = {"__metadata__": {**meta, "hatirmail_derived_from": f"{LOCK['repo']}@{LOCK['revision']}",
                                        "hatirmail_subset": "language_model.*"}}, 0
    for k in names:
        a, b = header[k]["data_offsets"]
        out_header[k] = {"dtype": header[k]["dtype"], "shape": header[k]["shape"], "data_offsets": [off, off + b - a]}
        off += b - a
    hb = json.dumps(out_header, separators=(",", ":")).encode("utf-8")
    hb += b" " * (-len(hb) % 8)
    dst.mkdir(parents=True, exist_ok=True)
    tmp = dst / "model.safetensors.part"
    with open(src / "model.safetensors", "rb") as fi, open(tmp, "wb") as fo:
        fo.write(struct.pack("<Q", len(hb)) + hb)
        for k in names:
            a, b = header[k]["data_offsets"]
            fi.seek(base + a)
            fo.write(fi.read(b - a))
    tmp.replace(dst / "model.safetensors")
    for name in COPY:
        shutil.copyfile(src / name, dst / name)
    (dst / "NOTICE.txt").write_text(
        "EmbeddingGemma 2 (Google DeepMind), Apache License 2.0: https://ai.google.dev/gemma/docs/gemma_4_license\n"
        f"Source: https://huggingface.co/{LOCK['repo']} revision {LOCK['revision']}\n"
        "Modified by HatirMail: model.safetensors keeps only the language_model.* tensors (text component).\n"
        "Tensor data is byte-identical to the source. The vision and audio tensors were removed.\n", encoding="utf-8")
    files = {p.name: {"size": p.stat().st_size, "sha256": sha256(p)} for p in sorted(dst.iterdir()) if p.is_file()}
    (dst / "hatirmail-model.json").write_text(json.dumps(
        {"repo": LOCK["repo"], "revision": LOCK["revision"], "subset": "language_model.*", "tensors": len(names),
         "files": files}, indent=1), encoding="utf-8")
    print(f"text-only model: {len(names)} tensors, {(dst / 'model.safetensors').stat().st_size / 2**20:.1f} MB")


if __name__ == "__main__":
    main(*sys.argv[1:3])
