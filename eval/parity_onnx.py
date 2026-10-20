"""Parity check: torch (hm_embed.TorchEncoder, the reference) vs ONNX Runtime (hm_embed.Encoder) on the same texts.

Texts: every query and every document chunk of eval/corpus.json (as the worker makes them), plus edge cases (1 token,
512 tokens, empty text, Turkish capitals) in batch 1 and in batches with padding. Pass: max |diff| <= 1e-4 and
cosine >= 0.99999 for every vector (the index stores float16, its step at 0.05 is ~3e-5).
Usage (Mac, .venv-onnx has torch): parity_onnx.py <torch model dir> <onnx model dir> <out.json> [threads]
  This also writes <out>.npz with the torch vectors.
Usage (2.19, no torch in the runtime): parity_onnx.py <ref.npz> <onnx model dir> <out.json> [threads]
  Compares ONNX Runtime on this machine with the torch vectors from the Mac.
"""
import json
import sys
import time
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "worker"))
import hm_embed  # noqa: E402
import hm_index  # noqa: E402

ref_src, onnx_dir, out = sys.argv[1], sys.argv[2], Path(sys.argv[3])
threads = int(sys.argv[4]) if len(sys.argv) > 4 else 4
corpus = json.loads((HERE / "corpus.json").read_text(encoding="utf-8"))
queries = json.loads((HERE / "queries.json").read_text(encoding="utf-8"))
docs = []
for m in corpus["messages"]:  # the model inputs that the worker makes for these mails
    rec, _, spans, _, _ = hm_index.prepare({"eid": m["id"], "folder": m["folder"], "subject": m["subject"], "body": m["body"],
                                            "sender": m["from_name"], "attachments": m["attachments"]})
    docs += hm_index.chunk_texts(rec["subject"], rec["sender"], rec["attachments"], rec["main_text"], spans)
qs = [hm_embed.query_text(q["q"]) for q in queries]
edge = ["a", " ".join(["Proje toplantısında bütçe, takvim ve riskler konuşuldu."] * 120), "", "İŞÇİ ığdır ŞÜKRÜ 4.250 TL"]
SETS = (("queries_b1", qs, hm_embed.MAX_QUERY_TOKENS, 1), ("queries_b8", qs, hm_embed.MAX_QUERY_TOKENS, 8),
        ("docs_b8", docs, hm_embed.MAX_DOC_TOKENS, 8), ("edge_b1", edge, hm_embed.MAX_DOC_TOKENS, 1),
        ("edge_b4", edge, hm_embed.MAX_DOC_TOKENS, 4))


def run(enc, texts, max_tokens, batch):
    return np.concatenate([enc.encode(texts[i:i + batch], max_tokens) for i in range(0, len(texts), batch)])


t = time.perf_counter()
ort = hm_embed.Encoder(onnx_dir, threads)
t_ort = time.perf_counter() - t
if ref_src.endswith(".npz"):
    ref, t_ref = dict(np.load(ref_src)), None
else:
    t = time.perf_counter()
    torch_enc = hm_embed.TorchEncoder(ref_src, threads)
    t_ref = time.perf_counter() - t
    ref = {name: run(torch_enc, texts, mt, batch) for name, texts, mt, batch in SETS}
    np.savez(str(out) + ".npz", **ref)

rows = {}
for name, texts, mt, batch in SETS:
    a = ref[name]
    t = time.perf_counter()
    b = run(ort, texts, mt, batch)
    tb = time.perf_counter() - t
    cos = (a * b).sum(1) / (np.linalg.norm(a, axis=1) * np.linalg.norm(b, axis=1))
    rows[name] = {"n": len(texts), "max_abs_diff": float(np.abs(a - b).max()), "min_cos": float(cos.min()),
                  # as stored in the index (float16): at most one rounding step apart
                  "f16_max_abs_diff": float(np.abs(a.astype(np.float16).astype(np.float32) - b.astype(np.float16).astype(np.float32)).max()),
                  "onnx_s": round(tb, 2)}
    print(name, rows[name], flush=True)
ok = all(r["max_abs_diff"] <= 1e-4 and r["min_cos"] >= 0.99999 for r in rows.values())
report = {"pass": ok, "platform": sys.platform, "threads": threads, "reference": "npz" if t_ref is None else "torch",
          "torch_load_s": t_ref and round(t_ref, 2), "onnx_load_s": round(t_ort, 2), "rows": rows}
out.write_text(json.dumps(report, indent=1))
print("PASS" if ok else "FAIL", "onnx_load_s", report["onnx_load_s"])
sys.exit(0 if ok else 1)
