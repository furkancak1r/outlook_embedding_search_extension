"""Dump raw ranking signals for every query (dev tool for calibration). Loads the model once, in-process.

Usage: python eval/dump_signals.py <data_dir_with_index> <model_dir> <out.json>
The index must already exist (run eval/run_eval.py first). Output: per query the top 30 candidates with
cosine, z, keyword coverage, keyword rank, semantic rank and date score. No final ranking decisions here.
"""
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "worker"))
import numpy as np  # noqa: E402

import hm_dates  # noqa: E402
import hm_embed  # noqa: E402
import hm_search as S  # noqa: E402
import hm_worker  # noqa: E402
from hm_store import Store  # noqa: E402
from hm_text import fold  # noqa: E402

data, model, out = sys.argv[1:4]
corpus = json.loads((ROOT / "eval" / "corpus.json").read_text(encoding="utf-8"))
queries = json.loads((ROOT / "eval" / "queries.json").read_text(encoding="utf-8"))
store = Store(str(hm_worker.db_path(data, corpus["store"]["id"])), hm_embed.MODEL_REVISION)
enc = hm_embed.Encoder(model, 4)
now = datetime.fromisoformat(corpus["now"])
scope = [f["id"] for f in corpus["folders"]]
ids, seq, mat = store.all_vectors(hm_embed.MODEL_REVISION, hm_embed.DIM, scope)
rows = []
for q in queries:
    windows, spans, keep = hm_dates.parse(q["q"], now)
    text_q = (hm_dates.remove_spans(q["q"], spans) + " " + " ".join(keep)).strip()
    groups = S.query_groups(text_q, store.vocab())
    regexes = [S.group_regex(vs) for _, vs in groups]
    qv = enc.encode([hm_embed.query_text(text_q or q["q"])], 128)[0]
    qv_full = enc.encode([hm_embed.query_text(q["q"])], 128)[0]
    sims, sims_full = mat @ qv, mat @ qv_full
    best, best_full = {}, {}
    for i, r in enumerate(ids.tolist()):
        best[r] = max(best.get(r, -1.0), float(sims[i]))
        best_full[r] = max(best_full.get(r, -1.0), float(sims_full[i]))
    vals = np.array(list(best.values()))
    med = float(np.median(vals))
    sig = float(np.median(np.abs(vals - med))) * 1.4826
    kw = [r for r, _ in store.fts(S.match_expr(groups))] if groups else []
    info = store.items(set(best) | set(kw))
    cand = []
    for r, it in info.items():
        hay = fold(" ".join(it[k] or "" for k in ("subject", "sender", "sender_addr", "recipients", "attachments", "main_text")))
        cov = sum(1 for rx in regexes if rx.search(hay)) / len(regexes) if regexes else 0.0
        dt = datetime.fromisoformat(it["sent_at"]) if it["sent_at"] else None
        cand.append({"eid": it["entry_id"], "subject": it["subject"], "cos": round(best.get(r, -1), 4),
                     "cos_full": round(best_full.get(r, -1), 4), "z": round((best.get(r, -1) - med) / sig, 2),
                     "cov": round(cov, 2), "kw_rank": kw.index(r) + 1 if r in kw else None,
                     "ds": round(hm_dates.date_score(dt, windows), 3) if windows and dt else 0.0})
    cand.sort(key=lambda c: -c["cos"])
    for i, c in enumerate(cand, 1):
        c["sem_rank"] = i
    top = cand[:30] + [c for c in cand[30:] if c["eid"] in q["targets"] or c["kw_rank"] and c["kw_rank"] <= 10]
    rows.append({**q, "text_q": text_q, "groups": [[g, vs] for g, vs in groups], "median": round(med, 4),
                 "sigma": round(sig, 4), "windows": hm_dates.describe(windows), "cands": top})
Path(out).write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
print("queries", len(rows))
