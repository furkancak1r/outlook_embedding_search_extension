"""Offline calibration of the fusion weights and the confidence thresholds (dev tool).

Input: tmp/signals.json from eval/dump_signals.py. Grid search on the dev half (odd query ids),
report on the test half (even query ids). Prints the chosen parameters; they are copied by hand into
worker/hm_search.py (the worker never reads this file).
"""
import itertools
import json
import sys


def rank_rows(rows, w_cov, w_date, w_kw, use_full):
    out = []
    for r in rows:
        n = max(1, len(r["groups"]))
        scored = []
        for c in r["cands"]:
            cos = c["cos_full"] if use_full else c["cos"]
            kw = 1.0 / (10 + c["kw_rank"]) if c["kw_rank"] else 0.0
            s = cos + w_cov * c["cov"] * min(1.0, n / 2) + w_date * c["ds"] + w_kw * kw
            scored.append((s, c))
        scored.sort(key=lambda x: -x[0])
        out.append((r, [c for _, c in scored[:10]]))
    return out


def metrics(ranked):
    pos = [(r, top) for r, top in ranked if r["targets"]]
    rk = [next((i + 1 for i, c in enumerate(top) if c["eid"] in r["targets"]), None) for r, top in pos]
    return {"r@1": round(sum(k == 1 for k in rk) / len(rk), 3), "r@3": round(sum(1 for k in rk if k and k <= 3) / len(rk), 3),
            "mrr": round(sum(1 / k for k in rk if k) / len(rk), 3)}


def main(path):
    rows = json.load(open(path, encoding="utf-8"))
    dev = [r for r in rows if int(r["id"][1:]) % 2 == 1]
    test = [r for r in rows if int(r["id"][1:]) % 2 == 0]
    best = None
    for w_cov, w_date, w_kw, full in itertools.product([0, .02, .04, .06, .08, .1, .15], [0, .02, .04, .06, .1],
                                                        [0, .5, 1, 2], [False, True]):
        m = metrics(rank_rows(dev, w_cov, w_date, w_kw, full))
        key = (m["mrr"], m["r@1"], -w_cov - w_date - w_kw)
        if best is None or key > best[0]:
            best = (key, (w_cov, w_date, w_kw, full), m)
    _, params, mdev = best
    print("best params (w_cov, w_date, w_kw, use_full_query):", params)
    print("dev :", mdev)
    print("test:", metrics(rank_rows(test, *params)))
    print("all :", metrics(rank_rows(rows, *params)))
    # Confidence: features of the top result for negatives vs the target for positives.
    ranked = rank_rows(rows, *params)
    pos, neg = [], []
    for r, top in ranked:
        n = len(r["groups"])
        if r["targets"]:
            t = next((c for c in top if c["eid"] in r["targets"]), None)
            if t:
                pos.append((r["id"], t["cos_full" if params[3] else "cos"], t["z"], t["cov"], n))
        else:
            for c in top[:3]:
                neg.append((r["id"], c["cos_full" if params[3] else "cos"], c["z"], c["cov"], n))
    print("targets  (id, cos, z, cov, n):", sorted(pos, key=lambda x: x[1])[:12])
    print("negatives top3 (id, cos, z, cov, n):", sorted(neg, key=lambda x: -x[1])[:12])


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "tmp/signals.json")
