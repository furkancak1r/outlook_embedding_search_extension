"""'Benzerlerini bul' evaluation: worker mode similar on a full eval index, with the real model vectors.

Run after eval/run_eval.py (it builds the index):
  python eval/similar_eval.py --index-from tmp/eval-xyz --out reports/similar-mac.json
The index file is copied into a new temp folder (the source stays unchanged). One kept-alive worker answers one similar
session per family mail. A hit = a mail of the same family (eval/make_corpus.py: one topic, other details).
Pass: family share in the top 5 >= 0.95 and 'yüksek' precision >= 0.85.
"""
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))


def families():
    """{mail id: family key} with the final ids of eval/corpus.json (make_corpus numbers the mails after sorting)."""
    import make_corpus as mc
    fams = mc.build()
    mc.MSGS.sort(key=lambda m: m["date"])
    new_id = {m["id"]: f"M{i:04d}" for i, m in enumerate(mc.MSGS, 1)}
    corpus = {m["id"]: m for m in json.loads((HERE / "corpus.json").read_text(encoding="utf-8"))["messages"]}
    out = {}
    for k, ms in fams.items():
        for m in ms:
            nid = new_id[m["id"]]
            assert corpus[nid]["subject"] == m["subject"], nid  # the corpus file was made by this generator
            out[nid] = k
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index-from", required=True, help="--data folder of a finished eval/run_eval.py run")
    ap.add_argument("--out", required=True)
    ap.add_argument("--python", default=sys.executable)
    args = ap.parse_args()
    corpus = json.loads((HERE / "corpus.json").read_text(encoding="utf-8"))
    name = hashlib.sha256(corpus["store"]["id"].encode()).hexdigest()[:24] + ".sqlite"
    fam = families()
    token = os.urandom(16).hex()
    with tempfile.TemporaryDirectory() as d:
        (Path(d) / "index").mkdir()
        shutil.copy(Path(args.index_from) / "index" / name, Path(d) / "index" / name)
        # Python 3.13 only: the Python 3.6 compatibility rule for Popen(encoding=) does not apply.
        p = subprocess.Popen([args.python, "-I", "-B", "-X", "utf8", str(HERE.parent / "worker" / "hm_worker.py"), "--data", d,  # nosemgrep: python36-compatibility-Popen2
                              "--model", "unused", "--fake-encoder"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                             stderr=subprocess.DEVNULL, text=True, encoding="utf-8", env=dict(os.environ, HM_SESSION=token))
        rows = []
        for eid, k in sorted(fam.items()):
            others = {e for e, f in fam.items() if f == k and e != eid}
            if len(others) < 2:
                continue
            p.stdin.write(json.dumps({"op": "begin", "session": token, "mode": "similar", "similar_eid": eid, "keep_alive_s": 60,
                                      "store": corpus["store"], "folders": [{"id": f["id"]} for f in corpus["folders"]]}) + "\n")
            p.stdin.flush()
            res = None
            while True:
                ev = json.loads(p.stdout.readline())
                if ev["ev"] == "results":
                    res = ev
                if ev["ev"] == "done":
                    break
            assert ev["status"] == "ok" and not ev["model_loaded"], ev
            got = [(r["eid"], r["confidence"], r["cos"]) for r in res["results"]]
            rows.append({"eid": eid, "family": k, "n_family": len(others), "results": got,
                         "at5": sum(e in others for e, _, _ in got[:5]) / min(5, len(others)), "self": any(e == eid for e, _, _ in got)})
        p.stdin.close()
        p.wait(timeout=30)
    lab = {c: [sum(1 for r in rows for e, cc, _ in r["results"] if cc == c and fam.get(e) == r["family"]),
               sum(1 for r in rows for _, cc, _ in r["results"] if cc == c)] for c in ("yüksek", "orta", "zayıf")}
    at5 = sum(r["at5"] for r in rows) / len(rows)
    prec = lab["yüksek"][0] / max(1, lab["yüksek"][1])
    ok = at5 >= 0.95 and prec >= 0.85 and not any(r["self"] for r in rows)
    report = {"queries": len(rows), "family_at5": round(at5, 3), "labels_family_of_total": lab, "high_precision": round(prec, 3),
              "self_in_results": sum(r["self"] for r in rows), "pass": ok, "rows": rows}
    Path(args.out).write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "rows"}, ensure_ascii=False))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
