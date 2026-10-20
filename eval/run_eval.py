"""End-to-end evaluation of the HatırMail worker with the real model (no Outlook needed).

It plays the add-in side of the JSON-lines protocol: one worker process per search, as in production.
Measures: first index time, cold search time (model load included), Recall@1/@10, MRR, negative-query behavior,
peak worker RAM, and that every worker process exits after its search.
Usage: python eval/run_eval.py --python <python> --model <text-only model dir> --data <empty dir> --out <json>
"""
import argparse
import json
import os
import secrets
import subprocess
import sys
import threading
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
WORKER = HERE.parent / "worker" / "hm_worker.py"


def to_item(m, folders):
    it = {"eid": m["id"], "folder": m["folder"], "cls": m["cls"], "mod": m["mod"], "subject": m["subject"],
          "sender": m["from_name"], "sender_addr": m["from_addr"], "to": "; ".join(n for n, _ in m["to"]),
          "cc": "; ".join(n for n, _ in m["cc"]), "date": m["date"], "attachments": m["attachments"], "body": m["body"]}
    if m.get("att_paths"):  # attachment test corpus: real files next to the corpus
        it["att_meta"] = [[i, Path(p).name, (HERE / p).stat().st_size, 1] for i, p in enumerate(m["att_paths"], 1)]
    return it


def rss_sampler(pid, stop, out):
    """Sample the worker memory (MB) from outside the process, every 100 ms (Windows: peak working set)."""
    peak = 0.0
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes

        class PMC(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD)] + [
                (n, ctypes.c_size_t) for n in ("PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                                               "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage",
                                               "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]
        k32 = ctypes.WinDLL("kernel32")
        k32.OpenProcess.restype = wintypes.HANDLE
        h = k32.OpenProcess(0x1000 | 0x0010, False, pid)  # QUERY_LIMITED_INFORMATION | VM_READ
        pmc = PMC()
        pmc.cb = ctypes.sizeof(PMC)
        commit = 0.0
        while not stop.is_set():
            if h and ctypes.WinDLL("psapi").GetProcessMemoryInfo(wintypes.HANDLE(h), ctypes.byref(pmc), pmc.cb):
                peak = max(peak, pmc.PeakWorkingSetSize / 2**20)
                commit = max(commit, pmc.PeakPagefileUsage / 2**20)
            stop.wait(0.1)
        if h:
            k32.CloseHandle(wintypes.HANDLE(h))
        out["sampled_peak_commit_mb"] = round(commit, 1)
    else:
        while not stop.is_set():
            try:
                r = subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True, timeout=5)
                peak = max(peak, float(r.stdout.strip() or 0) / 1024)
            except (ValueError, subprocess.SubprocessError, OSError):
                pass
            stop.wait(0.1)
    out["sampled_peak_mb"] = round(peak, 1)


def start(args):
    token = secrets.token_hex(16)
    env = dict(os.environ, HM_SESSION=token, PYTHONIOENCODING="utf-8")
    # Python 3.13 only: the Python 3.6 compatibility rule for Popen(encoding=) does not apply.
    p = subprocess.Popen([args.python, "-I", "-B", "-X", "utf8", str(WORKER), "--data", args.data, "--model", args.model,  # nosemgrep: python36-compatibility-Popen2
                          "--threads", str(args.threads)], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                         stderr=subprocess.DEVNULL, env=env, text=True, encoding="utf-8")
    p.hm_token = token
    return p


def session(args, corpus, query, change=None, cancel_after_s=None, proc=None):
    """One search. proc=None: a new worker process for this search only (cold, as with the keep-alive setting 0).
    proc = a kept-alive worker (search window open): the session runs in it and it stays alive."""
    folders = corpus["folders"]
    msgs = corpus["messages"] if change is None else change
    by = {m["id"]: m for m in msgs}
    att_dir = Path(args.data) / "tmp-att"
    att_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    p = proc or start(args)
    token = p.hm_token
    stop, mem = threading.Event(), {}
    th = threading.Thread(target=rss_sampler, args=(p.pid, stop, mem), daemon=True)
    th.start()

    def send(**m):
        p.stdin.write(json.dumps(m, ensure_ascii=False) + "\n")
        p.stdin.flush()

    first_results = {}

    def wait_for(name):
        evs = []
        while True:
            line = p.stdout.readline()
            if not line:
                return evs
            ev = json.loads(line)
            evs.append(ev)
            if ev["ev"] == "results" and ev.get("stage") not in first_results:
                first_results[ev["stage"]] = round(time.perf_counter() - t0, 3)
            if ev["ev"] == "need_att":  # play the add-in: save the asked files, as Attachment.SaveAsFile does
                files = []
                for eid, idx, _ in ev["atts"]:
                    dst = att_dir / ("%s-%d.bin" % (eid, idx))
                    dst.write_bytes((HERE / by[eid]["att_paths"][idx - 1]).read_bytes())
                    files.append([eid, idx, str(dst)])
                send(op="att_files", files=files)
            elif ev["ev"] in (name, "done"):
                return evs

    # No index_budget_s: a search without time limit (the index is complete before the first measured query).
    send(op="begin", session=token, query=query, now=corpus["now"], top_k=10, attachments=True, att_dir=str(att_dir),
         store=corpus["store"], folders=[{"id": f["id"], "path": f["path"]} for f in folders],
         keep_alive_s=600 if proc else 0)
    evs = wait_for("hello")
    for f in folders:
        items = [[m["id"], m["mod"], m["cls"], m["date"]] for m in msgs if m["folder"] == f["id"]]
        for i in range(0, max(len(items), 1), 200):
            send(op="manifest", folder=f["id"], items=items[i:i + 200], final=i + 200 >= len(items))
    send(op="scan_done")
    evs += wait_for("need")
    need = set(evs[-1].get("eids", []))
    todo = [to_item(by[e], folders) for e in need if e in by]
    for i in range(0, len(todo), 25):
        send(op="items", items=todo[i:i + 25])
    send(op="items_done", failed=[])
    if cancel_after_s is not None:
        def _cancel():
            time.sleep(cancel_after_s)
            try:
                send(op="cancel")
            except OSError:
                pass
        threading.Thread(target=_cancel, daemon=True).start()
    evs += wait_for("done")
    wall = time.perf_counter() - t0
    code = exit_s = alive = None
    if proc is None:
        try:
            p.stdin.close()
        except OSError:
            pass
        code = p.wait(timeout=60)
        exit_s = time.perf_counter() - t0 - wall
        alive = p.poll() is None
    stop.set()
    th.join(timeout=5)
    res = {e["stage"]: e for e in evs if e["ev"] == "results"}
    done = evs[-1] if evs and evs[-1]["ev"] == "done" else {"status": "no-done"}
    return {"wall_s": round(wall, 3), "exit_code": code, "exit_after_done_s": exit_s and round(exit_s, 3), "alive_after": alive,
            "done": done, "final": res.get("final"), "keyword": res.get("keyword"), "need": len(need),
            "keyword_results_s": first_results.get("keyword"), "final_results_s": first_results.get("final"), **mem}


def metrics(rows):
    pos = [r for r in rows if r["targets"]]
    neg = [r for r in rows if not r["targets"]]
    rank = lambda r: next((i + 1 for i, x in enumerate(r["ranked"]) if x in r["targets"]), None)
    out = {"n": len(rows), "n_pos": len(pos), "n_neg": len(neg)}
    if pos:
        ranks = [rank(r) for r in pos]
        out["recall@1"] = round(sum(1 for k in ranks if k == 1) / len(pos), 3)
        out["recall@3"] = round(sum(1 for k in ranks if k and k <= 3) / len(pos), 3)
        out["recall@10"] = round(sum(1 for k in ranks if k) / len(pos), 3)
        out["mrr@10"] = round(sum(1 / k for k in ranks if k) / len(pos), 3)
        out["target_conf_high_or_medium"] = round(sum(1 for r in pos if r["target_conf"] in ("yüksek", "orta")) / len(pos), 3)
    if neg:
        out["neg_no_high_medium"] = round(sum(1 for r in neg if not any(c in ("yüksek", "orta") for c in r["confs"])) / len(neg), 3)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--python", default=sys.executable)
    ap.add_argument("--model", required=True)
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--threads", type=int, default=4)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--cancel-after", type=float, default=8.0, help="seconds until cancel in the cancel test")
    ap.add_argument("--att", action="store_true", help="attachment test set (att_corpus.json, att_queries.json)")
    args = ap.parse_args()
    if args.att:
        return att_eval(args)
    corpus = json.loads((HERE / "corpus.json").read_text(encoding="utf-8"))
    queries = json.loads((HERE / "queries.json").read_text(encoding="utf-8"))
    if args.limit:
        queries = queries[:args.limit]
    Path(args.data).mkdir(parents=True, exist_ok=True)
    report = {"python": args.python, "threads": args.threads, "platform": sys.platform, "messages": len(corpus["messages"])}

    def save(stage):  # write after every stage, so an interruption keeps the finished part
        report["last_stage"] = stage
        Path(args.out).write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")

    # 1) First search on an empty index = first full index + first query.
    first = session(args, corpus, queries[0]["q"])
    report["first_index"] = {k: first.get(k) for k in ("wall_s", "exit_code", "alive_after", "need", "sampled_peak_mb",
                                                       "sampled_peak_commit_mb")}
    report["first_index"].update(first["done"].get("timing", {}), embed=first["done"].get("stats", {}).get("embed"),
                                 peak_rss_mb_self=first["done"].get("peak_rss_mb"),
                                 keyword_results_s=first["keyword_results_s"])
    save("first_index")

    # 2) Every query = one cold worker process (model load + query), unchanged index.
    rows, walls, peaks, exits, alive, kw_s = [], [], [], [], 0, []
    for q in queries:
        r = session(args, corpus, q["q"])
        fin = r["final"] or {"results": []}
        ranked = [x["eid"] for x in fin["results"]]
        confs = [x["confidence"] for x in fin["results"]]
        tconf = next((x["confidence"] for x in fin["results"] if x["eid"] in q["targets"]), None)
        rows.append({**q, "ranked": ranked, "confs": confs, "target_conf": tconf,
                     "scores": [x["score"] for x in fin["results"]],
                     "top1": fin["results"][0]["subject"] if fin["results"] else None,
                     "date_windows": fin.get("date_windows"), "terms": fin.get("terms"),
                     "wall_s": r["wall_s"], "need": r["need"], "status": r["done"].get("status")})
        walls.append(r["wall_s"])
        kw_s.append(r["keyword_results_s"] or 0)
        peaks.append(r.get("sampled_peak_mb") or 0)
        exits.append(r["exit_after_done_s"])
        alive += r["alive_after"]
        print(f'{q["id"]} {q["category"]:12} rank={next((i + 1 for i, x in enumerate(ranked) if x in q["targets"]), "-")} '
              f'conf={tconf} wall={r["wall_s"]:.2f}s', flush=True)
        report["rows"] = rows
        if len(rows) % 10 == 0:
            save(f"queries {len(rows)}")
    walls.sort()
    kw_s.sort()
    report["cold_search"] = {"n": len(walls), "median_s": walls[len(walls) // 2], "p90_s": walls[int(len(walls) * 0.9) - 1],
                             "max_s": walls[-1], "keyword_results_median_s": kw_s[len(kw_s) // 2],
                             "peak_rss_mb_max": max(peaks), "exit_after_done_max_s": max(exits),
                             "processes_alive_after": alive}
    report["metrics_all"] = metrics(rows)
    cats = sorted({r["category"] for r in rows})
    report["metrics_by_category"] = {c: metrics([r for r in rows if r["category"] == c]) for c in cats}
    dev = [r for r in rows if int(r["id"][1:]) % 2 == 1]
    test = [r for r in rows if int(r["id"][1:]) % 2 == 0]
    report["metrics_dev_odd"], report["metrics_test_even"] = metrics(dev), metrics(test)
    save("queries")

    # 2b) Search window open (keep-alive): one worker, status check + model preload, then every query in it.
    report["warm_search"] = warm(args, corpus, queries, rows)
    save("warm")

    # 3) Incremental run: change 3 mails, delete 2, move 1 -> only changed items are read and embedded.
    msgs = [dict(m) for m in corpus["messages"]]
    for m in msgs[10:13]:
        m["body"] = m["body"] + "\n\nEk not: güncellendi."
        m["mod"] = "2026-10-14T11:00:00"
    moved = msgs[20]
    moved["folder"] = "F-ARCH" if moved["folder"] != "F-ARCH" else "F-INBOX"
    del msgs[30:32]
    inc = session(args, corpus, "kargo faturası", change=msgs)
    report["incremental"] = {"need": inc["need"], "wall_s": inc["wall_s"], "stats": inc["done"].get("stats"),
                             "timing": inc["done"].get("timing")}
    save("incremental")

    # 4) Cancel during the embedding of a fresh store and resume.
    other = dict(corpus, store={"id": "STORE-SYNTH-CANCEL", "name": "İptal testi"})
    c1 = session(args, other, "fatura", cancel_after_s=args.cancel_after)
    c2 = session(args, other, "fatura")
    report["cancel_resume"] = {
        "cancel": {"status": c1["done"].get("status"), "exit_code": c1["exit_code"], "alive_after": c1["alive_after"],
                   "wall_s": c1["wall_s"]},
        "resume": {"status": c2["done"].get("status"), "need": c2["need"], "embed": c2["done"].get("stats", {}).get("embed"),
                   "coverage_complete": (c2["final"] or {}).get("coverage", {}).get("complete")}}
    save("done")
    print(json.dumps({k: v for k, v in report.items() if k != "rows"}, ensure_ascii=False, indent=1))


def warm(args, corpus, queries, cold_rows):
    """Window open: status check with preload (the model loads in the background), then the searches reuse the loaded
    model. Checks: the same ranking as the cold run, one process for all, exit 0 when stdin closes."""
    t0 = time.perf_counter()
    p = start(args)
    p.stdin.write(json.dumps({"op": "begin", "session": p.hm_token, "mode": "status", "store": corpus["store"],
                              "folders": [{"id": f["id"]} for f in corpus["folders"]], "keep_alive_s": 600,
                              "preload": True}) + "\n")
    p.stdin.flush()
    states, load_s = [], None
    while not states or states[-1] not in ("ready", "error"):
        line = p.stdout.readline()
        if not line:
            break
        ev = json.loads(line)
        if ev["ev"] == "model":
            states.append(ev["state"])
            load_s = ev.get("load_s")
    preload_s = round(time.perf_counter() - t0, 3)
    walls, same, warm_flags = [], 0, 0
    for q, cold in zip(queries, cold_rows):
        r = session(args, corpus, q["q"], proc=p)
        walls.append(r["wall_s"])
        same += [x["eid"] for x in (r["final"] or {"results": []})["results"]] == cold["ranked"]
        warm_flags += bool(r["done"].get("timing", {}).get("model_warm"))
    p.stdin.close()
    code = p.wait(timeout=60)
    walls.sort()
    return {"preload_states": states, "preload_wall_s": preload_s, "model_load_s": load_s,
            "n": len(walls), "median_s": walls[len(walls) // 2], "p90_s": walls[int(len(walls) * 0.9) - 1],
            "max_s": walls[-1], "sessions_with_model_warm": warm_flags, "same_ranking_as_cold": same,
            "exit_code_after_stdin_close": code}


def att_eval(args):
    """Attachment content: the query words are only inside the attachments. First search = attachment reading
    + indexing; then one cold worker per query. OCR categories need Windows (on the Mac they are expected misses)."""
    corpus = json.loads((HERE / "att_corpus.json").read_text(encoding="utf-8"))
    queries = json.loads((HERE / "att_queries.json").read_text(encoding="utf-8"))
    Path(args.data).mkdir(parents=True, exist_ok=True)
    report = {"python": args.python, "threads": args.threads, "platform": sys.platform, "messages": len(corpus["messages"])}
    first = session(args, corpus, queries[0]["q"])
    report["first_index"] = {"wall_s": first["wall_s"], "exit_code": first["exit_code"], "alive_after": first["alive_after"],
                             "sampled_peak_mb": first.get("sampled_peak_mb"), "att": first["done"].get("stats", {}).get("att"),
                             "timing": first["done"].get("timing"),
                             "att_tmp_files_left": len(list((Path(args.data) / "tmp-att").iterdir()))}
    rows = []
    for q in queries:
        r = session(args, corpus, q["q"])
        fin = r["final"] or {"results": []}
        ranked = [x["eid"] for x in fin["results"]]
        tconf = next((x["confidence"] for x in fin["results"] if x["eid"] in q["targets"]), None)
        in_att = next((x.get("match_attachment") for x in fin["results"] if x["eid"] in q["targets"]), None)
        rows.append({**q, "ranked": ranked, "confs": [x["confidence"] for x in fin["results"]], "target_conf": tconf,
                     "target_match_attachment": in_att, "wall_s": r["wall_s"], "alive_after": r["alive_after"]})
        print(f'{q["id"]} {q["category"]:10} rank={next((i + 1 for i, x in enumerate(ranked) if x in q["targets"]), "-")} '
              f'conf={tconf} ek={in_att}', flush=True)
    report["rows"] = rows
    report["metrics_all"] = metrics(rows)
    report["metrics_by_category"] = {c: metrics([r for r in rows if r["category"] == c]) for c in sorted({r["category"] for r in rows})}
    report["processes_alive_after"] = sum(r["alive_after"] for r in rows)
    Path(args.out).write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "rows"}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
