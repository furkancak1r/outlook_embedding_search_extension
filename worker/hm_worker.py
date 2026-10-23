"""HatırMail worker process.

The Outlook add-in starts this process only on a user action (window open, Ara, İndekslemeyi başlat). Modes (begin.mode):

  search: begin -> manifest* -> scan_done -> (need) -> items* -> items_done -> keyword results
          -> (need_att -> att_files)* -> embed new chunks -> final results -> done
          index_budget_s limits the indexing work (attachments + embedding); the rest waits for the next run.
  index:  the same without a query and without a time limit ("İndekslemeyi başlat"). No results. When the read phase
          ends without a stop, the listed folders count as indexed at this session (folders.synced_gen).
  update: "Güncelle": the same as index, but only for the changes after the last index run: changed mails, and new
          mails of folders that an index run read completely (new mails of other folders: stats.unsynced_skipped);
          attachments and embedding only for the delta (hm_store.DELTA). Deleted and moved mails leave the index.
  status: hello(coverage) -> done. No scan, no reading, no model.
  similar: "Benzerlerini bul": hello -> results(final) -> done. The mails nearest to begin.similar_eid by its stored
          vectors (hm_search.similar). No scan, no reading, no model, no index file for an unknown store.

Every mode first deletes the mails of the folders in begin.exclude (excluded in Ayarlar) from the index file
(Store.purge_folders); done.excluded_removed counts them. An excluded folder in the scope is a protocol error.

Keep-alive (begin.keep_alive_s > 0, while the search window is open): after done(ok) the process waits for the next
begin and keeps the model in memory. preload (status mode): load the model after done when the scope has vectors
(model events). stdin closed or no begin within keep_alive_s: exit 0. Missing/0: exit after the session (old behavior).

Protocol: UTF-8 JSON lines on stdin (add-in -> worker) and stdout (worker -> add-in). See docs/PROTOKOL.md.
No network, no telemetry, no log files. stderr holds only library warnings (no mail or query text).
"""
import argparse
import hashlib
import json
import os
import queue
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

PROTO = 1
QUIET_ENV = {"HF_HUB_OFFLINE": "1", "TRANSFORMERS_OFFLINE": "1", "HF_HUB_DISABLE_TELEMETRY": "1",
             "HF_HUB_DISABLE_PROGRESS_BARS": "1", "TRANSFORMERS_VERBOSITY": "error", "TQDM_DISABLE": "1",
             "TOKENIZERS_PARALLELISM": "false", "DO_NOT_TRACK": "1"}
os.environ.update(QUIET_ENV)
sys.path.insert(0, str(Path(__file__).resolve().parent))

import hm_attach  # noqa: E402
import hm_index  # noqa: E402
import hm_search  # noqa: E402
import hm_store  # noqa: E402
from hm_embed import DIM, MODEL_REVISION  # noqa: E402
from hm_store import INDEX_VERSION, Store  # noqa: E402


class Cancelled(Exception):
    pass


class ProtocolError(Exception):
    pass


class Out:
    def __init__(self, f):
        self.f, self.lock = f, threading.Lock()

    def send(self, ev, **kw):
        line = json.dumps({"ev": ev, **kw}, ensure_ascii=False, separators=(",", ":"))
        with self.lock:
            try:
                self.f.write(line + "\n")
                self.f.flush()
            except (BrokenPipeError, OSError, ValueError):  # the add-in is gone; nobody reads this
                pass


class Inbox:
    """Reads stdin on a thread so that 'cancel' and 'finish' arrive while the main thread computes."""

    def __init__(self, stream):
        self.q = queue.Queue()
        self.stop = threading.Event()       # cancel or finish: stop the embedding loop
        self.cancelled = threading.Event()  # cancel or stdin closed: end the session without results
        threading.Thread(target=self._run, args=(stream,), daemon=True).start()

    def _run(self, stream):
        for line in stream:
            if not line.strip():
                continue
            try:
                msg = json.loads(line)
            except ValueError:
                msg = {"op": "invalid"}
            op = msg.get("op") if isinstance(msg, dict) else "invalid"
            if op == "cancel":
                self.cancelled.set()
                self.stop.set()
                self.q.put(msg)  # wake up a waiting get()
            elif op == "finish":
                self.stop.set()
            else:
                self.q.put(msg)
        self.cancelled.set()  # parent closed the pipe or died
        self.stop.set()
        self.q.put({"op": "eof"})

    def get(self, timeout):
        try:
            msg = self.q.get(timeout=timeout)
        except queue.Empty:
            raise ProtocolError("timeout waiting for the add-in") from None
        if self.cancelled.is_set():
            raise Cancelled()
        return msg


def peak_rss_mb():
    try:
        if os.name == "nt":
            import ctypes
            from ctypes import wintypes

            class PMC(ctypes.Structure):
                _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD)] + [
                    (n, ctypes.c_size_t) for n in ("PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                                                   "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage",
                                                   "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]
            pmc = PMC()
            pmc.cb = ctypes.sizeof(PMC)
            k32 = ctypes.WinDLL("kernel32")
            k32.GetCurrentProcess.restype = wintypes.HANDLE
            # The pseudo handle (-1) comes back as a large int; pass it as HANDLE, not as int (OverflowError).
            h = wintypes.HANDLE(k32.GetCurrentProcess())
            ctypes.WinDLL("psapi").GetProcessMemoryInfo(h, ctypes.byref(pmc), pmc.cb)
            return round(pmc.PeakWorkingSetSize / 2**20, 1)
        import resource
        r = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        return round(r / 2**20 if sys.platform == "darwin" else r / 1024, 1)
    except Exception:  # measurement only
        return None


def db_path(data_dir, store_id):
    return Path(data_dir) / "index" / (hashlib.sha256(store_id.encode("utf-8")).hexdigest()[:24] + ".sqlite")


MODES = ("search", "index", "update", "status", "similar")


def load_encoder(args, enc):
    """The model loads once per process; the next sessions of a kept-alive process reuse it."""
    if "e" not in enc:
        import hm_embed
        enc["e"] = (hm_embed.FakeEncoder if args.fake_encoder else hm_embed.Encoder)(args.model, args.threads)
    return enc["e"]


def session(args, inbox, out, ctx, enc, begin=None):
    """Run one session. ctx receives the store, the coverage function (main() uses them for 'done') and the
    keep-alive values. enc holds the encoder of this process."""
    t0 = time.perf_counter()
    timing, stats = {"model_warm": "e" in enc}, {}
    begin = begin or inbox.get(60)
    if begin.get("op") != "begin":
        raise ProtocolError("begin expected")
    token = os.environ.get("HM_SESSION", "")
    if len(token) < 16 or begin.get("session") != token:
        raise ProtocolError("session id mismatch")
    mode = begin.get("mode") or "search"
    if mode not in MODES:
        raise ProtocolError("unknown mode")
    ctx["keep_alive"] = max(0.0, min(float(begin.get("keep_alive_s") or 0), 86400.0))
    store_info, folders = begin["store"], begin["folders"]
    scope = [f["id"] for f in folders]
    # Folders that the user excluded from the index (Ayarlar): their mails leave the index at the session start, and
    # they are never in a scope. A status check may have an empty scope (every folder excluded): it only deletes.
    exclude = [str(x) for x in begin.get("exclude") or []]
    if set(exclude) & set(scope):
        raise ProtocolError("excluded folder in scope")
    if not scope and mode != "status":
        raise ProtocolError("empty folder scope")
    query = str(begin.get("query") or "").strip()[:500] if mode == "search" else ""
    # Ara tab filters (sender, recipient: keys from done.facets; date_from, date_to: ISO local time; attach: "yes", "no"
    # or a file type). The folder filter is the scope itself.
    filters = begin.get("filters") or {}
    if not isinstance(filters, dict) or any(not isinstance(v, (str, type(None))) for v in filters.values()):
        raise ProtocolError("bad filters")
    if filters.get("attach") and filters["attach"] not in hm_search.ATT_FILTERS:
        raise ProtocolError("bad filters")
    for k in ("date_from", "date_to"):
        if filters.get(k):  # the same text form as items.sent_at: the worker compares the text
            try:
                filters[k] = datetime.fromisoformat(filters[k]).replace(tzinfo=None).isoformat(timespec="seconds")
            except ValueError:
                raise ProtocolError("bad filters") from None
    now = datetime.fromisoformat(begin.get("now") or datetime.now().isoformat(timespec="seconds"))
    top_k = max(1, min(int(begin.get("top_k", 10)), 50))
    # Search mode: attachment reading and embedding stop index_budget_s after the read phase starts (the add-in
    # stops reading new mails itself, after half of it). 0 = no indexing work; missing = no limit (older add-ins,
    # eval). Index mode: no limit.
    budget = begin.get("index_budget_s") if mode == "search" else None
    budget = None if budget is None else max(0.0, float(budget))
    idle = max(5.0, min(float(begin.get("idle_timeout_s", 300)), 3600.0))
    atts_on = bool(begin.get("attachments"))
    path = db_path(args.data, store_info["id"])
    hello = lambda: out.send("hello", pid=os.getpid(), proto=PROTO, index_version=INDEX_VERSION,
                             model_rev=MODEL_REVISION, mode=mode, coverage=ctx["cov"](), model_ready="e" in enc)
    similar_eid = str(begin.get("similar_eid") or "") if mode == "similar" else ""
    if mode == "similar" and not similar_eid:
        raise ProtocolError("similar_eid missing")
    if mode in ("status", "similar"):  # no scan, no reading, no model; an unknown store gets no index file
        store = Store(str(path) if path.exists() else ":memory:", MODEL_REVISION)
        ctx["store"] = store
        removed = store.purge_folders(exclude)
        ctx["cov"] = lambda: store.coverage(scope, atts=atts_on)
        ctx["facets"] = lambda: hm_search.facets(store)
        hello()
        if mode == "similar":  # "Benzerlerini bul": the stored vectors of the mail, no query, no model
            t = time.perf_counter()
            res = hm_search.similar(store, similar_eid, scope, MODEL_REVISION, DIM, top_k,
                                    allow=hm_search.allowed(store, scope, filters))
            out.send("results", stage="final", coverage=ctx["cov"](), **res)
            return {"mode": mode, "excluded_removed": removed, "timing": {"similar_s": round(time.perf_counter() - t, 3)},
                    "model_loaded": "e" in enc}
        # Preload only when a search can use the model: the scope has vectors.
        ctx["preload"] = bool(begin.get("preload")) and ctx["cov"]()["embedded"] > 0
        return {"mode": mode, "excluded_removed": removed}
    # Attachment reading: only when the add-in offers it (it saves the files) and gives its own temp folder.
    att_dir = begin.get("att_dir") if atts_on else None
    if att_dir and not os.path.isdir(att_dir):
        raise ProtocolError("att_dir missing")
    path.parent.mkdir(parents=True, exist_ok=True)
    store = Store(str(path), MODEL_REVISION)
    ctx["store"] = store
    removed = store.purge_folders(exclude)
    with store.tx():
        store.set_meta("store_name", store_info.get("name", ""))
        stats["att_crashed"] = store.fail_stale()
    ctx["cov"] = lambda: store.coverage(scope, atts=atts_on)
    ctx["facets"] = lambda: hm_search.facets(store)
    cov = ctx["cov"]
    hello()
    last_cov = [time.perf_counter()]

    def progress(phase, done, total, t_start, rate=None):
        """Progress with the rate and the time left in this phase (rate = average since t_start); the coverage at
        most every 2 s (the index panel shows it while the session runs)."""
        if rate is None:
            el = time.perf_counter() - t_start
            rate = done / el if el > 0 else 0.0
        extra = {}
        if time.perf_counter() - last_cov[0] >= 2:
            last_cov[0] = time.perf_counter()
            extra["coverage"] = cov()
        eta = 0 if done >= total else round((total - done) / rate) if rate > 0 else None
        out.send("progress", phase=phase, done=done, total=total, rate=round(rate, 2), eta_s=eta, **extra)

    # 1) Scan: the add-in sends the item list of every folder; the worker finds new/changed/deleted items.
    gen = time.time_ns() // 1_000_000
    paths = {f["id"]: f.get("path", "") for f in folders}
    rows, complete, need, mods = {}, [], [], {}
    # Update: a new entry id in a folder that an index run read completely came (or moved in) after it; in another
    # folder it can be any old mail: that folder needs "İndekslemeyi başlat". ponytail: entry ids, no change-time rule
    # (whether a PST move keeps LastModificationTime is not verified).
    synced = store.synced(scope) if mode == "update" else None
    stats["unsynced_skipped"] = 0
    seen = set()  # rowids of the known items in the manifests (Store.sweep)
    t, work = time.perf_counter(), 0.0  # work: the worker part of scan_s (the rest is the wait for the add-in)
    while True:
        msg = inbox.get(idle)
        tw = time.perf_counter()
        op = msg.get("op")
        if op == "manifest":
            fid = msg["folder"]
            if fid not in paths:
                raise ProtocolError("folder not in scope")
            rows.setdefault(fid, []).extend(r for r in msg["items"] if hm_index.searchable(r[2]))
            if msg.get("final"):
                items = rows.pop(fid)
                state = store.item_state([r[0] for r in items])
                changed = [r for r in items if r[0] not in state or state[r[0]][1] != r[1]]
                if synced is not None and fid not in synced:
                    stats["unsynced_skipped"] += sum(1 for r in changed if r[0] not in state)
                    changed = [r for r in changed if r[0] in state]
                known = [state[r[0]] for r in items if r[0] in state]
                seen.update(k[0] for k in known)
                with store.tx():
                    store.move_items([(k[0], fid) for k in known if k[2] != fid])  # folder_id follows moves
                    store.mark_stale([state[r[0]][0] for r in changed if r[0] in state])
                    store.set_folder(fid, paths[fid], len(items))
                need += [(r[3] or "", r[0]) for r in changed]
                mods.update((r[0], r[1]) for r in changed)  # the manifest value is the change key
                complete.append(fid)
                progress("scan", len(complete), len(scope), t)
        elif op == "scan_done":
            break
        else:
            raise ProtocolError("unexpected op during scan: %s" % op)
        work += time.perf_counter() - tw
    tw = time.perf_counter()
    with store.tx():
        stats["removed"] = store.sweep(complete, seen)
    need = [e for _, e in sorted(need, reverse=True)]  # newest first
    timing["scan_s"] = round(time.perf_counter() - t, 3)
    timing["scan_work_s"] = round(work + time.perf_counter() - tw, 3)
    stats.update(manifest_folders=len(complete), need=len(need))
    out.send("need", eids=need)

    # 2) Read: the add-in sends the content of the needed items; the worker stores and keyword-indexes them.
    #    In a time-limited search the add-in stops early (items_done.stopped); the rest stays "to read".
    #    items_done always comes, also after an empty "need".
    t = time.perf_counter()
    deadline = [t + budget] if budget is not None else None  # list: the model load extends it (see get_encoder)
    counts = {"new": 0, "changed": 0, "same": 0, "skipped": 0}
    while True:
        msg = inbox.get(idle)
        op = msg.get("op")
        if op == "items":
            with store.tx():
                for item in msg["items"]:
                    if item.get("folder") not in paths:
                        raise ProtocolError("item folder not in scope")
                    item["mod"] = mods.get(item.get("eid"), item.get("mod"))
                    counts[hm_index.ingest(store, item, gen, MODEL_REVISION)] += 1
            progress("read", sum(counts.values()), len(need), t)
        elif op == "items_done":
            stats["failed"] = len(msg.get("failed", []))
            stats["read_stopped"] = bool(msg.get("stopped"))
            if mode == "index" and not stats["read_stopped"]:
                # Every mail of these folders is in the index now (a failed read stays unknown: "Güncelle" asks for
                # it again). The embedding and attachment backlog is not part of this flag.
                with store.tx():
                    store.set_synced(complete, gen)
            break
        else:
            raise ProtocolError("unexpected op during read: %s" % op)
    stats.update(counts)
    timing["read_s"] = round(time.perf_counter() - t, 3)

    # 3) Keyword results first (no model needed).
    allow = hm_search.allowed(store, scope, filters) if mode == "search" else None  # after the read: new mails count
    if mode == "search":
        t = time.perf_counter()
        res = hm_search.run(store, query, now, scope, top_k=top_k, allow=allow)
        timing["keyword_s"] = round(time.perf_counter() - t, 3)
        out.send("results", stage="keyword", coverage=cov(), **res)

    # The model loads here, inside this process only, or is loaded already.
    def get_encoder():
        if "e" not in enc:
            if inbox.cancelled.is_set():
                raise Cancelled()
            out.send("progress", phase="model", done=0, total=1)
            tl = time.perf_counter()
            load_encoder(args, enc)
            timing["model_load_s"] = round(time.perf_counter() - tl, 3)
            if deadline is not None:  # the search needs the model anyway: its load time is not indexing time
                deadline[0] += timing["model_load_s"]
            out.send("progress", phase="model", done=1, total=1)
        if inbox.cancelled.is_set():  # cancel that arrived during the model load
            raise Cancelled()
        return enc["e"]

    # 3b) Semantic results with the vectors that the index has now, when this search still has attachment or embedding
    #     work (it can take until the deadline). The final results come after that work, as before, and replace these.
    #     The deadline grows by the time of this run: the indexing work of the search stays the same.
    if mode == "search" and query and budget != 0 and cov()["embedded"] and (
            store.missing_vectors(scope, MODEL_REVISION) or (att_dir and store.pending_atts(scope))):
        encoder = get_encoder()
        t = time.perf_counter()
        res = hm_search.run(store, query, now, scope, encoder=encoder, model_rev=MODEL_REVISION, dim=DIM, top_k=top_k,
                            allow=allow)
        out.send("results", stage="semantic", coverage=cov(), **res)
        timing["semantic_s"] = round(time.perf_counter() - t, 3)
        if deadline is not None:
            deadline[0] += time.perf_counter() - t

    # 4) Attachments: the worker asks for a few files at a time; the add-in saves them into att_dir; the worker
    #    reads the text and deletes each file at once. Pull = flow control and a small temp folder.
    t = time.perf_counter()
    stop = lambda: inbox.stop.is_set() or (deadline is not None and time.perf_counter() > deadline[0])
    since = gen - 1 if mode == "update" else None  # update: attachments and embedding of the delta only
    if since is not None:
        stats["delta"] = store.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
            f"SELECT count(*) FROM items i WHERE folder_id IN ({','.join('?' * len(scope))}) AND {hm_store.DELTA}",
            [*scope, since]).fetchone()[0]
    if att_dir:
        stats["att"] = read_attachments(store, inbox, out, scope, att_dir, idle, stop, progress, since)
        timing["att_s"] = round(time.perf_counter() - t, 3)
        if inbox.cancelled.is_set():
            raise Cancelled()

    # 5) Embed new chunks, then the hybrid search.
    t = time.perf_counter()
    # A time-limited search loads the model only for the query, and only when the scope has a semantic index;
    # new chunks are embedded with it until the deadline. Without a semantic index the search gives keyword results
    # and the İndeks tab builds the index (no model load for a few chunks). Chunks with known vectors always finish.
    want_model = mode in ("index", "update") or budget is None or bool(query and cov()["embedded"])
    stats["embed"] = hm_index.embed_pending(store, get_encoder, scope, MODEL_REVISION,
                                            stop=stop if want_model else (lambda: True),
                                            progress=lambda d, n, rate: progress("embed", d, n, t, rate), since=since)
    timing["embed_s"] = round(time.perf_counter() - t, 3)
    if inbox.cancelled.is_set():
        raise Cancelled()
    with store.tx():
        stats["att_texts_removed"] = store.gc_att_texts()  # text of deleted mails leaves the index
        if stats["embed"]["embedded_chunks"] or stats["removed"]:
            stats["vectors_removed"] = store.gc_vectors()

    if mode == "search":
        t = time.perf_counter()
        # No semantic index yet (budget 0 or a new store): keyword results only, the model is not loaded.
        encoder = get_encoder() if query and cov()["embedded"] else None
        res = hm_search.run(store, query, now, scope, encoder=encoder, model_rev=MODEL_REVISION, dim=DIM, top_k=top_k,
                            allow=allow)
        timing["final_s"] = round(time.perf_counter() - t, 3)
        out.send("results", stage="final", coverage=cov(), **res)
    timing["total_s"] = round(time.perf_counter() - t0, 3)
    return {"mode": mode, "timing": timing, "stats": stats, "model_loaded": "e" in enc, "excluded_removed": removed}


ATT_BATCH = 4  # files saved by the add-in per request


def read_attachments(store, inbox, out, scope, att_dir, idle, stop, progress, since=None):
    """Ask for pending attachments in small batches, read each saved file, delete it. Stops on cancel/finish/budget:
    the rest stays 'pending' for the next run. since: only the delta (update mode)."""
    todo = store.pending_atts(scope, since)  # (rowid, entry_id, idx, name, kind)
    counts = {"total": len(todo), "ok": 0, "empty": 0, "failed": 0, "unsupported": 0, "cached": 0}
    base = os.path.realpath(att_dir)
    t = time.perf_counter()
    for b in range(0, len(todo), ATT_BATCH):
        if stop():
            counts["stopped"] = True
            break
        part = todo[b:b + ATT_BATCH]
        out.send("need_att", atts=[[e, i, n] for _, e, i, n, _ in part])
        msg = inbox.get(idle)
        if msg.get("op") != "att_files":
            raise ProtocolError("att_files expected")
        files = {(f[0], int(f[1])): f[2] for f in msg.get("files") or []}  # (eid, idx) -> path or None
        for rowid, eid, idx, name, kind in part:
            path = files.get((eid, idx))
            data = None
            if path:
                full = os.path.realpath(path)
                if os.path.dirname(full) != base:  # only files in the session temp folder are read and deleted
                    raise ProtocolError("attachment path outside att_dir")
                try:
                    with open(full, "rb") as f:
                        data = f.read(hm_attach.MAX_FILE + 1)
                finally:
                    try:
                        os.remove(full)
                    except OSError:
                        pass
            counts[hm_index.apply_file(store, (rowid, idx, name, kind), data)] += 1
        progress("att", min(b + ATT_BATCH, len(todo)), len(todo), t)
    return counts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--model", required=True)
    ap.add_argument("--threads", type=int, default=4)
    ap.add_argument("--fake-encoder", action="store_true", help="tests only: hashed bag-of-words, no model")
    args = ap.parse_args()
    # Protocol channel = original stdout. Anything else that writes to fd 1 goes to stderr.
    proto = os.fdopen(os.dup(1), "w", encoding="utf-8", newline="\n")
    os.dup2(2, 1)
    sys.stdout = sys.stderr
    sys.stdin.reconfigure(encoding="utf-8-sig")  # .NET Process.Start can write a UTF-8 BOM first
    out, inbox = Out(proto), Inbox(sys.stdin)
    code, enc, begin = 0, {}, None

    def done(ctx, status, **kw):
        """'done' carries the index state after the work (the index panel shows it) and the filter values (Ara tab)."""
        extra = {}
        for k in ("cov", "facets"):
            try:
                extra[k] = ctx[k]() if k in ctx else None
            except Exception:  # extra information only; the status counts
                extra[k] = None
        out.send("done", status=status, coverage=extra["cov"], facets=extra["facets"], peak_rss_mb=peak_rss_mb(), **kw)

    while True:  # one session, or more while the add-in keeps this process (keep_alive_s)
        ctx = {}
        try:
            done(ctx, "ok", **session(args, inbox, out, ctx, enc, begin))
        except Cancelled:
            done(ctx, "cancelled")
            code = 3
        except Exception as e:  # report the type only; messages can hold data
            done(ctx, "error", error=type(e).__name__, detail=str(e)[:200] if isinstance(e, ProtocolError) else "")
            code = 2
        try:
            if "store" in ctx:
                ctx["store"].close()
        except Exception:
            pass
        # Cancel and error end the process; so does a session without keep-alive (the old behavior).
        if code or not ctx.get("keep_alive"):
            break
        if ctx.get("preload"):
            preload(args, out, enc, inbox)
        inbox.stop.clear()  # 'finish' was for the session that ended
        try:
            begin = inbox.get(ctx["keep_alive"])
        except (Cancelled, ProtocolError):  # stdin closed, cancel, or no new session in time: a normal end
            break
    try:
        proto.flush()
    except (BrokenPipeError, OSError, ValueError):
        pass
    os._exit(code)  # do not wait for the stdin thread


def preload(args, out, enc, inbox):
    """Load the model between sessions (the search window opened; the scope has vectors), so that the next search
    does not wait for it. A begin that arrives meanwhile waits in the inbox."""
    if inbox.cancelled.is_set():
        return
    t = time.perf_counter()
    if "e" not in enc:
        out.send("model", state="loading")
        try:
            load_encoder(args, enc)
        except Exception as e:  # the next search tries again and reports its own error
            out.send("model", state="error", error=type(e).__name__)
            return
    out.send("model", state="ready", load_s=round(time.perf_counter() - t, 3), peak_rss_mb=peak_rss_mb())


if __name__ == "__main__":
    main()
