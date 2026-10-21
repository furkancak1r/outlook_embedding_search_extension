"""Per-store SQLite index: items, chunks, FTS5 (contentless), float16 vectors keyed by chunk-text hash,
attachments and their extracted text.

One database file per Outlook store (account separation). Only one worker process opens it at a time.
"""
import json
import sqlite3
import time

import numpy as np

INDEX_VERSION = 2  # 2: attachment content (atts, att_texts, fts_att, chunks.src)
SCHEMA = """
PRAGMA journal_mode=WAL;
PRAGMA synchronous=NORMAL;
CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS folders(id TEXT PRIMARY KEY, path TEXT, total INTEGER, scanned_at REAL);
CREATE TABLE IF NOT EXISTS items(
  rowid INTEGER PRIMARY KEY, entry_id TEXT UNIQUE NOT NULL, folder_id TEXT, msg_class TEXT,
  last_mod TEXT, content_hash TEXT, subject TEXT, sender TEXT, sender_addr TEXT, recipients TEXT,
  sent_at TEXT, attachments TEXT, main_text TEXT, n_chunks INTEGER, embedded INTEGER DEFAULT 0,
  model_rev TEXT, index_version INTEGER, seen_gen INTEGER DEFAULT 0);
CREATE INDEX IF NOT EXISTS items_folder ON items(folder_id, seen_gen);
CREATE TABLE IF NOT EXISTS chunks(
  id INTEGER PRIMARY KEY, item_rowid INTEGER NOT NULL, seq INTEGER, start INTEGER, end INTEGER, text_hash TEXT,
  src INTEGER NOT NULL DEFAULT 0);
CREATE INDEX IF NOT EXISTS chunks_item ON chunks(item_rowid);
CREATE INDEX IF NOT EXISTS chunks_hash ON chunks(text_hash);
CREATE TABLE IF NOT EXISTS vectors(text_hash TEXT PRIMARY KEY, model_rev TEXT, dim INTEGER, vec BLOB);
CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(
  subject, people, body, quoted, attach, content='', contentless_delete=1,
  tokenize='unicode61 remove_diacritics 2', prefix='2 3 4');
CREATE VIRTUAL TABLE IF NOT EXISTS fts_vocab USING fts5vocab(fts, 'row');
CREATE TABLE IF NOT EXISTS atts(
  item_rowid INTEGER NOT NULL, idx INTEGER NOT NULL, name TEXT, size INTEGER, kind INTEGER, state TEXT,
  file_hash TEXT, PRIMARY KEY(item_rowid, idx));
CREATE INDEX IF NOT EXISTS atts_state ON atts(state);
CREATE TABLE IF NOT EXISTS att_texts(file_hash TEXT PRIMARY KEY, state TEXT, text TEXT);
CREATE VIRTUAL TABLE IF NOT EXISTS fts_att USING fts5(
  body, content='', contentless_delete=1, tokenize='unicode61 remove_diacritics 2', prefix='2 3 4');
"""
ITEM_COLS = ["rowid", "entry_id", "folder_id", "msg_class", "subject", "sender", "sender_addr", "recipients",
             "sent_at", "attachments", "main_text", "embedded"]
# Attachment states. pending: to read; extracting: being read now (a worker that dies here leaves it, and the
# next worker marks it failed: a file that crashes the reader is not tried again); ok/empty/failed: read;
# unsupported/too_big: not read.
ATT_READABLE = ("pending", "extracting", "ok", "empty", "failed")
ATT_DONE = ("ok", "empty", "failed")


# The SQL that _in() builds holds only '?' placeholders (and constant column names): values are always bound.
# Semgrep flags these f-strings as raw SQL; that is a false positive (sqlite3, not SQLAlchemy).
def _in(n):
    return ",".join("?" * n)


class Store:
    def __init__(self, path, model_rev):
        self.db = sqlite3.connect(path, timeout=30, isolation_level=None)
        self.db.executescript(SCHEMA)
        self._vocab = None
        with self.tx():
            if "src" not in {r[1] for r in self.db.execute("PRAGMA table_info(chunks)")}:  # index version 1
                self.db.execute("ALTER TABLE chunks ADD COLUMN src INTEGER NOT NULL DEFAULT 0")
            version = self.meta("index_version")
            if version is None:
                self.set_meta("index_version", INDEX_VERSION)
                self.set_meta("created_at", time.time())
            elif version < INDEX_VERSION:
                # The add-in reads every mail again once (attachment list with sizes). Vectors are kept: the
                # same chunk text has the same hash, so only new attachment chunks need the model.
                # last_mod NULL = "to read" (see mark_stale).
                self.db.execute("UPDATE items SET last_mod=NULL")
                self.set_meta("index_version", INDEX_VERSION)
            if self.meta("model_rev") != model_rev:  # new model: every item needs new vectors
                self.db.execute("UPDATE items SET embedded=0, model_rev=?", (model_rev,))
                self.set_meta("model_rev", model_rev)

    def tx(self):
        store = self

        class _Tx:
            def __enter__(self):
                store.db.execute("BEGIN IMMEDIATE")

            def __exit__(self, exc, *_):
                store.db.execute("ROLLBACK" if exc else "COMMIT")

        return _Tx()

    def meta(self, k, default=None):
        r = self.db.execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
        return json.loads(r[0]) if r else default

    def set_meta(self, k, v):
        self.db.execute("INSERT OR REPLACE INTO meta(k, v) VALUES(?, ?)", (k, json.dumps(v)))

    # -- sync ---------------------------------------------------------------------------------
    def item_state(self, entry_ids):
        """{entry_id: (rowid, last_mod)} for known ids."""
        out = {}
        for i in range(0, len(entry_ids), 500):
            part = entry_ids[i:i + 500]
            out.update((e, (r, m)) for e, r, m in self.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
                f"SELECT entry_id, rowid, last_mod FROM items WHERE entry_id IN ({_in(len(part))})", part))
        return out

    def mark_seen(self, rows, gen):
        """rows: [(rowid, folder_id)] present in the current manifest (folder_id follows moves)."""
        self.db.executemany("UPDATE items SET seen_gen=?, folder_id=? WHERE rowid=?", [(gen, f, r) for r, f in rows])

    def mark_stale(self, rowids):
        """Changed in Outlook, not read again yet: the item stays searchable with its old text, counts as not read,
        and the next scan asks for it again (a time-limited search can stop before it)."""
        self.db.executemany("UPDATE items SET last_mod=NULL WHERE rowid=?", [(r,) for r in rowids])

    def sweep(self, folder_ids, gen):
        """Delete items of completely listed folders that the manifest did not contain (deleted / moved out)."""
        if not folder_ids:
            return 0
        rows = [r for (r,) in self.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
            f"SELECT rowid FROM items WHERE folder_id IN ({_in(len(folder_ids))}) AND seen_gen<>?", [*folder_ids, gen])]
        for r in rows:
            self.delete_item(r)
        return len(rows)

    def delete_item(self, rowid):
        """The extracted attachment text stays in att_texts (keyed by file hash) until gc_att_texts() at the end of
        the session: a mail moved in a PST store gets a new entry id, and its attachments are not read again."""
        self.db.execute("DELETE FROM fts WHERE rowid=?", (rowid,))
        self.db.execute("DELETE FROM fts_att WHERE rowid=?", (rowid,))
        self.db.execute("DELETE FROM chunks WHERE item_rowid=?", (rowid,))
        self.db.execute("DELETE FROM atts WHERE item_rowid=?", (rowid,))
        self.db.execute("DELETE FROM items WHERE rowid=?", (rowid,))
        self._vocab = None

    def upsert_item(self, rec, fts_fields, spans, chunk_hashes, model_rev, gen, att_rows=()):
        """att_rows: [(idx, name, size, kind, state)]."""
        old = self.db.execute("SELECT rowid FROM items WHERE entry_id=?", (rec["entry_id"],)).fetchone()
        if old:
            self.delete_item(old[0])
        rowid = self.db.execute(
            "INSERT INTO items(entry_id, folder_id, msg_class, last_mod, content_hash, subject, sender, sender_addr,"
            " recipients, sent_at, attachments, main_text, n_chunks, embedded, model_rev, index_version, seen_gen)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,0,?,?,?)",
            (rec["entry_id"], rec["folder_id"], rec["msg_class"], rec["last_mod"], rec["content_hash"], rec["subject"],
             rec["sender"], rec["sender_addr"], rec["recipients"], rec["sent_at"], rec["attachments"],
             rec["main_text"], len(spans), model_rev, INDEX_VERSION, gen)).lastrowid
        self.db.execute("INSERT INTO fts(rowid, subject, people, body, quoted, attach) VALUES(?,?,?,?,?,?)",
                        (rowid, *fts_fields))
        self.db.executemany("INSERT INTO chunks(item_rowid, seq, start, end, text_hash) VALUES(?,?,?,?,?)",
                            [(rowid, i, a, b, h) for i, ((a, b), h) in enumerate(zip(spans, chunk_hashes))])
        self.db.executemany("INSERT OR REPLACE INTO atts(item_rowid, idx, name, size, kind, state) VALUES(?,?,?,?,?,?)",
                            [(rowid, *a) for a in att_rows])
        self._vocab = None
        return rowid

    def purge_folders(self, folder_ids):
        """Delete the mails of folders that the user excluded from the index (Ayarlar), with their keyword rows,
        chunks, attachments, and the attachment text and vectors that no other mail uses. The text must also leave the
        file: FTS5 keeps deleted words in its segments until a merge (optimize), and SQLite keeps freed pages until
        VACUUM (10-20 probe: VACUUM without optimize left the words in the file). meta 'compact' makes the next session
        do an interrupted VACUUM again. Returns the number of deleted mails."""
        ids = list(folder_ids)
        rows = []
        if ids:
            with self.tx():
                rows = [r for (r,) in self.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
                    f"SELECT rowid FROM items WHERE folder_id IN ({_in(len(ids))})", ids)]
                for r in rows:
                    self.delete_item(r)
                self.db.execute(f"DELETE FROM folders WHERE id IN ({_in(len(ids))})", ids)  # nosemgrep: sqlalchemy-execute-raw-query
                if rows:
                    self.gc_att_texts()
                    self.gc_vectors()
                    self.db.execute("INSERT INTO fts(fts) VALUES('optimize')")
                    self.db.execute("INSERT INTO fts_att(fts_att) VALUES('optimize')")
                    self.set_meta("compact", True)
        if self.meta("compact"):
            self.db.execute("VACUUM")
            self.set_meta("compact", False)
            self.db.execute("PRAGMA wal_checkpoint(TRUNCATE)")  # the old pages leave the WAL file too
        return len(rows)

    def gc_vectors(self):
        return self.db.execute("DELETE FROM vectors WHERE text_hash NOT IN (SELECT text_hash FROM chunks)").rowcount

    def gc_att_texts(self):
        """Delete extracted text that no attachment uses any more (deleted mails: their text leaves the index)."""
        return self.db.execute("DELETE FROM att_texts WHERE file_hash NOT IN"
                               " (SELECT file_hash FROM atts WHERE file_hash IS NOT NULL)").rowcount

    # -- attachments --------------------------------------------------------------------------
    def fail_stale(self):
        """'extracting' rows of a worker that died (crash or memory limit while it read the file) become 'failed'.
        The file hash is remembered as failed, so the same file in another mail is not tried again."""
        self.db.execute("INSERT OR IGNORE INTO att_texts(file_hash, state, text) SELECT DISTINCT file_hash,"
                        " 'failed', '' FROM atts WHERE state='extracting' AND file_hash IS NOT NULL")
        return self.db.execute("UPDATE atts SET state='failed' WHERE state='extracting'").rowcount

    def pending_atts(self, folder_ids):
        """[(item_rowid, entry_id, idx, name, kind)] to read, newest mails first."""
        return self.db.execute(
            "SELECT a.item_rowid, i.entry_id, a.idx, a.name, a.kind FROM atts a JOIN items i ON i.rowid=a.item_rowid"
            f" WHERE a.state='pending' AND i.folder_id IN ({_in(len(folder_ids))}) ORDER BY i.sent_at DESC, a.idx",
            list(folder_ids)).fetchall()

    def set_att(self, rowid, idx, state, file_hash=None):
        self.db.execute("UPDATE atts SET state=?, file_hash=? WHERE item_rowid=? AND idx=?", (state, file_hash, rowid, idx))

    def cached_att(self, file_hash):
        """(state, text) of a file that was read before (same content in another mail or a moved mail), or None."""
        return self.db.execute("SELECT state, text FROM att_texts WHERE file_hash=?", (file_hash,)).fetchone()

    def put_att_text(self, file_hash, state, text):
        self.db.execute("INSERT OR REPLACE INTO att_texts(file_hash, state, text) VALUES(?,?,?)", (file_hash, state, text))

    def att_texts(self, rowid):
        """[(idx, name, text)] of the attachments of an item with text, in attachment order."""
        return self.db.execute(
            "SELECT a.idx, a.name, t.text FROM atts a JOIN att_texts t ON t.file_hash=a.file_hash"
            " WHERE a.item_rowid=? AND a.state='ok' ORDER BY a.idx", (rowid,)).fetchall()

    def set_att_index(self, rowid, folded_text, chunk_rows):
        """Replace the attachment part of an item: its fts_att row and its attachment chunks.
        chunk_rows: [(src = attachment idx, start, end, text_hash)]. The item needs embedding again
        (body chunks keep their vectors: same text hash)."""
        self.db.execute("DELETE FROM fts_att WHERE rowid=?", (rowid,))
        if folded_text:
            self.db.execute("INSERT INTO fts_att(rowid, body) VALUES(?, ?)", (rowid, folded_text))
        self.db.execute("DELETE FROM chunks WHERE item_rowid=? AND src<>0", (rowid,))
        first = self.db.execute("SELECT coalesce(max(seq), -1) + 1 FROM chunks WHERE item_rowid=?", (rowid,)).fetchone()[0]
        self.db.executemany("INSERT INTO chunks(item_rowid, seq, start, end, text_hash, src) VALUES(?,?,?,?,?,?)",
                            [(rowid, first + i, a, b, h, s) for i, (s, a, b, h) in enumerate(chunk_rows)])
        self.db.execute("UPDATE items SET embedded=0 WHERE rowid=?", (rowid,))
        self._vocab = None

    # -- embeddings ---------------------------------------------------------------------------
    def pending(self, folder_ids):
        return self.db.execute(
            f"SELECT rowid, subject, sender, attachments, main_text FROM items WHERE embedded=0"
            f" AND folder_id IN ({_in(len(folder_ids))}) ORDER BY sent_at DESC", list(folder_ids)).fetchall()

    def chunk_rows(self, rowid):
        """[(src, start, end, text_hash)] in chunk order."""
        return self.db.execute("SELECT src, start, end, text_hash FROM chunks WHERE item_rowid=? ORDER BY seq",
                               (rowid,)).fetchall()

    def chunk(self, rowid, seq):
        """(src, start, end) of one chunk."""
        return self.db.execute("SELECT src, start, end FROM chunks WHERE item_rowid=? AND seq=?", (rowid, seq)).fetchone()

    def has_vectors(self, hashes, model_rev):
        out = set()
        for i in range(0, len(hashes), 500):
            part = hashes[i:i + 500]
            out.update(h for (h,) in self.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
                f"SELECT text_hash FROM vectors WHERE model_rev=? AND text_hash IN ({_in(len(part))})", [model_rev, *part]))
        return out

    def put_vectors(self, rows, model_rev):
        self.db.executemany("INSERT OR REPLACE INTO vectors(text_hash, model_rev, dim, vec) VALUES(?,?,?,?)",
                            [(h, model_rev, len(v), v.astype(np.float16).tobytes()) for h, v in rows])

    def set_embedded(self, rowids):
        self.db.executemany("UPDATE items SET embedded=1 WHERE rowid=?", [(r,) for r in rowids])

    def all_vectors(self, model_rev, dim, folder_ids):
        """(item_rowids int64, chunk_seq int32, matrix float32 [N, dim]) for the folder scope."""
        rows = self.db.execute(
            "SELECT c.item_rowid, c.seq, v.vec FROM items i JOIN chunks c ON c.item_rowid=i.rowid"
            " JOIN vectors v ON v.text_hash=c.text_hash"
            f" WHERE v.model_rev=? AND v.dim=? AND i.folder_id IN ({_in(len(folder_ids))})",
            [model_rev, dim, *folder_ids]).fetchall()
        n = len(rows)
        ids = np.fromiter((r[0] for r in rows), np.int64, n)
        seq = np.fromiter((r[1] for r in rows), np.int32, n)
        mat = np.frombuffer(b"".join(r[2] for r in rows), np.float16).reshape(n, dim).astype(np.float32)
        return ids, seq, mat

    # -- keyword search -----------------------------------------------------------------------
    def vocab(self):
        """{term: document count} of the mail FTS index (loaded once per worker run).
        ponytail: attachment words are not in it (spreadsheets bring many codes and numbers; the typo scan is
        linear). A typo in a word that is only in an attachment is not corrected; the exact word still matches."""
        if self._vocab is None:
            self._vocab = dict(self.db.execute("SELECT term, doc FROM fts_vocab"))
        return self._vocab

    def fts(self, match, limit=400):
        try:
            return self.db.execute(
                "SELECT rowid, bm25(fts, 6.0, 4.0, 1.0, 0.3, 3.0) AS s FROM fts WHERE fts MATCH ? ORDER BY s LIMIT ?",
                (match, limit)).fetchall()
        except sqlite3.OperationalError:
            return []

    def fts_att(self, match, limit=400):
        """Like fts(), on the attachment text."""
        try:
            return self.db.execute("SELECT rowid, bm25(fts_att) AS s FROM fts_att WHERE fts_att MATCH ? ORDER BY s LIMIT ?",
                                   (match, limit)).fetchall()
        except sqlite3.OperationalError:
            return []

    def fts_att_rowids(self, match):
        try:
            return {r for (r,) in self.db.execute("SELECT rowid FROM fts_att WHERE fts_att MATCH ?", (match,))}
        except sqlite3.OperationalError:
            return set()

    def scope_rowids(self, folder_ids):
        return {r for (r,) in self.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
            f"SELECT rowid FROM items WHERE folder_id IN ({_in(len(folder_ids))})", list(folder_ids))}

    def items(self, rowids):
        rowids = list(rowids)
        out = {}
        for i in range(0, len(rowids), 500):
            part = rowids[i:i + 500]
            out.update((r[0], dict(zip(ITEM_COLS, r))) for r in self.db.execute(  # nosemgrep: sqlalchemy-execute-raw-query
                f"SELECT {', '.join(ITEM_COLS)} FROM items WHERE rowid IN ({_in(len(part))})", part))
        return out

    # -- coverage -----------------------------------------------------------------------------
    def set_folder(self, fid, path, total):
        self.db.execute("INSERT OR REPLACE INTO folders(id, path, total, scanned_at) VALUES(?,?,?,?)",
                        (fid, path, total, time.time()))

    def coverage(self, folder_ids, atts=True):
        """Index state per folder (in folder_ids order) and in total. total: searchable mails in the last scan
        (None: never scanned); indexed: mails in the index; read: of them, read with their current version;
        embedded: in the semantic index; att_total/att_done: readable attachments of indexed mails / read ones.
        atts False: attachments are not read (setting off): not counted."""
        q, ids = _in(len(folder_ids)), list(folder_ids)
        meta = {r[0]: r[1:] for r in self.db.execute(f"SELECT id, path, total, scanned_at FROM folders WHERE id IN ({q})", ids)}  # nosemgrep: sqlalchemy-execute-raw-query
        items = {r[0]: r[1:] for r in self.db.execute(
            "SELECT folder_id, count(*), coalesce(sum(last_mod IS NOT NULL), 0), coalesce(sum(embedded), 0) FROM items"
            f" WHERE folder_id IN ({q}) GROUP BY folder_id", ids)}
        att = {r[0]: r[1:] for r in self.db.execute(
            f"SELECT i.folder_id, count(*), coalesce(sum(a.state IN ({_in(len(ATT_DONE))})), 0) FROM atts a"
            f" JOIN items i ON i.rowid=a.item_rowid WHERE a.state IN ({_in(len(ATT_READABLE))})"
            f" AND i.folder_id IN ({q}) GROUP BY i.folder_id", [*ATT_DONE, *ATT_READABLE, *ids])} if atts else {}
        folders = []
        for fid in dict.fromkeys(ids):
            path, total, scanned = meta.get(fid, (None, None, None))
            n, rd, emb = items.get(fid, (0, 0, 0))
            at, ad = att.get(fid, (0, 0))
            folders.append({"id": fid, "path": path, "total": total, "indexed": n, "read": rd, "embedded": emb,
                            "att_total": at, "att_done": ad, "scanned_at": scanned})
        out = {k: sum(f[k] for f in folders) for k in ("indexed", "read", "embedded", "att_total", "att_done")}
        out.update(folders=folders, total=sum(f["total"] or 0 for f in folders), complete=all(
            f["total"] is not None and f["read"] >= f["total"] and f["embedded"] >= f["indexed"]
            and f["att_done"] >= f["att_total"] for f in folders))
        return out

    def close(self):
        try:
            self.db.execute("PRAGMA optimize")
        finally:
            self.db.close()
