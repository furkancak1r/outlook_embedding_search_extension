"""Indexing: mail dict -> stored record, FTS text, chunk texts; attachment text; embedding of pending chunks."""
import hashlib
import re
import time

import hm_attach
import hm_embed
from hm_text import chunk_spans, content_hash, fold, split_body, text_hash

SUBJ_PREFIX = re.compile(r"^\s*(?:(?:re|fw|fwd|ynt|ilt|aw|wg|sv|vs|antw|cevap|yanit|yanıt)\s*(?:\[\d+\])?\s*:\s*)+", re.I)
NOISE_ATTACH = re.compile(r"^(?:image|outlook|att)\d*\.(?:png|jpe?g|gif|bmp)$", re.I)
SEARCHABLE = ("ipm.note", "ipm.schedule.meeting", "ipm.post")
MAX_BODY = 200_000
# An inline picture (image001.png) smaller than this is a logo or an icon: not read. A bigger one is a pasted
# screenshot or photo: read with OCR. ponytail: size rule, no image analysis. A big logo costs one OCR run per
# distinct file (the text is kept by file hash).
INLINE_MIN = 40_000
MAX_ATT_CHUNKS = 3          # embedded chunks per attachment; the full text stays keyword-searchable
MAX_ITEM_ATT_CHUNKS = 12    # embedded attachment chunks per mail
MAX_ITEM_ATT_TEXT = 300_000  # attachment chars per mail in the keyword index


def searchable(cls):
    c = (cls or "").lower()
    return c.startswith(SEARCHABLE) and not c.startswith(("ipm.note.rules", "ipm.note.storage"))


def clean_subject(s):
    return SUBJ_PREFIX.sub("", s or "").strip()


def chunk_texts(subject, sender, attachments, main, spans):
    """Model input for each body chunk. The first chunk also names the sender and the attachments."""
    title = clean_subject(subject)
    head = "Kimden: " + (sender or "?") + (" | Ekler: " + attachments if attachments else "")
    return [hm_embed.doc_text(title, (head + "\n" if i == 0 else "") + main[a:b]) for i, (a, b) in enumerate(spans)]


def att_chunk_text(subject, name, text, span):
    """Model input for an attachment chunk: the mail subject and the file name are the title."""
    return hm_embed.doc_text(clean_subject(subject) + " | Ek: " + (name or "?"), text[span[0]:span[1]])


def att_state(name, size, kind):
    """First state of an attachment: pending (to read) or a reason not to read it."""
    if NOISE_ATTACH.match(name or "") and (size or 0) < INLINE_MIN:
        return "skipped"
    if not hm_attach.readable(name, kind):
        return "unsupported"
    if (size or 0) > hm_attach.MAX_FILE:
        return "too_big"
    return "pending"


def prepare(item):
    subject = item.get("subject") or ""
    body = (item.get("body") or "")[:MAX_BODY]
    main, quoted = split_body(body, subject)
    atts = "; ".join(a for a in item.get("attachments") or [] if a and not NOISE_ATTACH.match(a))
    recipients = "; ".join(x for x in (item.get("to"), item.get("cc")) if x)
    # att_meta: [[index (1-based), name, size, Outlook attachment type]]. Missing in v1 add-ins and in the
    # name-only test corpus: then no attachment is read.
    meta = [(int(m[0]), str(m[1] or ""), int(m[2] or 0), int(m[3] or 0)) for m in item.get("att_meta") or []]
    rec = {
        "entry_id": item["eid"], "folder_id": item["folder"], "msg_class": item.get("cls") or "",
        "last_mod": item.get("mod") or "", "subject": subject, "sender": item.get("sender") or "",
        "sender_addr": item.get("sender_addr") or "", "recipients": recipients, "sent_at": item.get("date") or "",
        "attachments": atts, "main_text": main,
        "content_hash": content_hash([subject, item.get("sender"), item.get("sender_addr"), recipients,
                                      item.get("date"), atts, body, meta]),
    }
    people = " ; ".join(x for x in (rec["sender"], rec["sender_addr"], recipients) if x)
    fts = (fold(subject), fold(people), fold(main), fold(quoted[:20000]), fold(atts))
    spans = chunk_spans(main)
    hashes = [text_hash(t) for t in chunk_texts(subject, rec["sender"], atts, main, spans)]
    att_rows = [(i, n, s, k, att_state(n, s, k)) for i, n, s, k in meta]
    return rec, fts, spans, hashes, att_rows


def ingest(store, item, gen, model_rev):
    """Store one mail. Returns 'skipped', 'same', 'changed' or 'new'. The caller owns the transaction."""
    if not searchable(item.get("cls")):
        return "skipped"
    rec, fts, spans, hashes, att_rows = prepare(item)
    old = store.db.execute("SELECT rowid, content_hash FROM items WHERE entry_id=?", (rec["entry_id"],)).fetchone()
    if old and old[1] == rec["content_hash"]:  # only flags/read state changed
        store.db.execute("UPDATE items SET last_mod=?, folder_id=?, seen_gen=? WHERE rowid=?",
                         (rec["last_mod"], rec["folder_id"], gen, old[0]))
        return "same"
    store.upsert_item(rec, fts, spans, hashes, model_rev, gen, att_rows)
    return "changed" if old else "new"


# --- attachments ---------------------------------------------------------------------------------

def apply_file(store, att, data):
    """Read one saved attachment into the index. att = (item rowid, idx, name, kind); data = file bytes, or None
    when the add-in could not save the file. Returns the new state, or 'cached' when the same file (sha256) was
    read before. Each step commits, so a crash inside a reader leaves the state 'extracting' (Store.fail_stale)."""
    rowid, idx, name, kind = att
    if data is None:
        with store.tx():
            store.set_att(rowid, idx, "failed")
        return "failed"
    h = hashlib.sha256(data).hexdigest()[:32]
    with store.tx():
        cached = store.cached_att(h)
        if cached:
            store.set_att(rowid, idx, cached[0], h)
            index_attachments(store, rowid)
            return "cached"
        store.set_att(rowid, idx, "extracting", h)
    try:
        state, text = hm_attach.extract_bytes(data, name, kind)
    except Exception:  # corrupt, encrypted or strange file: record it, never stop the session
        state, text = "failed", ""
    with store.tx():
        if state != "unsupported":  # "unsupported" is not kept by hash: a later version can read it
            store.put_att_text(h, state, text)
        store.set_att(rowid, idx, state, h)
        index_attachments(store, rowid)
    return state


def index_attachments(store, rowid):
    """Rebuild the attachment part of one mail (keyword row and attachment chunks) from its read attachments.
    The caller owns the transaction."""
    subject = store.db.execute("SELECT subject FROM items WHERE rowid=?", (rowid,)).fetchone()[0]
    parts, rows, size = [], [], 0
    for idx, name, text in store.att_texts(rowid):
        text = text[:max(0, MAX_ITEM_ATT_TEXT - size)]
        if not text:
            break
        size += len(text)
        parts.append(fold(text))
        for a, b in chunk_spans(text, max_chunks=MAX_ATT_CHUNKS)[:max(0, MAX_ITEM_ATT_CHUNKS - len(rows))]:
            rows.append((idx, a, b, text_hash(att_chunk_text(subject, name, text, (a, b)))))
    store.set_att_index(rowid, "\n".join(parts), rows)


# --- embedding -----------------------------------------------------------------------------------

def item_chunk_texts(store, rowid, subject, sender, atts, main):
    """[(text_hash, model input)] of all chunks of one mail, body and attachments."""
    rows = store.chunk_rows(rowid)  # (src, start, end, hash)
    body = iter(chunk_texts(subject, sender, atts, main, [(a, b) for s, a, b, _ in rows if s == 0]))
    files = {i: (n, t) for i, n, t in store.att_texts(rowid)} if any(s for s, _, _, _ in rows) else {}
    out = []
    for s, a, b, h in rows:
        if s == 0:
            out.append((h, next(body)))
        elif s in files:
            out.append((h, att_chunk_text(subject, files[s][0], files[s][1], (a, b))))
    return out


def embed_pending(store, get_encoder, folder_ids, model_rev, batch=8, window=64, stop=lambda: False,
                  progress=lambda done, total, rate: None, since=None):
    """Embed the chunks of not-yet-embedded items in the scope, newest first. since: only the delta ("Güncelle").

    Vectors are saved after every model batch (a killed process loses one batch). An item is marked embedded when
    every chunk has a vector, also after stop(). Chunks with a known text hash reuse the stored vector (moved, re-read
    or re-saved mails cost nothing); only the model work waits for the next run.
    progress(done, total, rate) counts model work only: chunks embedded / distinct chunks without a vector;
    rate = chunks per second since the first batch (the model load is not counted). Returns a stats dict.
    get_encoder() loads the model only if there is work.
    """
    items = store.pending(folder_ids, since)
    todo = [(rowid, item_chunk_texts(store, rowid, subject, sender, atts, main))
            for rowid, subject, sender, atts, main in items]
    distinct = list({h for _, cs in todo for h, _ in cs})
    total = len(distinct) - len(store.has_vectors(distinct, model_rev))
    embedded, reused, stopped, t0, t_enc = 0, 0, False, time.perf_counter(), None
    rate = lambda: embedded / max(1e-9, time.perf_counter() - t_enc) if t_enc else 0.0
    i = 0
    while i < len(todo):
        group, n = [], 0
        while i < len(todo) and (n < window or not group):
            group.append(todo[i])
            n += len(todo[i][1])
            i += 1
        known = store.has_vectors([h for _, cs in group for h, _ in cs], model_rev)
        need = {h: t for _, cs in group for h, t in cs if h not in known}
        reused += sum(1 for _, cs in group for h, _ in cs if h in known)
        if need and not stopped and not stop():
            enc = get_encoder()
            t_enc = t_enc or time.perf_counter()
            pairs = sorted(need.items(), key=lambda kv: len(kv[1]))  # similar lengths: less padding
            for b in range(0, len(pairs), batch):
                if stop():  # also right after the model load
                    stopped = True
                    break
                part = pairs[b:b + batch]
                vecs = enc.encode([t for _, t in part])
                with store.tx():
                    store.put_vectors(list(zip([h for h, _ in part], vecs)), model_rev)
                known.update(h for h, _ in part)
                embedded += len(part)
                progress(embedded, total, rate())
        elif need:
            stopped = True  # model work waits for the next run
        # Items with a vector for every chunk count as embedded, also when the run stopped inside this window.
        ready = [r for r, cs in group if all(h in known for h, _ in cs)]
        if ready:
            with store.tx():
                store.set_embedded(ready)
    if total:
        progress(embedded, total, rate())
    return {"items": len(items), "chunks": sum(len(c) for _, c in todo), "model_chunks": total,
            "embedded_chunks": embedded, "reused_chunks": reused, "stopped": stopped,
            "seconds": round(time.perf_counter() - t0, 3)}
