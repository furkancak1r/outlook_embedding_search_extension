"""Hybrid search: FTS5 keyword (Turkish folding, light stemming, typo expansion) + vector cosine + soft date boost."""
import re
from datetime import datetime

import numpy as np

import hm_dates
from hm_text import fold, tokens

STOP = set(fold(w) for w in """
ve veya ile bir bu şu o da de mi mı mu mü ki için gibi olan olarak hakkında hakkındaki ilgili konu konusu konulu
gelen giden gönderilen gönderdiği göndermişti göndermiş gönderdi gönderen yollamıştı yollamış yollanan yazmıştı yazmış
attığı mail maili mailde mailini mailı e-posta eposta posta postası postayı mesaj mesajı mesajını ileti iletisi galiba
sanırım herhalde belki bana beni benim ben biz bize bizim var vardı yoktu olabilir olmalı neydi ne nedir hangi hani
şey şeyi şeyle dair üzerine ait içinde içeren içeriyordu olduğu aramıştım arıyorum bul bulur musun lütfen bulamıyorum
the a an of to from about with for in on at by some something that this mail email e-mail message sent regarding
i me my we our it was is were think maybe there had has have who what which any find please one thing someone
""".split())
SYNONYMS = {"excel": ["xlsx", "xls", "csv"], "word": ["docx", "doc"], "powerpoint": ["pptx", "ppt"],
            "sunum": ["pptx", "ppt"], "tablo": ["xlsx", "xls"], "spreadsheet": ["xlsx", "xls"],
            "resim": ["jpg", "jpeg", "png"], "fotograf": ["jpg", "jpeg", "png"], "photo": ["jpg", "jpeg", "png"],
            "zip": ["zip", "rar", "7z"]}
TR_SUFFIXES = sorted("""lar ler larin lerin lari leri larda lerde lardan lerden lara lere nin nun in un si su
da de ta te dan den tan ten ndan nden nda nde yla yle la le ya ye yi yu ki dir dur tir tur mis mus i u a e
siyla siyle sini sunu sine suna sinde sunda""".split(), key=len, reverse=True)
EN_SUFFIXES = ["ing", "ed", "es", "s"]
# Fusion weights and confidence thresholds. Chosen by grid search on the odd-numbered (dev) synthetic queries
# with eval/calibrate.py and checked on the even-numbered (test) queries.
# score = cosine(full query) + W_COV * keyword coverage * min(1, words/2) + W_DATE * date score + W_KW / (10 + BM25 rank)
W_COV, W_DATE, W_KW = 0.08, 0.04, 0.5
W_ATT = 0.5  # attachment BM25 rank relative to the mail BM25 rank (not calibrated: chosen, tested on 13 attachment queries)
COS_HIGH, COS_HIGH_Z, Z_HIGH, COS_MEDIUM, COV_HIGH, COV_MEDIUM, COV_FULL = 0.80, 0.75, 4.5, 0.70, 0.6, 0.5, 0.99


def stem(t):
    """Light suffix stripping on a folded token (keeps at least 4 chars)."""
    for _ in range(2):
        for suf in TR_SUFFIXES + EN_SUFFIXES:
            if t.endswith(suf) and len(t) - len(suf) >= 4:
                t = t[: -len(suf)]
                break
        else:
            break
    return t


def edit_distance_le(a, b, k):
    """Optimal string alignment (Damerau-Levenshtein) distance <= k."""
    if abs(len(a) - len(b)) > k:
        return False
    prev2, prev = None, list(range(len(b) + 1))
    for i in range(1, len(a) + 1):
        cur = [i] + [0] * len(b)
        for j in range(1, len(b) + 1):
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] != b[j - 1]))
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                cur[j] = min(cur[j], prev2[j - 2] + 1)
        if min(cur) > k:
            return False
        prev2, prev = prev, cur
    return prev[-1] <= k


def fuzzy(tok, vocab, limit=6):
    """Vocabulary terms within edit distance 1 (len < 8) or 2 (len >= 8), most frequent first.

    ponytail: linear scan over the FTS vocabulary; fine up to ~10^5 terms (a few 100 ms).
    Upgrade path: a trigram FTS table over the vocabulary.
    """
    k = 1 if len(tok) < 8 else 2
    out = []
    for term, df in vocab.items():
        if abs(len(term) - len(tok)) > k or term.isdigit() or term[0] != tok[0] and term[1:2] != tok[1:2]:
            continue
        if edit_distance_le(tok, term, k):
            out.append((df, term))
    return [t for _, t in sorted(out, reverse=True)[:limit]]


def _known(t, s, vocab):
    return t in vocab or (len(s) >= 4 and any(term.startswith(s) for term in vocab))


def query_groups(text, vocab):
    """[(token, [(term, is_prefix)])] for the keyword part of the query. One group per query word."""
    text = re.sub(r"(\w)['’](\w+)", r"\1", text)  # Ahmet'ten -> Ahmet, Sarah's -> Sarah
    groups, seen = [], set()
    for t in tokens(fold(text)):
        if t in STOP or t in seen or len(t) < 2:
            continue
        seen.add(t)
        variants = {(t, False)}
        if not t.isdigit():
            s = stem(t)
            if len(s) >= 4:
                variants.add((s, True))
            for syn in SYNONYMS.get(t, []):
                variants.add((syn, False))
            if len(t) >= 4 and not _known(t, s, vocab):
                for f in fuzzy(t, vocab):
                    variants.add((f, False))
        groups.append((t, sorted(variants)))
    return groups


def match_expr(groups):
    q = lambda t, p: '"%s"%s' % (t.replace('"', '""'), "*" if p else "")
    return " OR ".join("(" + " OR ".join(q(t, p) for t, p in vs) + ")" for _, vs in groups)


def group_regex(vs):
    return re.compile(r"\b(?:" + "|".join(re.escape(t) + ("" if p else r"\b") for t, p in vs) + ")")


def snippet(text, regexes, span=None, width=220):
    """Short window of the original text with the most query-term hits (else the best semantic chunk)."""
    if not text:
        return ""
    f = fold(text)
    hits = sorted(m.start() for r in regexes for m in r.finditer(f))
    if hits:
        counts = [sum(1 for x in hits[i:] if x < h + width * 0.8) for i, h in enumerate(hits)]
        start = max(0, hits[counts.index(max(counts))] - 50)
    else:
        start = span[0] if span else 0
    if start > 0:
        sp = text.rfind(" ", 0, start)
        start = sp + 1 if sp >= 0 and start - sp < 30 else start
    end = min(len(text), start + width)
    if end < len(text):
        sp = text.rfind(" ", start, end)
        end = sp if sp > start + width // 2 else end
    return ("… " if start > 0 else "") + " ".join(text[start:end].split()) + (" …" if end < len(text) else "")


def marks(text, regexes):
    """[[start, length]] of the query words in text, in UTF-16 units (the add-in highlights them in the preview).
    A prefix match marks the whole word."""
    f = fold(text or "")  # same length as text
    spans = []
    for rx in regexes:
        for m in rx.finditer(f):
            b = m.end()
            while b < len(f) and f[b].isalnum():
                b += 1
            spans.append((m.start(), b))
    merged = []
    for a, b in sorted(spans):
        if merged and a <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    u16 = lambda i: len(text[:i].encode("utf-16-le")) // 2
    return [[u16(a), u16(b) - u16(a)] for a, b in merged]


def confidence(cos, zv, cov, semantic, n_groups):
    """'yüksek' / 'orta' / 'zayıf' for the semantic stage; 'kelime' or None (hidden) for the keyword-only stage."""
    if not semantic:  # keyword-only stage: every query word must match
        return "kelime" if n_groups and cov >= COV_FULL else None
    cos, zv = (-1.0 if cos is None else cos), (-9.0 if zv is None else zv)
    if cos >= COS_HIGH or (cos >= COS_HIGH_Z and zv >= Z_HIGH) or (cov >= COV_HIGH and cos >= COS_MEDIUM):
        return "yüksek"
    if cos >= COS_MEDIUM or cov >= COV_MEDIUM:
        return "orta"
    return "zayıf"


def run(store, query, now, folder_ids, encoder=None, model_rev=None, dim=512, top_k=10):
    """Search the folder scope. Without an encoder the result is keyword-only."""
    windows, spans, keep = hm_dates.parse(query, now)
    text_q = (hm_dates.remove_spans(query, spans) + " " + " ".join(keep)).strip()
    groups = query_groups(text_q, store.vocab())
    regexes = [group_regex(vs) for _, vs in groups]
    scope = store.scope_rowids(folder_ids)

    kw_rank, att_rank, att_hits = {}, {}, {}
    if groups:
        match = match_expr(groups)
        cand = [r for r, _ in store.fts(match) if r in scope]  # sorted by BM25
        kw_rank = {r: i for i, r in enumerate(cand, 1)}
        cand = [r for r, _ in store.fts_att(match) if r in scope]
        att_rank = {r: i for i, r in enumerate(cand, 1)}
        # Per query word: the mails whose attachments hold it (for coverage; one FTS query per word).
        att_hits = {g: store.fts_att_rowids(match_expr([(g, vs)])) for g, vs in groups}

    sem_rank, cos, best_seq, z, sem_stats = {}, {}, {}, {}, None
    if encoder is not None:
        qv = encoder.encode([encoder_query(query)], max_tokens=128)[0]  # the full query, date words included
        ids, seq, mat = store.all_vectors(model_rev, dim, folder_ids)
        if len(ids):
            sims = mat @ qv
            order = np.lexsort((-sims, ids))
            first = np.ones(len(order), bool)
            first[1:] = ids[order][1:] != ids[order][:-1]
            best = order[first]
            item_ids, item_cos, item_seq = ids[best], sims[best], seq[best]
            med = float(np.median(item_cos))
            sigma = float(np.median(np.abs(item_cos - med))) * 1.4826 or 1e-6
            sem_stats = {"median": round(med, 4), "sigma": round(sigma, 4), "n": int(len(item_ids))}
            for rank, i in enumerate(np.argsort(-item_cos)[:400].tolist(), 1):
                r = int(item_ids[i])
                sem_rank[r], cos[r], best_seq[r] = rank, float(item_cos[i]), int(item_seq[i])
                z[r] = (cos[r] - med) / sigma

    info = store.items(set(kw_rank) | set(att_rank) | set(sem_rank))
    coverage, scored = {}, []
    for r, it in info.items():
        hay = fold(" ".join(it[k] or "" for k in ("subject", "sender", "sender_addr", "recipients", "attachments",
                                                   "main_text")))
        found = sum(1 for (g, _), rx in zip(groups, regexes) if rx.search(hay) or r in att_hits.get(g, ()))
        coverage[r] = found / len(regexes) if regexes else 0.0
        dt = _dt(it["sent_at"])
        ds = hm_dates.date_score(dt, windows) if windows and dt else 0.0
        # An attachment hit counts less than a hit in the mail itself (W_ATT): attachments are long.
        kw = (1.0 / (10 + kw_rank[r]) if r in kw_rank else 0.0) + (W_ATT / (10 + att_rank[r]) if r in att_rank else 0.0)
        if encoder is not None:
            s = cos.get(r, -1.0) + W_COV * coverage[r] * min(1.0, len(groups) / 2) + W_DATE * ds + W_KW * kw
        else:  # keyword-only stage: coverage first, then BM25, date as tie-breaker
            s = coverage[r] + 0.1 * kw + 0.05 * ds
        conf = confidence(cos.get(r), z.get(r), coverage[r], encoder is not None, len(groups))
        if conf:
            scored.append((s, r, ds, conf))
    scored.sort(key=lambda x: (-x[0], x[1]))

    out = []
    for s, r, ds, conf in scored[:top_k]:
        it = info[r]
        chunk = store.chunk(r, best_seq[r]) if r in best_seq else None  # (src, start, end)
        snip, att_name = _snippet(store, r, it, regexes, chunk)
        why = []
        if r in coverage:
            why.append("kelime %d/%d" % (round(coverage[r] * len(groups)), len(groups)))
        if r in cos:
            why.append("anlam %.2f" % cos[r])
        if att_name:
            why.append("ekte")
        if windows:
            why.append("tarih uygun" if ds >= 0.999 else "tarih yakın" if ds > 0.3 else "tarih dışı")
        out.append({
            "eid": it["entry_id"], "folder": it["folder_id"], "subject": it["subject"] or "(konu yok)",
            "sender": it["sender"] or "", "recipients": it["recipients"] or "", "date": it["sent_at"],
            "snippet": snip, "match_attachment": att_name, "hl": marks(snip, regexes),
            "hl_subject": marks(it["subject"] or "", regexes),
            "attachments": it["attachments"] or "", "score": round(s, 5), "confidence": conf, "why": why,
            "cos": round(cos[r], 4) if r in cos else None, "z": round(z[r], 2) if r in z else None,
            "coverage": round(coverage.get(r, 0.0), 2),
        })
    return {"results": out, "date_windows": hm_dates.describe(windows), "terms": [g for g, _ in groups],
            "semantic": encoder is not None, "sem_stats": sem_stats}


def _snippet(store, r, it, regexes, chunk):
    """(snippet, attachment name or None) from the text with the most query words: the mail or one of its
    attachments (a tie goes to the mail). Without a word hit: the best semantic chunk (body or attachment)."""
    main = it["main_text"] or ""
    texts = store.att_texts(r) if regexes or (chunk and chunk[0]) else []  # [(idx, name, text)]
    def hits(t):
        f = fold(t)
        return sum(1 for rx in regexes if rx.search(f))

    best, best_hits = (main, None), hits(main) if main else 0
    for idx, name, text in texts:
        n = hits(text)
        if n > best_hits:
            best, best_hits = (text, name), n
    if best_hits:
        return snippet(best[0], regexes), best[1]
    if chunk and chunk[0]:
        for idx, name, text in texts:
            if idx == chunk[0]:
                return snippet(text, [], (chunk[1], chunk[2])), name
    return snippet(main or it["attachments"] or "", regexes, chunk[1:] if chunk else None), None


def encoder_query(text):
    import hm_embed
    return hm_embed.query_text(text)


def _dt(s):
    try:
        return datetime.fromisoformat(s) if s else None
    except ValueError:
        return None
