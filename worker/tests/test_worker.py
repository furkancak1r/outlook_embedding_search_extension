"""Unit and protocol tests for the HatırMail worker. Run: python -m unittest discover -s worker/tests -v

The protocol tests start the real worker process with --fake-encoder (no model, no torch import).
"""
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import unittest.mock
from datetime import datetime
from pathlib import Path

WORKER = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WORKER))

import hm_attach  # noqa: E402
import hm_dates  # noqa: E402
import hm_index  # noqa: E402
import hm_search  # noqa: E402
from hm_store import Store  # noqa: E402
from hm_text import chunk_spans, fold, split_body  # noqa: E402

NOW = datetime(2026, 10, 14, 12, 0)
REV = "test-rev"
ATT = WORKER.parent / "eval" / "att"  # fixture files from eval/make_att_fixtures.py


def mail(eid, subject, body, date="2026-09-10T10:00:00", sender="Ayşe Kaya", folder="F1", mod=None,
         atts=(), cls="IPM.Note", to="Furkan Çakır"):
    return {"eid": eid, "folder": folder, "cls": cls, "mod": mod or date, "subject": subject, "sender": sender,
            "sender_addr": sender.split()[0].lower() + "@x.example", "to": to, "cc": "", "date": date,
            "attachments": list(atts), "body": body}


class TextTests(unittest.TestCase):
    def test_fold_turkish(self):
        self.assertEqual(fold("İSTANBUL IŞIK ığdır Çağrı ŞÜKRÜ"), "istanbul isik igdir cagri sukru")
        self.assertEqual(len(fold("İŞÇİ")), 4)

    def test_split_reply_signature_disclaimer(self):
        body = ("Merhaba,\n\nİade onaylandı, 12.500 TL.\n\nSaygılarımla,\nMehmet Yılmaz\nMuhasebe\n\n"
                "Bu e-posta gizli bilgi içerebilir. Muhatabı değilseniz lütfen göndereni bilgilendirin ve mesajı "
                "imha edin. Yasal uyarı: şirketimiz sorumlu tutulamaz.\n\n"
                "Kimden: Ayşe <a@x.example>\nGönderildi: 3 Mart 2025\nKime: Mehmet\nKonu: İade\n\nİade ne oldu?")
        main, quoted = split_body(body, "YNT: İade")
        self.assertEqual(main, "Merhaba,\n\nİade onaylandı, 12.500 TL.")
        self.assertIn("İade ne oldu?", quoted)

    def test_forward_keeps_forwarded_text(self):
        body = "Bilgin olsun.\n\n-----Original Message-----\nFrom: X\nSent: Monday\nTo: Y\n\nForklift tamiri 4.800 TL."
        main, _ = split_body(body, "FW: Forklift")
        self.assertIn("Forklift tamiri", main)

    def test_english_on_wrote_and_mobile(self):
        body = "Sounds good.\n\nSent from my iPhone\n\nOn Mon, 3 Mar 2025 at 10:12, Sarah <s@x.example> wrote:\n> old"
        main, quoted = split_body(body, "Re: plan")
        self.assertEqual(main, "Sounds good.")
        self.assertIn("old", quoted)

    def test_chunk_spans_cover_text(self):
        text = ("Kelime " * 900).strip()
        spans = chunk_spans(text)
        self.assertEqual(spans[0][0], 0)
        self.assertEqual(spans[-1][1], len(text))
        self.assertTrue(all(b > a for a, b in spans))
        self.assertTrue(all(spans[i + 1][0] < spans[i][1] for i in range(len(spans) - 1)))  # overlap


class DateTests(unittest.TestCase):
    def window(self, q):
        w, _, _ = hm_dates.parse(q, NOW)
        return [(s.date().isoformat(), (e - hm_dates.D).date().isoformat()) for s, e, _ in w]

    def test_relative(self):
        self.assertEqual(self.window("geçen ay fatura"), [("2026-09-01", "2026-09-30")])
        self.assertEqual(self.window("last month invoice"), [("2026-09-01", "2026-09-30")])
        self.assertEqual(self.window("geçen hafta"), [("2026-10-05", "2026-10-11")])
        self.assertEqual(self.window("dün"), [("2026-10-13", "2026-10-13")])

    def test_months_and_parts(self):
        self.assertEqual(self.window("mart başında teklif"), [("2026-03-01", "2026-03-10")])
        self.assertEqual(self.window("geçen yıl kasımda sözleşme"), [("2025-11-01", "2025-11-30")])
        self.assertEqual(self.window("15 mart 2025"), [("2025-03-15", "2025-03-15")])
        self.assertEqual(self.window("early march 2025"), [("2025-03-01", "2025-03-10")])
        self.assertEqual(self.window("yılbaşından önce izin"), [("2025-12-01", "2025-12-31")])

    def test_ambiguous_words_are_not_dates(self):
        for q in ("Kasım'ın gönderdiği teklif", "aralığı genişletin", "nişan fotoğrafları", "Martin's report",
                  "may I see the budget"):
            self.assertEqual(self.window(q), [], q)

    def test_date_phrase_removed_from_text_query(self):
        w, spans, keep = hm_dates.parse("geçen ay Ahmet'ten gelen kargo faturası", NOW)
        self.assertEqual(hm_dates.remove_spans("geçen ay Ahmet'ten gelen kargo faturası", spans),
                         "Ahmet'ten gelen kargo faturası")

    def test_date_score_soft(self):
        w, _, _ = hm_dates.parse("geçen ay", NOW)
        self.assertEqual(hm_dates.date_score(datetime(2026, 9, 15), w), 1.0)
        near, far = hm_dates.date_score(datetime(2026, 10, 3), w), hm_dates.date_score(datetime(2026, 3, 1), w)
        self.assertTrue(0.3 < near < 1.0 and far < 0.01, (near, far))


class StoreSearchTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.store = Store(os.path.join(self.tmp.name, "t.sqlite"), REV)
        import hm_embed
        self.enc = hm_embed.FakeEncoder()
        mails = [
            mail("E1", "Kargo faturası", "Eylül kargo faturası ekte. Tutar 4.250 TL.", "2026-09-12T09:00:00",
                 "Ahmet Demir", atts=["Fatura_2026-09.pdf"]),
            mail("E2", "Kargo faturası", "Mart kargo faturası ekte. Tutar 3.100 TL.", "2026-03-12T09:00:00",
                 "Ahmet Demir"),
            mail("E3", "Sunucu bakımı", "Cumartesi gecesi planlı bakım var.", "2026-09-20T09:00:00", "Sarah Miller"),
            mail("E4", "Toplantı notları", "Bütçe ve takvim konuşuldu. Sözleşme taslağı gelecek hafta.",
                 "2026-10-07T15:00:00", "Mehmet Yılmaz", atts=["Bütçe_2026_Q4.xlsx"]),
            mail("E5", "NDR", "Teslim edilemedi", cls="REPORT.IPM.Note.NDR"),
        ]
        with self.store.tx():
            for m in mails:
                hm_index.ingest(self.store, m, 1, REV)
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)

    def tearDown(self):
        self.store.close()
        self.tmp.cleanup()

    def search(self, q, **kw):
        return hm_search.run(self.store, q, NOW, ["F1"], encoder=self.enc, model_rev=REV, **kw)["results"]

    def test_non_mail_classes_skipped(self):
        self.assertIsNone(self.store.db.execute("SELECT 1 FROM items WHERE entry_id='E5'").fetchone())

    def test_date_boost_picks_right_month(self):
        self.assertEqual(self.search("geçen ay Ahmet'ten gelen kargo faturası")[0]["eid"], "E1")
        self.assertEqual(self.search("mart ayındaki kargo faturası")[0]["eid"], "E2")

    def test_ascii_and_typo(self):
        self.assertEqual(self.search("toplanti notlari butce")[0]["eid"], "E4")
        self.assertEqual(self.search("sozlesmee taslagi")[0]["eid"], "E4")
        self.assertEqual(self.search("sunucu bakimi")[0]["eid"], "E3")

    def test_attachment_synonym(self):
        self.assertEqual(self.search("bütçe excel dosyası")[0]["eid"], "E4")

    def test_keyword_stage_needs_full_match(self):
        res = hm_search.run(self.store, "uzay mekiği", NOW, ["F1"])["results"]
        self.assertEqual(res, [])

    def test_snippet_from_original_text(self):
        r = self.search("cumartesi bakım")[0]
        self.assertIn("Cumartesi gecesi", r["snippet"])

    def test_moved_item_reuses_vectors(self):
        n_vec = self.store.db.execute("SELECT count(*) FROM vectors").fetchone()[0]
        with self.store.tx():
            self.assertEqual(hm_index.ingest(self.store, mail("E3", "Sunucu bakımı", "Cumartesi gecesi planlı bakım var.",
                                                               "2026-09-20T09:00:00", "Sarah Miller", folder="F2"), 2, REV),
                             "same")
        st = hm_index.embed_pending(self.store, lambda: self.fail("model must not load"), ["F1", "F2"], REV)
        self.assertEqual(st["embedded_chunks"], 0)
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM vectors").fetchone()[0], n_vec)

    def test_changed_item_reindexed_and_old_text_gone(self):
        with self.store.tx():
            self.assertEqual(hm_index.ingest(self.store, mail("E3", "Sunucu bakımı", "Pazar sabahı yazılım güncellemesi.",
                                                               "2026-09-20T09:00:00", "Sarah Miller"), 2, REV), "changed")
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        self.assertFalse(any(r["eid"] == "E3" and "Cumartesi" in r["snippet"] for r in self.search("cumartesi gecesi")))
        self.assertEqual(self.search("yazılım güncellemesi")[0]["eid"], "E3")

    def test_cache_follows_every_change(self):
        """The vector matrix and the filter values of a kept worker: the same object while the index does not change, and
        the same values as a new query after each kind of change (new mail, moved mail, deleted mail, new vector, new
        attachment chunk, deleted vector). The arrays are read-only."""
        import numpy as np
        fs = ["F1", "F2"]
        fresh = lambda: (self.store.all_vectors(REV, 512, fs, cache=False), hm_search._facets(self.store, 1000))
        same = lambda: self.assertTrue(all(np.array_equal(a, b) for a, b in zip(self.store.all_vectors(REV, 512, fs), fresh()[0]))
                                       and hm_search.facets(self.store) == fresh()[1])
        first = self.store.all_vectors(REV, 512, fs)
        self.assertIs(self.store.all_vectors(REV, 512, fs)[2], first[2])  # no change: the same arrays
        self.assertIs(hm_search.facets(self.store), hm_search.facets(self.store))
        self.assertFalse(first[2].flags.writeable)
        same()
        steps = [
            lambda: hm_index.ingest(self.store, mail("E9", "Yeni", "Yeni posta metni.", sender="Zeynep Arslan"), 2, REV),
            lambda: hm_index.ingest(self.store, mail("E3", "Sunucu bakımı", "Cumartesi gecesi planlı bakım var.",
                                                     "2026-09-20T09:00:00", "Sarah Miller", folder="F2"), 2, REV),
            lambda: self.store.delete_item(self.store.item_state(["E2"])["E2"][0]),
            lambda: hm_index.embed_pending(self.store, lambda: self.enc, fs, REV),
            lambda: self.store.set_att_index(self.store.item_state(["E1"])["E1"][0], "ek metni",
                                             [(1, 0, 8, "hash-att-1")]),
            lambda: self.store.db.execute("DELETE FROM vectors WHERE text_hash IN (SELECT text_hash FROM vectors LIMIT 1)"),
        ]
        for i, step in enumerate(steps):
            gen = self.store.meta("data_gen")
            if i == 3:
                step()  # embed_pending owns its transactions
            else:
                with self.store.tx():
                    step()
            self.assertGreater(self.store.meta("data_gen"), gen)
            same()
        with self.store.tx():  # only a flag: the index state for the caches stays
            gen = self.store.meta("data_gen")
            self.store.mark_stale([self.store.item_state(["E1"])["E1"][0]])
        self.assertEqual(self.store.meta("data_gen"), gen)

    def test_vocab_cache(self):
        """The FTS vocabulary of a kept worker: one object for the sessions while the index does not change, the same
        terms as a new query after a new and a deleted mail, read only."""
        path = self.store.path
        fresh = lambda s: dict(s.db.execute("SELECT term, doc FROM fts_vocab"))
        a, b = Store(path, REV), Store(path, REV)  # two sessions of one worker
        self.assertIs(a.vocab(), b.vocab())
        self.assertEqual(dict(a.vocab()), fresh(a))
        with self.assertRaises(TypeError):
            a.vocab()["neu"] = 1
        for step in (lambda s: hm_index.ingest(s, mail("E9", "Zeplin", "Zeplin sevkiyatı yarın."), 2, REV),
                     lambda s: s.delete_item(s.item_state(["E9"])["E9"][0])):
            with a.tx():
                step(a)
            c = Store(path, REV)
            for s in (a, c):  # the writer and a new session see the new terms
                self.assertEqual(dict(s.vocab()), fresh(s))
                self.assertEqual("zeplin" in s.vocab(), "zeplin" in fresh(s))
            c.close()
        a.close(), b.close()

    def test_typo_scan_same_as_full_scan(self):
        """The prefix check and the typo scan on ranges of the sorted terms give what a check of every term gives: words
        of the vocabulary with 1-2 edits, random strings, digits, Turkish and other letters, 2-3 char words, an empty
        term and terms with a digit or a non-ASCII first char in the vocabulary. (query_groups gives 4+ chars; with 1
        char the old scan failed on the empty term.)"""
        import random

        def full_fuzzy(tok, vocab, limit=6):  # the scan before 10-26
            k = 1 if len(tok) < 8 else 2
            out = [(df, term) for term, df in vocab.items()
                   if not (abs(len(term) - len(tok)) > k or term.isdigit() or term[0] != tok[0] and term[1:2] != tok[1:2])
                   and hm_search.edit_distance_le(tok, term, k)]
            return [t for _, t in sorted(out, reverse=True)[:limit]]

        rnd = random.Random(11)
        alpha = "abcçdefgğhıijklmnoöprsştuüvyz0123456789-é"
        vocab = dict(self.store.vocab())
        for _ in range(3000):
            vocab["".join(rnd.choice(alpha) for _ in range(rnd.randint(1, 12)))] = rnd.randint(1, 50)
        vocab.update({"": 1, "9kargo": 2, "çarşamba": 3, "ömür": 4, "ıslak": 5})
        terms = sorted(vocab)
        words = list(vocab)
        sample = [w[:rnd.randint(2, 3)] for w in rnd.sample(words, 80) if len(w) >= 2]
        for _ in range(600):
            w = list(rnd.choice(words) or "a")
            for _ in range(rnd.randint(1, 2)):
                p = rnd.randrange(len(w) + 1)
                if rnd.random() < 0.4 or p == len(w):
                    w.insert(p, rnd.choice(alpha))
                elif rnd.random() < 0.5:
                    w[p] = rnd.choice(alpha)
                elif len(w) > 1:
                    del w[p]
            sample.append("".join(w))
        sample += ["".join(rnd.choice(alpha) for _ in range(rnd.randint(2, 12))) for _ in range(200)]
        for w in (w for w in sample if len(w) >= 2):
            self.assertEqual(hm_search.fuzzy(w, vocab, terms), full_fuzzy(w, vocab), w)
            s = hm_search.stem(w)
            self.assertEqual(hm_search._prefix_known(s, terms), len(s) >= 4 and any(t.startswith(s) for t in vocab), w)
        # The whole query: the same groups as the code before 10-26 (prefix check and typo scan over every term). The
        # scan must go to the end of the sorted list ("z..." and "ü..." are the last terms; fold() makes ü in a query u).
        vocab.update({"zeytin": 6, "zeytinlik": 2, "uretim": 7, "uretici": 3, "üzüm": 2})
        old_prefix = lambda s, terms: len(s) >= 4 and any(term.startswith(s) for term in vocab)
        with unittest.mock.patch.object(hm_search, "_prefix_known", old_prefix), \
                unittest.mock.patch.object(hm_search, "fuzzy", lambda tok, v, terms: full_fuzzy(tok, v)):
            old = [hm_search.query_groups(q, vocab) for q in sample + ["zeytinn üretimm", "kargo faturasi zeytn"]]
        new = [hm_search.query_groups(q, vocab) for q in sample + ["zeytinn üretimm", "kargo faturasi zeytn"]]
        self.assertEqual(new, old)
        self.assertIn(("zeytin", False), new[-2][0][1])
        self.assertIn(("uretim", False), new[-2][1][1])
        self.assertIn(("zeytin", False), new[-1][2][1])

    def test_covering_index_and_stable_orders(self):
        """items_cov: an old file gets it at the next open; coverage() and pending() give the same values and order as
        without it. best_per_item: with equal cosines of one mail, the lowest seq wins in every row order."""
        import numpy as np
        fs = ["F1"]
        with self.store.tx():
            for i in range(6):  # the same date: pending() orders them by rowid
                hm_index.ingest(self.store, mail("P%d" % i, "Paket %d" % i, "Paket teslim.", "2026-09-01T09:00:00"), 2, REV)
        before = (self.store.coverage(fs), self.store.pending(fs))
        self.assertEqual([r[1] for r in before[1]], ["Paket %d" % i for i in range(6)])
        import sqlite3

        def plan():  # a new connection: a cached EXPLAIN statement keeps its old plan after a schema change
            c = sqlite3.connect(self.store.path)
            try:
                return " ".join(r[3] for r in c.execute(
                    "EXPLAIN QUERY PLAN SELECT folder_id, count(*), sum(last_mod IS NOT NULL), sum(embedded) FROM items"
                    " WHERE folder_id IN (?) GROUP BY folder_id", fs))
            finally:
                c.close()
        self.assertIn("COVERING INDEX items_cov", plan())
        self.store.db.execute("DROP INDEX items_cov")  # a file from before 10-26
        self.assertNotIn("items_cov", plan())
        self.assertEqual((self.store.coverage(fs), self.store.pending(fs)), before)
        self.store.close()
        self.store = Store(self.store.path, REV)  # the next session creates the index (no statistics yet)
        has = "SELECT count(*) FROM sqlite_master WHERE name='items_cov'"
        self.assertEqual(self.store.db.execute(has).fetchone()[0], 1)
        self.assertEqual((self.store.coverage(fs), self.store.pending(fs)), before)
        self.store.close()  # PRAGMA optimize analyzes the new index: from the next session on, the planner uses it
        self.store = Store(self.store.path, REV)
        self.assertIn("COVERING INDEX items_cov", plan())
        self.assertEqual((self.store.coverage(fs), self.store.pending(fs)), before)
        ids, seq, sims = np.array([7, 7, 7, 3]), np.array([0, 1, 2, 0], np.int32), np.array([0.5, 0.9, 0.9, 0.1], np.float32)
        for p in ([0, 1, 2, 3], [3, 2, 1, 0], [2, 0, 3, 1]):
            got = hm_search.best_per_item(ids[p], seq[p], sims[p])
            self.assertEqual((got[0].tolist(), got[2].tolist()), ([3, 7], [0, 1]), p)

    def test_session_memo_and_lazy_coverage(self):
        """run() with one memo for the keyword and the final stage gives the same results as run() without it, also
        after the index changes between two runs (a new mail with the query words). Ties: mails with equal scores keep
        the rowid order, and the 10th place is the same with a top_k of 1, 2 and 3."""
        memo = {}
        for q in ("kargo faturası", "toplanti notlari butce", "geçen ay Ahmet'ten gelen kargo faturası", "uzay mekiği"):
            plain = (hm_search.run(self.store, q, NOW, ["F1"]), self.search_full(q))
            with_memo = (hm_search.run(self.store, q, NOW, ["F1"], memo=memo), self.search_full(q, memo=memo))
            self.assertEqual(plain, with_memo, q)
        with self.store.tx():
            hm_index.ingest(self.store, mail("E7", "Kargo faturası", "Kasım kargo faturası.", "2026-09-28T09:00:00"), 2, REV)
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        self.assertEqual(self.search_full("kargo faturası", memo=memo), self.search_full("kargo faturası"))
        self.assertIn("E7", [r["eid"] for r in self.search_full("kargo faturası")["results"]])
        with self.store.tx():  # equal mails: the same score, the order is the rowid
            for i in range(6):
                hm_index.ingest(self.store, mail("T%d" % i, "Paket", "Paket teslim edildi.", "2026-09-01T09:00:00"), 3, REV)
        full = [r["eid"] for r in hm_search.run(self.store, "paket teslim", NOW, ["F1"], top_k=10)["results"]]
        self.assertEqual(full[:6], ["T%d" % i for i in range(6)])
        for k in (1, 2, 3):
            self.assertEqual([r["eid"] for r in hm_search.run(self.store, "paket teslim", NOW, ["F1"], top_k=k)["results"]],
                             full[:k])

    def search_full(self, q, **kw):
        return hm_search.run(self.store, q, NOW, ["F1"], encoder=self.enc, model_rev=REV, **kw)

    def test_memo_follows_changed_text(self):
        """A memo entry belongs to a text, not to a rowid: SQLite gives the next new row the rowid of a deleted last
        row, so a changed mail can come back with its old rowid and a new text."""
        last = max(self.store.scope_rowids(["F1"]))
        eid = self.store.items([last])[last]["entry_id"]
        memo = {}
        before = self.search_full("zeplin kargosu", memo=memo)
        self.assertNotIn(eid, [r["eid"] for r in before["results"] if r["coverage"] == 1.0])
        with self.store.tx():
            hm_index.ingest(self.store, mail(eid, "Zeplin kargosu", "Zeplin kargosu yarın gelir."), 2, REV)
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        self.assertEqual(self.store.item_state([eid])[eid][0], last)  # the same rowid, a new text
        self.assertEqual(self.search_full("zeplin kargosu", memo=memo), self.search_full("zeplin kargosu"))

    def test_lazy_coverage_stop_rule(self):
        """The read stops only when the next upper limit is below the top_k-th exact score of the mails read so far.
        Fixed cosines (query vector e1): C0 0.75 and 48 others 0.74 without the query words, C1 0.70 and C50 0.66 with
        both words. Upper limits: C0 0.83 > C1 0.825 > the 48 at 0.82 (the first read batch of 50) > C50 0.782.
        Exact scores: C1 0.825 > C50 0.782 > C0 0.75 > the 48 at 0.74. So the 2nd result, C50, is after the first batch,
        and the 2nd mail of that batch by upper limit (C1) has a higher score than the 2nd best of the batch (C0)."""
        import numpy as np
        with self.store.tx():
            hm_index.ingest(self.store, mail("C0", "Teslimat", "Teslim edildi.", sender="Ece Su"), 2, REV)
            hm_index.ingest(self.store, mail("C1", "Zeplin palet", "Zeplin palet.", sender="Ece Su"), 2, REV)
            for i in range(48):
                hm_index.ingest(self.store, mail("F%d" % i, "Teslimat %d" % i, "Teslim edildi.", sender="Ece Su"), 2, REV)
            hm_index.ingest(self.store, mail("C50", "Zeplin ve palet", "Yarın zeplin ile palet gelir, depo hazır olsun.",
                                             sender="Ece Su"), 2, REV)
        cos = {"C0": 0.75, "C1": 0.70, "C50": 0.66}
        vec = lambda c: np.array([c, (1 - c * c) ** 0.5] + [0.0] * 510, np.float32)
        rows = self.store.db.execute("SELECT i.entry_id, c.text_hash FROM items i JOIN chunks c ON c.item_rowid=i.rowid"
                                     " WHERE i.entry_id LIKE 'C%' OR i.entry_id LIKE 'F%'").fetchall()
        with self.store.tx():
            self.store.put_vectors([(h, vec(cos.get(e, 0.74))) for e, h in rows], REV)

        class E1:
            def encode(self, texts, max_tokens=128):
                return np.array([vec(1.0)])

        res = lambda k: [r["eid"] for r in hm_search.run(self.store, "zeplin palet", NOW, ["F1"], encoder=E1(),
                                                         model_rev=REV, top_k=k)["results"]]
        self.assertEqual(res(10**6)[:3], ["C1", "C50", "C0"])
        for k in (1, 2, 3, 10):
            self.assertEqual(res(k), res(10**6)[:k], k)

    def test_lazy_coverage_many_candidates(self):
        """More candidates than one read batch (50): the first top_k results are the first top_k of a full evaluation
        (top_k larger than all candidates), in the keyword and the final stage, with and without a date and a filter."""
        import random
        rnd = random.Random(4)
        words = "zeplin kargo fatura teslim depo palet bakım rapor sevk stok".split()
        with self.store.tx():
            for i in range(160):
                body = " ".join(rnd.choice(words) for _ in range(rnd.randint(1, 30)))
                subj = " ".join(rnd.sample(words, rnd.randint(1, 3)))
                d = "2026-%02d-%02dT09:00:00" % (rnd.randint(1, 10), rnd.randint(1, 28))
                hm_index.ingest(self.store, mail("R%d" % i, subj, body, d, sender=rnd.choice(["Ali Can", "Ece Su"])), 2, REV)
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        allow = hm_search.allowed(self.store, ["F1"], {"sender": hm_search.sender_key("Ali Can", "ali@x.example")})
        for q in ("zeplin kargo", "geçen ay fatura teslim", "depo palet stok sevk", "bakım"):
            for enc in (None, self.enc):
                for a in (None, allow):
                    full = hm_search.run(self.store, q, NOW, ["F1"], encoder=enc, model_rev=REV, top_k=10**6, allow=a)
                    for k in (1, 3, 10):
                        got = hm_search.run(self.store, q, NOW, ["F1"], encoder=enc, model_rev=REV, top_k=k, allow=a)
                        self.assertEqual(got["results"], full["results"][:k], (q, enc is not None, a is not None, k))

    def test_sweep_removes_deleted(self):
        with self.store.tx():
            state = self.store.item_state(["E1", "E2", "E3"])
            removed = self.store.sweep(["F1"], {s[0] for s in state.values()})
        self.assertEqual(removed, 1)  # E4 is not in the new manifest
        self.assertTrue(all(r["eid"] != "E4" for r in self.search("toplantı notları")))

    def test_resume_after_stop(self):
        with self.store.tx():
            for i in range(10):
                hm_index.ingest(self.store, mail(f"N{i}", f"Konu {i}", f"Benzersiz metin numara {i} " * 40), 3, REV)
        calls = {"n": 0}

        def stop():
            calls["n"] += 1
            return calls["n"] > 2

        st = hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV, window=2, stop=stop)
        self.assertTrue(st["stopped"])
        left = self.store.db.execute("SELECT count(*) FROM items WHERE embedded=0").fetchone()[0]
        self.assertTrue(0 < left < 10)
        st = hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        self.assertFalse(st["stopped"])
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM items WHERE embedded=0").fetchone()[0], 0)

    def test_filters_narrow_results(self):
        """Ara filters: only mails of the chosen sender / recipient; Turkish upper and lower case are one key; the
        confidence of a mail is the same with and without a filter (z statistics on the whole scope)."""
        with self.store.tx():
            hm_index.ingest(self.store, mail("E6", "Kargo faturası", "Ekim kargo faturası ekte.", "2026-10-02T09:00:00",
                                             "Zeynep Arslan", to="ŞÜKRÜ IŞIK; 'Furkan Çakır'"), 1, REV)
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        ahmet = hm_search.sender_key("Ahmet Demir", "ahmet@x.example")
        run = lambda f: hm_search.run(self.store, "kargo faturası", NOW, ["F1"], encoder=self.enc, model_rev=REV,
                                      allow=hm_search.allowed(self.store, ["F1"], f))["results"]
        full = {r["eid"]: r for r in run(None)}
        self.assertTrue({"E1", "E2", "E6"} <= set(full))
        only = run({"sender": ahmet})
        self.assertEqual({r["eid"] for r in only}, {"E1", "E2"})
        self.assertEqual([(r["confidence"], r["cos"]) for r in only], [(full[r["eid"]]["confidence"], full[r["eid"]]["cos"]) for r in only])
        self.assertEqual([r["eid"] for r in run({"recipient": hm_search.sender_key("Şükrü Işık", "")})], ["E6"])
        self.assertEqual({r["eid"] for r in run({"recipient": "furkan cakir"})} & {"E1", "E2", "E6"}, {"E1", "E2", "E6"})
        self.assertEqual(run({"sender": ahmet, "recipient": "sukru isik"}), [])
        self.assertEqual(hm_search.allowed(self.store, ["F1"], {"sender": None, "recipient": ""}), None)
        kw = hm_search.run(self.store, "sunucu bakımı", NOW, ["F1"], allow=hm_search.allowed(self.store, ["F1"], {"sender": ahmet}))
        self.assertEqual(kw["results"], [])  # keyword stage too

    def test_date_and_attachment_filters(self):
        """Ara filters Tarih (date_from inclusive, date_to exclusive) and Ek (yes / no / file type), alone and with the
        sender filter; the search keeps only these mails."""
        eids = lambda f: {it["entry_id"] for it in self.store.items(hm_search.allowed(self.store, ["F1"], f)).values()}
        sep = {"date_from": "2026-09-01T00:00:00", "date_to": "2026-10-01T00:00:00"}
        self.assertEqual(eids(sep), {"E1", "E3"})
        self.assertEqual(eids({"date_from": "2026-09-20T09:00:00"}), {"E3", "E4"})  # a mail at date_from is inside
        self.assertEqual(eids({"date_to": "2026-09-20T09:00:00"}), {"E1", "E2"})    # a mail at date_to is outside
        self.assertEqual(eids({"attach": "yes"}), {"E1", "E4"})
        self.assertEqual(eids({"attach": "no"}), {"E2", "E3"})
        self.assertEqual((eids({"attach": "pdf"}), eids({"attach": "excel"}), eids({"attach": "image"})), ({"E1"}, {"E4"}, set()))
        ahmet = hm_search.sender_key("Ahmet Demir", "ahmet@x.example")
        self.assertEqual(eids({"sender": ahmet, "attach": "no"}), {"E2"})
        self.assertEqual(hm_search.allowed(self.store, ["F1"], {"date_from": None, "attach": ""}), None)
        with self.store.tx():
            hm_index.ingest(self.store, mail("E8", "Tarihsiz", "Tarih alanı boş.", date=""), 1, REV)
        self.assertNotIn("E8", eids({"date_to": "2026-09-20T09:00:00"}))  # no date: outside every date range
        self.assertIn("E8", eids({"attach": "no"}))
        res = [r["eid"] for r in self.search("kargo faturası", allow=hm_search.allowed(self.store, ["F1"], sep))]
        self.assertEqual((res[0], set(res) <= {"E1", "E3"}), ("E1", True))  # E3: a weak semantic result inside the dates

    def test_similar(self):
        """"Benzerlerini bul": the nearest mails by the stored vectors; the mail itself is not a result; the same topic
        comes first; unknown mail / mail without vectors; filters apply."""
        sim = lambda eid, **kw: hm_search.similar(self.store, eid, ["F1"], REV, **kw)
        res = sim("E1")
        eids = [r["eid"] for r in res["results"]]
        self.assertEqual((res["similar_to"]["eid"], res["similar_to"]["embedded"]), ("E1", True))
        self.assertNotIn("E1", eids)
        self.assertEqual(eids[0], "E2")  # the other "Kargo faturası"
        self.assertEqual(len(eids), 3)  # E2, E3, E4 (E5 is an NDR: not indexed)
        self.assertTrue(res["results"][0]["why"][0].startswith("benzerlik "))
        self.assertEqual(sim("X9")["similar_to"], None)
        with self.store.tx():
            hm_index.ingest(self.store, mail("E7", "Yeni", "Henüz vektörü yok."), 1, REV)
        self.assertEqual((sim("E7")["similar_to"]["embedded"], sim("E7")["results"]), (False, []))
        allow = hm_search.allowed(self.store, ["F1"], {"sender": hm_search.sender_key("Sarah Miller", "sarah@x.example")})
        self.assertEqual([r["eid"] for r in sim("E1", allow=allow)["results"]], ["E3"])

    def test_facets(self):
        f = hm_search.facets(self.store)
        self.assertEqual(f["senders"][0], ["ahmet@x.example", "Ahmet Demir", "ahmet@x.example", 2])
        self.assertEqual(f["recipients"], [["furkan cakir", "Furkan Çakır", 4]])
        self.assertEqual(len(hm_search.facets(self.store, limit=1)["senders"]), 1)

    def test_coverage(self):
        with self.store.tx():
            self.store.set_folder("F1", "\\Gelen Kutusu", 6)
        cov = self.store.coverage(["F1", "F9"])
        self.assertFalse(cov["complete"])
        self.assertEqual(cov["indexed"], 4)
        self.assertIsNone([f for f in cov["folders"] if f["id"] == "F9"][0]["total"])


def att_mail(eid, subject, files, date="2026-09-10T10:00:00", body="Dosya ektedir.", **kw):
    """A mail with fixture attachments: att_meta = [index, name, size, type 1 (file)]."""
    m = mail(eid, subject, body, date, atts=files, **kw)
    m["att_meta"] = [[i, n, (ATT / n).stat().st_size if (ATT / n).exists() else 10, 1] for i, n in enumerate(files, 1)]
    return m


class AttachmentTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.store = Store(os.path.join(self.tmp.name, "t.sqlite"), REV)
        import hm_embed
        self.enc = hm_embed.FakeEncoder()

    def tearDown(self):
        self.store.close()
        self.tmp.cleanup()

    def ingest(self, *mails):
        with self.store.tx():
            for m in mails:
                hm_index.ingest(self.store, m, 1, REV)

    def read_all(self, data=None):
        """Play the add-in: give every pending attachment its fixture file (or data)."""
        out = []
        for rowid, eid, idx, name, kind in self.store.pending_atts(["F1"]):
            blob = data if data is not None else (ATT / name).read_bytes()
            out.append(hm_index.apply_file(self.store, (rowid, idx, name, kind), blob))
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1"], REV)
        return out

    def search(self, q):
        return hm_search.run(self.store, q, NOW, ["F1"], encoder=self.enc, model_rev=REV)["results"]

    def test_formats(self):
        for name, words in [("Stok_Sayim_Eylul.xlsx", "Galvanizli raf dikmesi"), ("Kira_Sozlesmesi_Taslak.docx", "2.400 metrekare"),
                            ("Ceyrek_Sunumu_Q3.pptx", "yüzde 2,7"), ("Fatura_TR4471.pdf", "Soğuk zincir"),
                            ("Kalibrasyon_Notlari.txt", "KX-2291"), ("Ambalaj_Fiyatlari.csv", "Balonlu naylon"),
                            ("Bakim_Arsivi.zip", "Nem ölçer"), ("Ic_Bulten.html", "Gece vardiyası")]:
            state, text = hm_attach.extract_bytes((ATT / name).read_bytes(), name)
            self.assertEqual(state, "ok", name)
            self.assertIn(words, text, name)
        self.assertNotIn("color:#333", hm_attach.extract_bytes((ATT / "Ic_Bulten.html").read_bytes(), "x.html")[1])
        # Third-party samples (licenses in tests/data): legacy Excel (xlrd, BSD) and an Outlook .msg (hfig/mapi, MIT).
        data = Path(__file__).resolve().parent / "data"
        self.assertIn("PROFILEDEF", hm_attach.extract_bytes((data / "profiles.xls").read_bytes(), "profiles.xls")[1])
        state, text = hm_attach.extract_bytes((data / "sample.msg").read_bytes(), "Eklenen ileti", 5)
        self.assertEqual(state, "ok")
        self.assertIn("Manuel Lemos", text)

    @unittest.skipUnless(sys.platform == "win32", "Windows OCR")
    def test_ocr_image_and_scanned_pdf_then_model_loads(self):
        # A fresh process, as in production: OCR first, then the model runtime (with torch this order failed:
        # WinError 1114, winrt's old msvcp140.dll). The worker now loads ONNX Runtime.
        code = ("import sys; sys.path.insert(0, sys.argv[1]); import hm_attach, hm_embed; from pathlib import Path\n"
                "att = Path(sys.argv[2])\n"
                "for n, w in [('Hasarli_Konteyner.jpg', '558213'), ('Ekran_Goruntusu.png', 'E-4012'), ('Tarama_0042.pdf', 'YANGIN')]:\n"
                "    st, t = hm_attach.extract_bytes((att / n).read_bytes(), n)\n"
                "    assert st == 'ok' and w in t, (n, st, t)\n"
                "s = hm_embed.session(sys.argv[3], 1)\n"
                "import numpy as np; print('ort', s.run(None, {'x': np.ones((1, 2), np.float32)})[0].sum())\n")
        r = subprocess.run([sys.executable, "-I", "-B", "-c", code, str(WORKER), str(ATT),
                            str(Path(__file__).resolve().parent / "data" / "identity.onnx")], capture_output=True,
                           text=True, timeout=300)
        self.assertEqual(r.returncode, 0, r.stderr[-1500:])
        self.assertIn("ort 2.0", r.stdout)

    @unittest.skipUnless(importlib.util.find_spec("onnxruntime"), "onnxruntime not installed")
    def test_onnx_runtime_telemetry_off_and_no_torch(self):
        """A fresh process, as in production: hm_embed.Encoder starts ONNX Runtime with telemetry off. Without the
        flag, ORT 1.31 writes a device id and an event database under the user profile (measured on the Mac, 20 Oct);
        with it, nothing. On Windows: OCR first, then ONNX Runtime (the DLL load order that failed with torch,
        WinError 1114). torch and transformers are never imported. Model: a 1-op test graph (identity, input x)."""
        with tempfile.TemporaryDirectory() as home:
            code = ("import sys; sys.path.insert(0, sys.argv[1]); import hm_attach, hm_embed; from pathlib import Path\n"
                    "if sys.platform == 'win32':\n"
                    "    st, t = hm_attach.extract_bytes((Path(sys.argv[2]) / 'Ekran_Goruntusu.png').read_bytes(), 'Ekran_Goruntusu.png')\n"
                    "    assert st == 'ok' and 'E-4012' in t, (st, t)\n"
                    "s = hm_embed.session(sys.argv[3], 1)\n"
                    "import numpy as np; print('ort', s.run(None, {'x': np.ones((1, 2), np.float32)})[0].sum())\n"
                    "print('torch' in sys.modules, 'transformers' in sys.modules)\n")
            env = {k: v for k, v in os.environ.items() if k != "ORT_DISABLE_TELEMETRY"}
            env.update(HOME=home, USERPROFILE=home, LOCALAPPDATA=home, APPDATA=home)
            r = subprocess.run([sys.executable, "-I", "-B", "-c", code, str(WORKER), str(ATT),
                                str(Path(__file__).resolve().parent / "data" / "identity.onnx")],
                               capture_output=True, text=True, timeout=300, env=env)
            self.assertEqual(r.returncode, 0, r.stderr[-1500:])
            self.assertIn("ort 2.0", r.stdout)
            self.assertIn("False False", r.stdout)
            self.assertEqual([str(p) for p in Path(home).rglob("*")], [])  # no device id, no event database

    def test_ocr_types_only_on_windows(self):
        self.assertEqual(hm_attach.readable("foto.jpg"), sys.platform == "win32")
        self.assertFalse(hm_attach.readable("arsiv.rar"))
        self.assertFalse(hm_attach.readable("link.pdf", 4))  # attachment type 4 = link to a file
        self.assertTrue(hm_attach.readable("Toplantı notu", 5))  # 5 = attached Outlook item (.msg)

    def test_zip_member_limit(self):
        import io
        import zipfile
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("big.txt", "a" * (hm_attach.MAX_ZIP_MEMBER + 10))
            z.writestr("small.txt", "küçük metin")
        state, text = hm_attach.extract_bytes(b.getvalue(), "x.zip")
        self.assertEqual((state, text), ("ok", "[small.txt]\nküçük metin"))

    def test_attachment_states(self):
        m = mail("E1", "Ekler", "x", atts=["a.pdf", "image001.png", "b.rar", "c.pdf"])
        m["att_meta"] = [[1, "a.pdf", 100, 1], [2, "image001.png", 3000, 1], [3, "b.rar", 100, 1],
                         [4, "c.pdf", hm_attach.MAX_FILE + 1, 1]]
        self.ingest(m)
        states = [s for (s,) in self.store.db.execute("SELECT state FROM atts ORDER BY idx")]
        self.assertEqual(states, ["pending", "skipped", "unsupported", "too_big"])

    def test_search_finds_text_inside_attachment(self):
        self.ingest(att_mail("A1", "Sayım dosyası", ["Stok_Sayim_Eylul.xlsx"]),
                    att_mail("A2", "Sayım dosyası", ["Stok_Sayim_Ekim.xlsx"]),
                    att_mail("A3", "Fatura", ["Fatura_TR4471.pdf"]))
        self.assertEqual(self.search("galvanizli raf dikmesi"), [])  # not read yet: nothing to find
        self.assertEqual(self.read_all(), ["ok", "ok", "ok"])
        r = self.search("galvanizli raf dikmesi")[0]
        self.assertEqual((r["eid"], r["match_attachment"], r["match_att_index"], r["match_att_member"]),
                         ("A1", "Stok_Sayim_Eylul.xlsx", 1, None))  # xlsx sheet lines "[Sayfa]" are not members
        self.assertIn("Galvanizli raf dikmesi", r["snippet"])
        self.assertIn("ekte", r["why"])
        self.assertEqual(self.search("euro palet sayımı")[0]["eid"], "A2")
        self.assertEqual(self.search("soğuk zincir faturası")[0]["eid"], "A3")
        cov = self.store.coverage(["F1"])
        self.assertEqual((cov["att_total"], cov["att_done"]), (3, 3))

    def test_match_inside_zip_member(self):
        """'Eki aç': the result names the attachment (Outlook index) and the zip member that holds the words."""
        self.ingest(att_mail("A1", "Dosyalar", ["Kalibrasyon_Notlari.txt", "Bakim_Arsivi.zip"], body="Merhaba, ekler ektedir."))
        self.assertEqual(self.read_all(), ["ok", "ok"])
        for q, member in [("nem ölçer sensörü arızası", "Bakim/bakim_notu.docx"), ("arşiv dosyası", "Bakim/okuma.txt")]:
            r = self.search(q)[0]
            self.assertEqual((r["eid"], r["match_attachment"], r["match_att_index"], r["match_att_member"]),
                             ("A1", "Bakim_Arsivi.zip", 2, member), q)
        r = self.search("ekler ektedir")[0]  # a hit in the mail text: no attachment
        self.assertEqual((r["match_attachment"], r["match_att_index"], r["match_att_member"]), (None, None, None))

    def test_zip_member_rule(self):
        text = "[a/Rapor.docx]\nbir\n[Sayfa1]\niki\n[b.txt]\nüç"
        self.assertEqual(hm_search.zip_member("x.zip", text, text.index("iki")), "a/Rapor.docx")  # "[Sayfa1]": no file name
        self.assertEqual(hm_search.zip_member("X.ZIP", text, text.index("üç")), "b.txt")
        self.assertEqual(hm_search.zip_member("x.zip", text, 1), "a/Rapor.docx")  # a hit in the member name line
        self.assertIsNone(hm_search.zip_member("x.zip", "metin", 0))
        self.assertIsNone(hm_search.zip_member("x.xlsx", text, text.index("üç")))
        self.assertIsNone(hm_search.zip_member(None, text, 0))

    def test_same_file_read_once_and_moved_mail_keeps_text(self):
        self.ingest(att_mail("A1", "Fatura", ["Fatura_TR4471.pdf"]), att_mail("A2", "Kopya", ["Fatura_TR4471.pdf"]))
        self.assertEqual(sorted(self.read_all()), ["cached", "ok"])
        with self.store.tx():  # PST move: the mail comes back with a new entry id
            self.store.delete_item(self.store.item_state(["A1"])["A1"][0])
        self.ingest(att_mail("A9", "Fatura", ["Fatura_TR4471.pdf"]))
        self.assertEqual(self.read_all(), ["cached"])
        self.assertEqual({r["eid"] for r in self.search("soğuk zincir taşıma")[:2]}, {"A2", "A9"})

    def test_deleted_mail_text_leaves_index(self):
        self.ingest(att_mail("A1", "Fatura", ["Fatura_TR4471.pdf"]))
        self.read_all()
        with self.store.tx():
            self.store.delete_item(self.store.item_state(["A1"])["A1"][0])
            self.assertEqual(self.store.gc_att_texts(), 1)
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM att_texts").fetchone()[0], 0)

    def test_excluded_folder_leaves_index(self):
        """Ayarlar: a folder excluded from the index. Its mails, attachment text and vectors leave the index, and the
        words leave the file bytes too (FTS optimize + VACUUM + WAL checkpoint). Other folders stay searchable."""
        path = os.path.join(self.tmp.name, "t.sqlite")
        raw = lambda: b"".join(Path(p).read_bytes() for p in (path, path + "-wal") if os.path.exists(p)).lower()
        self.ingest(mail("E1", "Kargo faturası", "Eylül kargo faturası ekte."),
                    dict(att_mail("P1", "Özel sayım", ["Stok_Sayim_Eylul.xlsx"], body="zebragizli kişisel not"), folder="F2"))
        for rowid, eid, idx, name, kind in self.store.pending_atts(["F2"]):
            hm_index.apply_file(self.store, (rowid, idx, name, kind), (ATT / name).read_bytes())
        hm_index.embed_pending(self.store, lambda: self.enc, ["F1", "F2"], REV)
        with self.store.tx():
            self.store.set_folder("F2", "\\Özel", 1)
        count = lambda t: self.store.db.execute(f"SELECT count(*) FROM {t}").fetchone()[0]  # nosemgrep: sqlalchemy-execute-raw-query
        n_vec = count("vectors")
        self.assertTrue(b"zebragizli" in raw() and b"galvanizli" in raw())
        self.assertEqual(self.store.purge_folders(["F2", "F9"]), 1)
        self.assertEqual(self.store.purge_folders([]), 0)  # nothing to do: no second VACUUM
        self.assertEqual([e for (e,) in self.store.db.execute("SELECT entry_id FROM items")], ["E1"])
        self.assertEqual((count("atts"), count("att_texts"), count("folders")), (0, 0, 0))
        self.assertLess(count("vectors"), n_vec)
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM vectors WHERE text_hash NOT IN"
                                               " (SELECT text_hash FROM chunks)").fetchone()[0], 0)
        self.assertFalse(self.store.meta("compact"))
        data = raw()
        self.assertFalse(b"zebragizli" in data or b"galvanizli" in data, "excluded text left in the file")  # no byte dump
        self.assertEqual(hm_search.run(self.store, "zebragizli", NOW, ["F1", "F2"])["results"], [])
        self.assertEqual(self.search("kargo faturası")[0]["eid"], "E1")  # the index still works

    def test_corrupt_file_failed_and_crash_not_retried(self):
        self.ingest(att_mail("A1", "Fatura", ["Fatura_TR4471.pdf"]))
        self.assertEqual(self.read_all(data=b"%PDF-broken"), ["failed"])
        self.ingest(att_mail("A2", "Tablo", ["Stok_Sayim_Eylul.xlsx"]))
        with self.store.tx():  # a worker died while it read this file
            rowid = self.store.item_state(["A2"])["A2"][0]
            self.store.set_att(rowid, 1, "extracting", "deadbeef")
        with self.store.tx():
            self.assertEqual(self.store.fail_stale(), 1)
        self.assertEqual(self.store.cached_att("deadbeef")[0], "failed")
        self.assertEqual(self.store.pending_atts(["F1"]), [])

    def test_mail_change_reads_attachments_again(self):
        self.ingest(att_mail("A1", "Sayım", ["Stok_Sayim_Eylul.xlsx"]))
        self.read_all()
        self.ingest(att_mail("A1", "Sayım", ["Stok_Sayim_Ekim.xlsx"]))  # attachment replaced
        self.assertEqual(self.read_all(), ["ok"])
        keyword = lambda q: hm_search.run(self.store, q, NOW, ["F1"])["results"]  # word matches only
        self.assertEqual(keyword("galvanizli raf"), [])  # the old attachment text is gone
        self.assertEqual(keyword("plastik kasa")[0]["eid"], "A1")

    def test_index_v1_migration(self):
        self.store.close()
        path = os.path.join(self.tmp.name, "v1.sqlite")
        import sqlite3
        db = sqlite3.connect(path)
        db.executescript("CREATE TABLE meta(k TEXT PRIMARY KEY, v TEXT);"
                         "CREATE TABLE items(rowid INTEGER PRIMARY KEY, entry_id TEXT UNIQUE NOT NULL, folder_id TEXT,"
                         " msg_class TEXT, last_mod TEXT, content_hash TEXT, subject TEXT, sender TEXT, sender_addr TEXT,"
                         " recipients TEXT, sent_at TEXT, attachments TEXT, main_text TEXT, n_chunks INTEGER,"
                         " embedded INTEGER DEFAULT 0, model_rev TEXT, index_version INTEGER, seen_gen INTEGER DEFAULT 0);"
                         "CREATE TABLE chunks(id INTEGER PRIMARY KEY, item_rowid INTEGER NOT NULL, seq INTEGER,"
                         " start INTEGER, end INTEGER, text_hash TEXT);"
                         "INSERT INTO meta VALUES('index_version', '1'), ('model_rev', '\"test-rev\"');"
                         "INSERT INTO items(entry_id, last_mod, embedded, model_rev) VALUES('E1', '2026-01-01', 1, 'test-rev');")
        db.commit()
        db.close()
        self.store = Store(path, REV)
        self.assertEqual(self.store.meta("index_version"), 2)
        self.assertIsNone(self.store.item_state(["E1"])["E1"][1])  # read again once, for the attachment list
        self.assertEqual(self.store.db.execute("SELECT embedded FROM items").fetchone()[0], 1)  # vectors kept

    def test_sync_columns_migration(self):
        """A v2 file from before "Güncelle": a folder that was read completely counts as indexed at its last scan; a
        partly read folder does not. A later scan (set_folder) keeps the flag."""
        self.ingest(mail("E1", "a", "a"), mail("E2", "b", "b"), mail("P1", "c", "c", folder="F2"))
        with self.store.tx():
            self.store.set_folder("F1", "\\F1", 2)
            self.store.set_folder("F2", "\\F2", 3)  # 1 of 3 read
        self.store.db.execute("ALTER TABLE items DROP COLUMN added_gen")
        self.store.db.execute("ALTER TABLE folders DROP COLUMN synced_gen")
        scanned = self.store.db.execute("SELECT scanned_at FROM folders WHERE id='F1'").fetchone()[0]
        self.store.close()
        self.store = Store(os.path.join(self.tmp.name, "t.sqlite"), REV)
        self.assertEqual(self.store.synced(["F1", "F2"]), {"F1"})
        self.assertEqual(self.store.db.execute("SELECT synced_gen FROM folders WHERE id='F1'").fetchone()[0], int(scanned * 1000))
        self.assertEqual({a for (a,) in self.store.db.execute("SELECT added_gen FROM items")}, {0})
        with self.store.tx():
            self.store.set_folder("F1", "\\F1 yeni", 5)
        self.assertEqual(self.store.synced(["F1"]), {"F1"})


class ProtocolTests(unittest.TestCase):
    """Start the real worker process (fake encoder) and talk JSON lines to it."""

    def start(self, data, token="s" * 32):
        env = dict(os.environ, HM_SESSION=token, PYTHONIOENCODING="utf-8")
        return subprocess.Popen([sys.executable, "-I", "-B", "-X", "utf8", str(WORKER / "hm_worker.py"), "--data", data,
                                 "--model", "unused", "--fake-encoder"], stdin=subprocess.PIPE,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env, text=True,
                                encoding="utf-8")

    def send(self, p, **msg):
        p.stdin.write(json.dumps(msg, ensure_ascii=False) + "\n")
        p.stdin.flush()

    def events(self, p, until="done", timeout=60):
        out, t0 = [], time.time()
        while time.time() - t0 < timeout:
            line = p.stdout.readline()
            if not line:
                break
            ev = json.loads(line)
            out.append(ev)
            if ev["ev"] in (until, "done"):  # a kept-alive worker sends nothing after 'done' (no EOF)
                break
        return out

    def run_session(self, data, mails, query, folders=("F1",), att_dir=None, **extra):
        """Play the add-in for one session in a new worker process. att_dir: attachments on, files from eval/att."""
        p = self.start(data)
        ev, need = self.play(p, mails, query, folders, att_dir, **extra)
        self.close(p)
        return ev, p.returncode, need

    def play(self, p, mails, query, folders=("F1",), att_dir=None, **extra):
        """One session on a running worker process."""
        if att_dir:
            extra.update(attachments=True, att_dir=att_dir)
        self.send(p, op="begin", session="s" * 32, query=query, now="2026-10-14T12:00:00", top_k=10,
                  store={"id": "S1", "name": "Test"}, folders=[{"id": f, "path": "\\" + f} for f in folders], **extra)
        ev = self.events(p, "hello")
        for f in folders:
            items = [[m["eid"], m["mod"], m["cls"], m["date"]] for m in mails if m["folder"] == f]
            self.send(p, op="manifest", folder=f, items=items, final=True)
        self.send(p, op="scan_done")
        ev += self.events(p, "need")
        need = set(ev[-1]["eids"])
        self.send(p, op="items", items=[m for m in mails if m["eid"] in need])
        self.send(p, op="items_done")
        while True:  # answer need_att until done
            ev += self.events(p, "need_att")
            if ev[-1]["ev"] != "need_att":
                break
            files = []
            for eid, idx, name in ev[-1]["atts"]:
                path = os.path.join(att_dir, "%s-%d.bin" % (eid, idx))
                Path(path).write_bytes((ATT / name).read_bytes())
                files.append([eid, idx, path])
            self.send(p, op="att_files", files=files)
        return ev, need

    def begin_status(self, p, **extra):
        self.send(p, op="begin", session="s" * 32, mode="status", store={"id": "S1"}, folders=[{"id": "F1"}], **extra)
        return self.events(p)

    def test_keep_alive_preload_and_reuse(self):
        """Search window open: one process for many sessions. The model loads once (after the status check, when the
        scope has vectors); the searches after it do not load it. stdin closed -> exit 0."""
        with tempfile.TemporaryDirectory() as d:
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                     mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00")]
            p = self.start(d)
            ev = self.begin_status(p, keep_alive_s=30, preload=True)
            self.assertEqual(([e["ev"] for e in ev], ev[0]["model_ready"]), (["hello", "done"], False))  # no vectors: no preload
            pid = ev[0]["pid"]
            ev, _ = self.play(p, mails, "", mode="index", keep_alive_s=30)  # same process: index run
            self.assertEqual((ev[0]["pid"], ev[-1]["status"], ev[-1]["model_loaded"]), (pid, "ok", True))
            self.close(p)
            self.assertEqual(p.returncode, 0)

            p = self.start(d)  # window opened again: status check + preload
            ev = self.begin_status(p, keep_alive_s=30, preload=True)
            ev += self.events(p, "model")
            ev += self.events(p, "model")
            self.assertEqual([e["ev"] for e in ev], ["hello", "done", "model", "model"])
            self.assertEqual([e["state"] for e in ev[2:]], ["loading", "ready"])
            for q, eid in (("kargo faturası", "E1"), ("cumartesi bakım", "E2")):
                ev, _ = self.play(p, mails, q, keep_alive_s=30)
                self.assertTrue(ev[0]["model_ready"] and ev[-1]["timing"]["model_warm"])
                self.assertNotIn("model", [e.get("phase") for e in ev])  # no model load in this session
                final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
                self.assertEqual((final["semantic"], final["results"][0]["eid"]), (True, eid))
            t0 = time.time()
            self.close(p)
            self.assertEqual(p.returncode, 0)
            self.assertLess(time.time() - t0, 5)

    def test_search_during_preload(self):
        """Window open, then Ara at once: the model loads on a thread (2 s here). The search scans and sends its keyword
        results during the load, waits for that load (one load, not two) and gives semantic final results."""
        with tempfile.TemporaryDirectory() as d:
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                     mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00")]
            self.run_session(d, mails, "", mode="index")
            env = dict(os.environ, HM_SESSION="s" * 32, PYTHONIOENCODING="utf-8", HM_FAKE_LOAD_S="2")
            p = subprocess.Popen([sys.executable, "-I", "-B", "-X", "utf8", str(WORKER / "hm_worker.py"), "--data", d,
                                  "--model", "unused", "--fake-encoder"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE, env=env, text=True, encoding="utf-8")
            t0 = time.time()
            self.assertEqual([e["ev"] for e in self.begin_status(p, keep_alive_s=30, preload=True)], ["hello", "done"])
            ev, _ = self.play(p, mails, "kargo faturası", keep_alive_s=30)
            ev += self.events(p, "model") if "ready" not in [e.get("state") for e in ev] else []
            names = [e["ev"] if e["ev"] != "results" else e["stage"] for e in ev]
            self.assertLess(names.index("keyword"), names.index("final"))
            self.assertEqual([e["state"] for e in ev if e["ev"] == "model"], ["loading", "ready"])
            self.assertLess(names.index("keyword"), names.index("model", names.index("model") + 1))  # before "ready"
            done = [e for e in ev if e["ev"] == "done"][0]
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertEqual((done["status"], final["semantic"], final["results"][0]["eid"]), ("ok", True, "E1"))
            self.assertLess(time.time() - t0, 3.8)  # one 2 s load, not two
            self.close(p)
            self.assertEqual(p.returncode, 0)

    def test_keep_alive_ends(self):
        """No new session within keep_alive_s -> exit 0. An error in a kept session -> exit 2. Cancel -> exit 3."""
        with tempfile.TemporaryDirectory() as d:
            p = self.start(d)
            t0 = time.time()
            self.begin_status(p, keep_alive_s=1)
            self.assertEqual(p.wait(timeout=10), 0)
            self.assertLess(time.time() - t0, 6)
            self.close(p)
            p = self.start(d)
            self.begin_status(p, keep_alive_s=30)
            self.send(p, op="begin", session="wrong", store={"id": "S1"}, folders=[{"id": "F1"}])
            ev = self.events(p)
            self.assertEqual((ev[-1]["status"], ev[-1]["detail"]), ("error", "session id mismatch"))
            self.assertEqual(p.wait(timeout=10), 2)
            self.close(p)
            p = self.start(d)
            self.begin_status(p, keep_alive_s=30)
            self.send(p, op="begin", session="s" * 32, query="x", store={"id": "S1"}, folders=[{"id": "F1"}],
                      keep_alive_s=30)
            self.events(p, "hello")
            self.send(p, op="cancel")
            self.assertEqual(self.events(p)[-1]["status"], "cancelled")
            self.assertEqual(p.wait(timeout=10), 3)
            self.close(p)

    def status(self, data, folders=("F1",)):
        p = self.start(data)
        self.send(p, op="begin", session="s" * 32, mode="status", store={"id": "S1"}, attachments=True,
                  folders=[{"id": f} for f in folders])
        ev = self.events(p)
        self.close(p)
        return ev, p.returncode

    def test_status_mode_reads_nothing(self):
        with tempfile.TemporaryDirectory() as d:
            ev, code = self.status(d)
            self.assertEqual(([e["ev"] for e in ev], ev[-1]["status"], code), (["hello", "done"], "ok", 0))
            self.assertIsNone(ev[-1]["coverage"]["folders"][0]["total"])
            self.assertFalse(os.path.exists(os.path.join(d, "index")))  # unknown store: no index file
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                     mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00")]
            self.run_session(d, mails, "", mode="index")
            ev, code = self.status(d, ("F1", "F9"))
            cov = ev[-1]["coverage"]
            f1 = cov["folders"][0]
            self.assertEqual((f1["id"], f1["total"], f1["indexed"], f1["read"], f1["embedded"]), ("F1", 2, 2, 2, 2))
            self.assertEqual(cov["folders"][1]["id"], "F9")  # the order of the request
            self.assertFalse(cov["complete"])  # F9 never scanned

    def test_index_mode_no_query_no_results(self):
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as att_dir:
            mails = [att_mail("A1", "Sayım dosyası", ["Stok_Sayim_Eylul.xlsx"]),
                     mail("E2", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00")]
            ev, code, need = self.run_session(d, mails, "galvanizli", att_dir=att_dir, mode="index")
            self.assertEqual((code, ev[-1]["status"], ev[-1]["mode"]), (0, "ok", "index"))
            self.assertNotIn("results", [e["ev"] for e in ev])
            self.assertTrue(ev[-1]["coverage"]["complete"])
            self.assertEqual((ev[-1]["coverage"]["att_done"], ev[-1]["coverage"]["att_total"]), (1, 1))
            embed = [e for e in ev if e["ev"] == "progress" and e["phase"] == "embed"]
            self.assertEqual((embed[-1]["done"], embed[-1]["total"], embed[-1]["eta_s"]), (3, 3, 0))  # 2 body + 1 attachment chunk

    def test_search_budget_zero_and_stopped_read(self):
        """Time-limited search: no attachment reading and no embedding after the limit; mails that the add-in did
        not read stay searchable with their old text and are asked for again by the next scan."""
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as att_dir:
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                     mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00")]
            self.run_session(d, mails, "", mode="index")
            new = att_mail("A3", "Sayım dosyası", ["Stok_Sayim_Eylul.xlsx"], date="2026-10-01T09:00:00")
            changed = [dict(m, mod="2026-10-02T09:00:00", body=m["body"] + " Güncel.") for m in mails]
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, query="kargo faturası", now="2026-10-14T12:00:00",
                      store={"id": "S1"}, folders=[{"id": "F1"}], index_budget_s=0, attachments=True, att_dir=att_dir)
            self.events(p, "hello")
            self.send(p, op="manifest", folder="F1", final=True,
                      items=[[m["eid"], m["mod"], m["cls"], m["date"]] for m in changed + [new]])
            self.send(p, op="scan_done")
            self.assertEqual(self.events(p, "need")[-1]["eids"], ["A3", "E2", "E1"])  # newest first
            self.send(p, op="items", items=[new])  # the add-in read one mail, then its time was up
            self.send(p, op="items_done", stopped=True)
            ev = self.events(p)
            self.close(p)
            self.assertEqual((p.returncode, ev[-1]["status"]), (0, "ok"))
            self.assertNotIn("need_att", [e["ev"] for e in ev])
            done = ev[-1]
            self.assertTrue(done["stats"]["read_stopped"] and done["stats"]["embed"]["stopped"])
            cov = done["coverage"]
            self.assertEqual((cov["indexed"], cov["read"], cov["embedded"], cov["att_done"], cov["att_total"]),
                             (3, 1, 2, 0, 1))
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertTrue(final["semantic"])  # the old vectors of E1, E2 still work
            self.assertEqual(final["results"][0]["eid"], "E1")  # still found with its old text
            self.assertEqual(final["results"][0]["hl_subject"], [[0, 5], [6, 8]])  # "Kargo", "faturası"
            # The next run asks for E1 and E2 again and reads the attachment of A3.
            ev, code, need = self.run_session(d, changed + [new], "", att_dir=att_dir, mode="index")
            self.assertEqual(need, {"E1", "E2"})
            self.assertTrue(ev[-1]["coverage"]["complete"])

    def test_attachments_pending_without_changed_mails(self):
        """No mail changed (empty need) but attachments wait: items_done comes before need_att."""
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as att_dir:
            mails = [att_mail("A1", "Sayım dosyası", ["Stok_Sayim_Eylul.xlsx"])]
            self.run_session(d, mails, "x", att_dir=att_dir, index_budget_s=0)
            ev, code, need = self.run_session(d, mails, "galvanizli raf", att_dir=att_dir)
            self.assertEqual((need, code, ev[-1]["status"]), (set(), 0, "ok"))
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertEqual(final["results"][0]["match_attachment"], "Stok_Sayim_Eylul.xlsx")

    def test_exclude_folder_protocol(self):
        """begin.exclude: the status check at the window open deletes the excluded folder from the index; later
        sessions do not see it. An excluded folder in the scope is an error. A status check with an empty scope (every
        folder excluded) works and makes no file for an unknown store; other modes still need a scope."""
        with tempfile.TemporaryDirectory() as d:
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, mode="status", store={"id": "S0"}, folders=[], exclude=["F2"])
            ev = self.events(p)
            self.close(p)
            self.assertEqual((ev[-1]["status"], ev[-1]["excluded_removed"], p.returncode), ("ok", 0, 0))
            self.assertFalse(os.path.exists(os.path.join(d, "index")))
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                     mail("P1", "Özel", "zebragizli kişisel not.", "2026-09-20T09:00:00", folder="F2")]
            self.run_session(d, mails, "", folders=("F1", "F2"), mode="index")
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, mode="status", store={"id": "S1"}, folders=[{"id": "F1"}],
                      exclude=["F2"])
            ev = self.events(p)
            self.close(p)
            self.assertEqual((ev[-1]["status"], ev[-1]["excluded_removed"], ev[-1]["coverage"]["indexed"]), ("ok", 1, 1))
            ev, code, need = self.run_session(d, mails, "zebragizli kişisel not", exclude=["F2"])
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertEqual((code, need, ev[-1]["excluded_removed"]), (0, set(), 0))
            self.assertTrue(all(r["eid"] != "P1" for r in final["results"]))
            data = b"".join(f.read_bytes() for f in Path(d, "index").iterdir())
            self.assertFalse(b"zebragizli" in data, "excluded text left in the file")
            for folders, exclude, mode, detail in (([{"id": "F1"}], ["F1"], "status", "excluded folder in scope"),
                                                   ([], ["F2"], "search", "empty folder scope")):
                p = self.start(d)
                self.send(p, op="begin", session="s" * 32, mode=mode, store={"id": "S1"}, folders=folders, exclude=exclude)
                ev = self.events(p)
                self.close(p)
                self.assertEqual((ev[-1]["status"], ev[-1]["detail"], p.returncode), ("error", detail, 2))

    def test_update_mode_processes_only_changes(self):
        """"Güncelle": after an index run, only the changes: changed mails, new mails of completely indexed folders, and
        their attachments and vectors. Old backlog (an attachment of an old mail, a mail of a never-indexed folder) waits
        for "İndekslemeyi başlat". Deleted mails leave the index. A search scan keeps the folder flag."""
        import sqlite3
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as att_dir:
            e1 = mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00")
            e2 = mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00")
            a1 = att_mail("A1", "Sayım dosyası", ["Stok_Sayim_Eylul.xlsx"], date="2026-09-01T09:00:00")
            n1 = mail("N1", "Teklif", "Raf sistemi teklifi ektedir.", "2026-10-01T09:00:00")
            b1 = mail("B1", "Bülten", "Haftalık bülten.", "2026-10-02T09:00:00", folder="F2")
            n2 = att_mail("N2", "Fatura", ["Fatura_TR4471.pdf"], date="2026-10-05T09:00:00")
            b2 = mail("B2", "Bülten iki", "İkinci bülten.", "2026-10-06T09:00:00", folder="F2")
            e2c = dict(e2, mod="2026-10-03T09:00:00", body="Pazar sabahı bakım.")
            both = ("F1", "F2")
            self.run_session(d, [e1, e2, a1], "", mode="index")  # F1 indexed; attachments off: A1's file waits
            _, _, need = self.run_session(d, [e1, e2, a1, n1, b1], "kargo", folders=both, index_budget_s=0)
            self.assertEqual(need, {"N1", "B1"})  # read by a search, not embedded (no indexing time)
            now = [e2c, a1, n1, n2, b1, b2]
            ev, code, need = self.run_session(d, now, "", folders=both, att_dir=att_dir, mode="update")
            done = ev[-1]
            self.assertEqual((code, done["status"], done["mode"]), (0, "ok", "update"))
            self.assertEqual(need, {"E2", "N2"})  # B2: new in a never-indexed folder
            self.assertEqual((done["stats"]["unsynced_skipped"], done["stats"]["removed"]), (1, 1))
            self.assertEqual([a for e in ev if e["ev"] == "need_att" for a in e["atts"]], [["N2", 1, "Fatura_TR4471.pdf"]])
            self.assertNotIn("results", [e["ev"] for e in ev])
            db = sqlite3.connect(next(Path(d, "index").glob("*.sqlite")))
            embedded = dict(db.execute("SELECT entry_id, embedded FROM items"))
            att_state = db.execute("SELECT a.state FROM atts a JOIN items i ON i.rowid=a.item_rowid WHERE entry_id='A1'").fetchone()[0]
            db.close()
            self.assertEqual(embedded, {"E2": 1, "A1": 1, "N1": 1, "N2": 1, "B1": 0})
            self.assertEqual(att_state, "pending")
            self.run_session(d, now, "", folders=both, att_dir=att_dir, mode="index")  # F2 indexed now
            ev, code, need = self.run_session(d, now, "", folders=both, att_dir=att_dir, mode="update")
            self.assertEqual((need, ev[-1]["stats"]["unsynced_skipped"], ev[-1]["stats"]["embed"]["embedded_chunks"]),
                             (set(), 0, 0))
            self.assertFalse(ev[-1]["model_loaded"] or "model" in [e.get("phase") for e in ev])  # nothing to do: no model
            self.assertTrue(ev[-1]["coverage"]["complete"])

    def test_semantic_stage_before_embedding(self):
        """New mails without vectors: semantic results with the indexed mails come before the embedding, then the final
        results. The final list is the same as a search on the index after the session (the quality does not change).
        Nothing to embed: no semantic stage (one list less)."""
        import hm_embed
        with tempfile.TemporaryDirectory() as d:
            old = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                   mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00")]
            self.run_session(d, old, "", mode="index")
            new = old + [mail("N1", "Kargo faturası ekim", "Ekim kargo faturası ektedir.", "2026-10-10T09:00:00")]
            ev, code, need = self.run_session(d, new, "kargo faturası", index_budget_s=20)
            names = [e["stage"] if e["ev"] == "results" else e.get("phase") or e["ev"] for e in ev]
            self.assertEqual((code, need), (0, {"N1"}))
            self.assertLess(names.index("keyword"), names.index("semantic"))
            self.assertLess(names.index("semantic"), names.index("embed"))
            self.assertLess(names.index("embed"), names.index("final"))
            res = {e["stage"]: e for e in ev if e["ev"] == "results"}
            self.assertTrue(res["semantic"]["semantic"])
            cos = lambda stage: {r["eid"]: r["cos"] for r in res[stage]["results"]}
            self.assertIsNone(cos("semantic")["N1"])  # read, found by its words; no vector yet
            self.assertIsNotNone(cos("final")["N1"])
            st = Store(str(next(Path(d, "index").glob("*.sqlite"))), hm_embed.MODEL_REVISION)
            ref = hm_search.run(st, "kargo faturası", NOW, ["F1"], encoder=hm_embed.FakeEncoder(),
                                model_rev=hm_embed.MODEL_REVISION, dim=hm_embed.DIM)
            st.close()
            self.assertEqual(res["final"]["results"], ref["results"])
            ev, code, need = self.run_session(d, new, "kargo faturası", index_budget_s=20)
            self.assertEqual([e["stage"] for e in ev if e["ev"] == "results"], ["keyword", "final"])
            ev, code, need = self.run_session(d, new + [mail("N2", "Kargo", "Kasım kargo.", "2026-11-01T09:00:00")],
                                              "kargo faturası", index_budget_s=0)  # no indexing time: no semantic stage
            self.assertEqual([e["stage"] for e in ev if e["ev"] == "results"], ["keyword", "final"])

    def test_scan_moved_and_deleted(self):
        """The scan writes only the moved mails (folder_id) and deletes the mails that no manifest holds. A moved mail
        is not read again and keeps its vectors."""
        import sqlite3
        with tempfile.TemporaryDirectory() as d:
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00"),
                     mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00"),
                     mail("E3", "Toplantı", "Bütçe toplantısı.", "2026-09-22T09:00:00", folder="F2")]
            self.run_session(d, mails, "", folders=("F1", "F2"), mode="index")
            now = [dict(mails[0], folder="F2"), mails[2]]  # E1 moved to F2, E2 deleted
            ev, code, need = self.run_session(d, now, "kargo faturası", folders=("F1", "F2"))
            done = ev[-1]
            self.assertEqual((code, need, done["stats"]["removed"], done["stats"]["embed"]["embedded_chunks"]),
                             (0, set(), 1, 0))
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertEqual((final["results"][0]["eid"], final["results"][0]["folder"]), ("E1", "F2"))
            db = sqlite3.connect(next(Path(d, "index").glob("*.sqlite")))
            self.assertEqual(dict(db.execute("SELECT entry_id, folder_id FROM items")), {"E1": "F2", "E3": "F2"})
            db.close()
            # Only F1 listed: E1 (now in F2) stays; a folder that is not in the scan keeps its mails.
            ev, code, need = self.run_session(d, now, "kargo", folders=("F1",))
            self.assertEqual((code, ev[-1]["stats"]["removed"], ev[-1]["coverage"]["indexed"]), (0, 0, 0))
            ev, _ = self.status(d, ("F2",))
            self.assertEqual(ev[-1]["coverage"]["indexed"], 2)

    def test_unknown_mode_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            for extra, detail in (({"mode": "delete"}, "unknown mode"), ({"filters": ["x"]}, "bad filters"),
                                  ({"filters": {"sender": 5}}, "bad filters"), ({"filters": {"attach": "exe"}}, "bad filters"),
                                  ({"filters": {"date_from": "geçen ay"}}, "bad filters")):
                p = self.start(d)
                self.send(p, op="begin", session="s" * 32, store={"id": "S"}, folders=[{"id": "F1"}], **extra)
                ev = self.events(p)
                self.close(p)
                self.assertEqual((ev[-1]["status"], ev[-1]["detail"], p.returncode), ("error", detail, 2))

    def test_similar_mode(self):
        """Worker mode similar: hello -> results -> done, no scan, no model; unknown store: no index file; a missing
        similar_eid is an error."""
        with tempfile.TemporaryDirectory() as d:
            begin = lambda p, **kw: self.send(p, op="begin", session="s" * 32, mode="similar", store={"id": "S1"},
                                              folders=[{"id": "F1"}], **kw)
            p = self.start(d)
            begin(p, similar_eid="E1")
            ev = self.events(p)
            self.close(p)
            self.assertEqual(([e["ev"] for e in ev], ev[1]["similar_to"], p.returncode), (["hello", "results", "done"], None, 0))
            self.assertFalse(os.path.exists(os.path.join(d, "index")))
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası ekte.", "2026-09-12T09:00:00"),
                     mail("E2", "Kargo faturası", "Ekim kargo faturası ekte.", "2026-10-12T09:00:00"),
                     mail("E3", "Sunucu bakımı", "Cumartesi gecesi planlı bakım.", "2026-09-20T09:00:00")]
            self.run_session(d, mails, "", mode="index")
            p = self.start(d)
            begin(p, similar_eid="E1", keep_alive_s=30)
            ev = self.events(p)
            self.assertEqual(([e["ev"] for e in ev], ev[-1]["status"], ev[-1]["model_loaded"]), (["hello", "results", "done"], "ok", False))
            self.assertEqual([r["eid"] for r in ev[1]["results"]], ["E2", "E3"])
            self.assertNotIn("model", [e.get("phase") for e in ev])
            begin(p)  # the next session in the same process: no similar_eid
            ev = self.events(p)
            self.close(p)
            self.assertEqual((ev[-1]["status"], ev[-1]["detail"], p.returncode), ("error", "similar_eid missing", 2))

    def test_search_filters_and_facets_protocol(self):
        """done.facets lists the senders and recipients of the index file (also in status mode); begin.filters narrows
        the keyword and the final results."""
        with tempfile.TemporaryDirectory() as d:
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası.", "2026-09-12T09:00:00", "Ahmet Demir"),
                     mail("E2", "Kargo faturası", "Ekim kargo faturası.", "2026-10-12T09:00:00", "Zeynep Arslan", folder="F2")]
            self.run_session(d, mails, "", folders=("F1", "F2"), mode="index")
            ev, _ = self.status(d)  # scope F1 only: the facets are the whole file
            self.assertEqual([s[0] for s in ev[-1]["facets"]["senders"]], ["ahmet@x.example", "zeynep@x.example"])
            ev, code, need = self.run_session(d, mails, "kargo faturası", folders=("F1", "F2"),
                                              filters={"sender": "zeynep@x.example", "recipient": None})
            res = {e["stage"]: [r["eid"] for r in e["results"]] for e in ev if e["ev"] == "results"}
            self.assertEqual((code, res), (0, {"keyword": ["E2"], "final": ["E2"]}))
            # Date filter as the add-in sends it (a date only, and a time with an offset: the local time counts).
            for f in ({"date_to": "2026-10-01"}, {"date_to": "2026-10-01T00:00:00+03:00", "attach": "no"}):
                ev, code, _ = self.run_session(d, mails, "kargo faturası", folders=("F1", "F2"), filters=f)
                res = {e["stage"]: [r["eid"] for r in e["results"]] for e in ev if e["ev"] == "results"}
                self.assertEqual((code, res), (0, {"keyword": ["E1"], "final": ["E1"]}), f)

    def close(self, p):
        for f in (p.stdin, p.stdout, p.stderr):
            if not f.closed:
                f.close()
        p.wait(timeout=30)

    def test_full_session_incremental_and_exit(self):
        with tempfile.TemporaryDirectory() as d:
            mails = [mail("E1", "Kargo faturası", "Eylül kargo faturası, 4.250 TL.", "2026-09-12T09:00:00"),
                     mail("E2", "Sunucu bakımı", "Cumartesi bakım.", "2026-09-20T09:00:00"),
                     mail("E3", "Görev", "x", cls="IPM.Task")]
            ev, code, need = self.run_session(d, mails, "geçen ay kargo faturası")
            self.assertEqual(code, 0)
            self.assertEqual(need, {"E1", "E2"})
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertEqual(final["results"][0]["eid"], "E1")
            self.assertTrue(final["coverage"]["complete"])
            self.assertEqual(ev[-1]["status"], "ok")
            # Second run: E2 changed, E1 deleted -> only E2 is read again, E1 is removed.
            mails2 = [dict(mails[1], mod="2026-10-01T09:00:00", body="Pazar sabahı bakım.")]
            ev, code, need = self.run_session(d, mails2, "bakım")
            self.assertEqual(need, {"E2"})
            done = ev[-1]
            self.assertEqual(done["stats"]["removed"], 1)
            self.assertEqual(done["stats"]["embed"]["embedded_chunks"], 1)
            # Third run: nothing changed -> nothing read, no new vectors, model loads only for the query.
            ev, code, need = self.run_session(d, mails2, "bakım")
            self.assertEqual(need, set())
            self.assertEqual(ev[-1]["stats"]["embed"]["embedded_chunks"], 0)

    def test_attachment_phase(self):
        """need_att -> the add-in saves files into att_dir -> the worker reads and deletes them."""
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as att_dir:
            mails = [att_mail("A1", "Sayım dosyası", ["Stok_Sayim_Eylul.xlsx", "Fatura_TR4471.pdf"]),
                     att_mail("A2", "Toplantı", [], body="Toplantı yarın.")]
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, query="galvanizli raf dikmesi", now="2026-10-14T12:00:00",
                      store={"id": "S1"}, folders=[{"id": "F1", "path": "\\F1"}], attachments=True, att_dir=att_dir)
            self.events(p, "hello")
            self.send(p, op="manifest", folder="F1", items=[[m["eid"], m["mod"], m["cls"], m["date"]] for m in mails],
                      final=True)
            self.send(p, op="scan_done")
            self.events(p, "need")
            self.send(p, op="items", items=mails)
            self.send(p, op="items_done")
            ev = self.events(p, "need_att")
            asked = ev[-1]["atts"]
            self.assertEqual(sorted(a[2] for a in asked), ["Fatura_TR4471.pdf", "Stok_Sayim_Eylul.xlsx"])
            files = []
            for eid, idx, name in asked:
                path = os.path.join(att_dir, "%s-%d.bin" % (eid, idx))
                Path(path).write_bytes((ATT / name).read_bytes())
                files.append([eid, idx, path])
            self.send(p, op="att_files", files=files)
            ev = self.events(p, "done")
            self.close(p)
            self.assertEqual(p.returncode, 0)
            self.assertEqual(os.listdir(att_dir), [])  # every file deleted after reading
            final = [e for e in ev if e["ev"] == "results" and e["stage"] == "final"][0]
            self.assertEqual(final["results"][0]["eid"], "A1")
            self.assertEqual(final["results"][0]["match_attachment"], "Stok_Sayim_Eylul.xlsx")
            self.assertTrue(final["coverage"]["complete"])
            self.assertEqual(ev[-1]["stats"]["att"]["ok"], 2)

    def test_attachment_path_outside_att_dir_rejected(self):
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as att_dir:
            m = att_mail("A1", "Sayım", ["Stok_Sayim_Eylul.xlsx"])
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, query="x", now="2026-10-14T12:00:00", store={"id": "S1"},
                      folders=[{"id": "F1"}], attachments=True, att_dir=att_dir)
            self.events(p, "hello")
            self.send(p, op="manifest", folder="F1", items=[[m["eid"], m["mod"], m["cls"], m["date"]]], final=True)
            self.send(p, op="scan_done")
            self.events(p, "need")
            self.send(p, op="items", items=[m])
            self.send(p, op="items_done")
            self.events(p, "need_att")
            outside = os.path.join(d, "victim.txt")
            Path(outside).write_text("do not delete")
            self.send(p, op="att_files", files=[["A1", 1, outside]])
            ev = self.events(p)
            self.close(p)
            self.assertEqual((ev[-1]["status"], p.returncode), ("error", 2))
            self.assertTrue(os.path.exists(outside))

    def test_bad_session_id_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.start(d, token="t" * 32)
            self.send(p, op="begin", session="wrong", query="x", store={"id": "S"}, folders=[{"id": "F1"}])
            ev = self.events(p)
            self.close(p)
            self.assertEqual(ev[-1]["status"], "error")
            self.assertEqual(p.returncode, 2)
            self.assertEqual(os.listdir(d), [])  # nothing written before the check

    def test_cancel_and_parent_death_end_process(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, query="x", store={"id": "S"}, folders=[{"id": "F1"}])
            self.events(p, "hello")
            t0 = time.time()
            self.send(p, op="cancel")
            ev = self.events(p)
            self.close(p)
            self.assertEqual((ev[-1]["status"], p.returncode), ("cancelled", 3))
            self.assertLess(time.time() - t0, 5)
            p = self.start(d)  # parent closes stdin without a word (Outlook crash)
            self.send(p, op="begin", session="s" * 32, query="x", store={"id": "S"}, folders=[{"id": "F1"}])
            self.events(p, "hello")
            t0 = time.time()
            self.close(p)
            self.assertEqual(p.returncode, 3)
            self.assertLess(time.time() - t0, 5)

    def test_idle_timeout(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.start(d)
            self.send(p, op="begin", session="s" * 32, query="x", store={"id": "S"}, folders=[{"id": "F1"}],
                      idle_timeout_s=5)
            t0 = time.time()
            ev = self.events(p, timeout=30)
            self.close(p)
            self.assertEqual(ev[-1]["status"], "error")
            self.assertLess(time.time() - t0, 20)

    def test_no_mail_text_on_stderr(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.start(d)
            secret = "GizliKelime987"
            self.send(p, op="begin", session="s" * 32, query=secret, now="2026-10-14T12:00:00",
                      store={"id": "S"}, folders=[{"id": "F1"}])
            self.send(p, op="manifest", folder="F1", items=[["E1", "m", "IPM.Note", "2026-01-01"]], final=True)
            self.send(p, op="scan_done")
            self.send(p, op="items", items=[mail("E1", secret, secret)])
            self.send(p, op="items_done")
            self.events(p)
            p.stdin.close()
            err = p.stderr.read()
            self.close(p)
            self.assertNotIn(secret, err)
            self.assertEqual([x for x in os.listdir(d) if x != "index"], [])


if __name__ == "__main__":
    unittest.main()
