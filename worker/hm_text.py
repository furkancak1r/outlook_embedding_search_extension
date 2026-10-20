"""Text helpers: Turkish-aware folding, quote/signature/disclaimer removal, chunking, hashing.

fold() keeps the string length, so positions in the folded text are also positions in the original text.
"""
import hashlib
import json
import re
import unicodedata


class _FoldMap(dict):
    """str.translate table: lower case + strip diacritics, one output char for each input char."""

    def __missing__(self, cp):
        c = chr(cp)
        low = c.lower()
        if len(low) != 1:
            low = c
        base = "".join(ch for ch in unicodedata.normalize("NFKD", low) if not unicodedata.combining(ch))
        r = base if len(base) == 1 else low
        r = _SPECIAL.get(r, r)
        self[cp] = r
        return r


_SPECIAL = {"ı": "i", "ł": "l", "ø": "o", "đ": "d", "’": "'", "‘": "'", "`": "'", "´": "'"}
_FOLD = _FoldMap()
_FOLD[ord("İ")] = "i"
_FOLD[ord("I")] = "i"


def fold(s):
    """Lower case with Turkish rules (I/İ/ı/i all become i) and ASCII-like letters (ş->s, ğ->g, ...)."""
    return s.translate(_FOLD) if s else ""


_TOKEN = re.compile(r"[^\W_]+")


def tokens(folded):
    """Split like the SQLite unicode61 tokenizer: runs of letters and digits."""
    return _TOKEN.findall(folded)


# --- reply chain, signature and disclaimer removal -------------------------------------------

_HDR_LINE = re.compile(
    r"^\s*(?:-{2,}\s*(?:original message|orijinal ileti|ozgun ileti|forwarded message|iletilen ileti|ileti)\s*-{2,}|_{10,})\s*$")
_ON_WROTE = re.compile(r"^\s*(?:on\s.{4,200}\swrote:|.{4,200}\s(?:tarihinde|zamaninda).{0,200}\s(?:sunu yazdi|yazdi):?)\s*$")
_FROM = re.compile(r"^\s*\*?(?:from|kimden|gonderen)\s*:\*?\s*\S")
_SENT = re.compile(r"^\s*\*?(?:sent|date|gonderildi|gonderilme tarihi|tarih|to|kime)\s*:\*?\s*\S")
_SIGNOFF = re.compile(
    r"^\s*(?:saygilarimla|saygilarimizla|saygilar|iyi calismalar|iyi gunler|selamlar|sevgiler|tesekkurler|tesekkur ederim"
    r"|kind regards|best regards|warm regards|regards|best wishes|best|thanks|thank you|many thanks|cheers|sincerely"
    r"|yours sincerely|br)\s*[,.!]?\s*$")
_MOBILE = re.compile(
    r"^\s*(?:sent from my \w+|iphone'?umdan gonderildi|ipad'?imden gonderildi|sent from outlook for \w+|get outlook for \w+"
    r"|android icin outlook'?tan gonderildi|\w+ icin outlook'?u edinin)\s*$")
_DISCLAIMER_WORDS = ("gizli", "yasal uyari", "sorumluluk", "muhatab", "alicisi degilseniz", "confidential",
                     "intended recipient", "disclaimer", "privileged", "unauthorized", "kvkk", "kisisel veri",
                     "imha", "bilgilendirme amaci", "virus")
_FORWARD_SUBJECT = re.compile(r"^\s*(?:fw|fwd|ilt|iletilen)\s*:")


def _is_disclaimer(folded_para):
    return len(folded_para) > 120 and sum(w in folded_para for w in _DISCLAIMER_WORDS) >= 2


def split_body(body, subject=""):
    """Return (main, quoted).

    main: the new text without quoted reply chain, signature, mobile footer and legal disclaimer.
    quoted: the removed reply chain (kept for low-weight keyword search only).
    ponytail: heuristic line rules for Outlook/Gmail TR+EN formats; unusual layouts stay in main text.
    """
    body = (body or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = body.split("\n")
    flines = [fold(x) for x in lines]
    forward = bool(_FORWARD_SUBJECT.match(fold(subject or "")))

    keep, quoted = [], []
    cut = len(lines)
    if not forward:  # in a forward the forwarded text is the content
        for i, fl in enumerate(flines):
            if _HDR_LINE.match(fl) or _ON_WROTE.match(fl) or (
                    _FROM.match(fl) and any(_SENT.match(x) for x in flines[i + 1:i + 5])):
                cut = i
                break
    for i in range(cut):
        (quoted if flines[i].lstrip().startswith(">") else keep).append(i)
    quoted_text = "\n".join([lines[i].lstrip("> ") for i in quoted] + lines[cut:]).strip()

    keep = [i for i in keep if not _MOBILE.match(flines[i])]
    # Signature: last sign-off line in the last 15 kept lines, if real text stays before it.
    tail = keep[-15:]
    for i in reversed(tail):
        if _SIGNOFF.match(flines[i]):
            before = [j for j in keep if j < i]
            if sum(len(flines[j].strip()) for j in before) >= 20:
                keep = before
            break
    main = "\n".join(lines[i] for i in keep)
    paras = re.split(r"\n\s*\n", main)
    main = "\n\n".join(p for p in paras if not _is_disclaimer(fold(p)))
    main = re.sub(r"\n{3,}", "\n\n", main).strip()
    return main, quoted_text


# --- chunking -------------------------------------------------------------------------------

def chunk_spans(text, max_chars=1200, overlap=150, max_chunks=12):
    """Split text into (start, end) spans at paragraph/sentence/space limits.

    ponytail: char-based limit (~300-450 tokens) instead of exact token counts; the encoder truncates at
    MAX_DOC_TOKENS as a guard. Text after max_chunks spans stays keyword-searchable only.
    """
    n = len(text)
    if n == 0:
        return [(0, 0)]
    spans, a = [], 0
    while a < n and len(spans) < max_chunks:
        b = min(n, a + max_chars)
        if b < n:
            window = text[a:b]
            cutpos = max(window.rfind("\n\n"), window.rfind(". "), window.rfind("\n"))
            if cutpos < max_chars // 2:
                cutpos = window.rfind(" ")
            if cutpos > max_chars // 3:
                b = a + cutpos + 1
        spans.append((a, b))
        if b >= n:
            break
        a = max(b - overlap, a + 1)
        while a < n and not text[a - 1].isspace() and a < b:  # start chunks at a word start
            a += 1
    return spans


def content_hash(fields):
    return hashlib.sha256(json.dumps(fields, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:32]


def text_hash(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:32]
