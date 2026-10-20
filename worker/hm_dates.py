"""Approximate date expressions (Turkish and English) -> soft date windows.

parse(query, now) -> (windows, spans, keep)
  windows: [(start, end_exclusive, label)]
  spans:   character spans of the date phrases in the query (the caller removes them from the text query)
  keep:    words that the caller keeps as keywords (years, the name-like month "Kasım")
The search uses the windows only as a soft boost, never as a hard filter.
"""
import re
from datetime import datetime, timedelta

from hm_text import fold

D = timedelta(days=1)
MONTHS_TR = {"ocak": 1, "ocag": 1, "subat": 2, "mart": 3, "nisan": 4, "mayis": 5, "haziran": 6, "temmuz": 7,
             "agustos": 8, "eylul": 9, "ekim": 10, "kasim": 11, "aralik": 12, "aralig": 12}
MONTHS_EN = {"january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6, "july": 7, "august": 8,
             "september": 9, "october": 10, "november": 11, "december": 12, "jan": 1, "feb": 2, "mar": 3,
             "apr": 4, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}
# These words have a second meaning (ekim = sowing, aralık = range, kasım = a name, may = modal verb ...).
# They become a month only with date context.
AMBIG = {"nisan", "ekim", "kasim", "aralik", "aralig", "may", "march", "mar", "jan", "feb", "apr", "jun", "jul",
         "aug", "sep", "sept", "oct", "nov", "dec"}
TR_SUFFIX = re.compile(r"^'?(?:|[dt][ae](?:ki|n)?|n?[iu]n|[iu]|y?[ae]|n?[iu]n?[dt][ae](?:ki)?)$")
LOCATIVE = re.compile(r"^'?(?:[dt][ae](?:ki)?|n?[iu]n?[dt][ae](?:ki)?)$")
PART = {"basi": "early", "basinda": "early", "baslarinda": "early", "baslari": "early", "basindaki": "early",
        "early": "early", "beginning": "early", "start": "early",
        "ortasi": "mid", "ortasinda": "mid", "ortalarinda": "mid", "ortalari": "mid", "mid": "mid", "middle": "mid",
        "sonu": "late", "sonunda": "late", "sonlarinda": "late", "sonlari": "late", "sonundaki": "late",
        "late": "late", "end": "late"}
MONTH_WORD = {"ay", "ayi", "ayinda", "ayindaki", "aylarinda"}
PREV = {"gecen", "onceki", "last", "previous"}
THIS = {"bu", "this"}
LEAD = PREV | THIS | {"in", "during", "since", "of", "the", "on"}
NUM = {"bir": 1, "iki": 2, "uc": 3, "dort": 4, "bes": 5, "alti": 6, "yedi": 7, "sekiz": 8, "dokuz": 9, "on": 10,
       "a": 1, "an": 1, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8,
       "nine": 9, "ten": 10}
FEW = {"birkac": (2, 4), "several": (3, 6), "few": (2, 4), "couple": (2, 2)}
UNIT = {"gun": "d", "day": "d", "days": "d", "hafta": "w", "week": "w", "weeks": "w", "ay": "m", "month": "m",
        "months": "m", "yil": "y", "sene": "y", "year": "y", "years": "y"}
UNIT_TR = re.compile(r"^(gun|hafta|ay|yil|sene)(?:[dt][ae](?:ki)?|ki|n?[iu]n|[iu]|n?[iu]n[dt][ae]|ler|lar|lerde|larda)?$")
AGO = {"once", "onceden", "evvel", "ago", "oncesi"}
WEEKDAYS = {"pazartesi": 0, "sali": 1, "carsamba": 2, "persembe": 3, "cuma": 4, "cumartesi": 5, "pazar": 6,
            "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3, "friday": 4, "saturday": 5, "sunday": 6}
WEEKDAY_AMBIG = {"pazar", "sali"}  # market, (Tuesday / "salı" also a common word in forms)
SEASONS = {"yazin": 6, "summer": 6, "kisin": 12, "winter": 12, "ilkbahar": 3, "ilkbaharda": 3, "baharda": 3,
           "sonbahar": 9, "sonbaharda": 9, "autumn": 9}
SEASONS_CTX = {"yaz": 6, "kis": 12, "bahar": 3, "spring": 3, "fall": 9, "guz": 9}  # need a lead word
TOKEN = re.compile(r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[./]\d{1,2}[./]\d{2,4}|\d{1,2}[./]\d{4}|[^\W_]+(?:'[^\W_]+)?")
DAYNUM = re.compile(r"^(\d{1,2})(?:st|nd|rd|th)?$")


def _add_months(dt, n):
    y, m = divmod(dt.month - 1 + n, 12)
    return datetime(dt.year + y, m + 1, 1)


def _month(y, m):
    s = datetime(y, m, 1)
    return s, _add_months(s, 1)


def _part(s, e, part):
    if part == "early":
        return s, s + 10 * D
    if part == "mid":
        return s + 9 * D, s + 21 * D
    if part == "late":
        return s + 19 * D, e
    return s, e


def _month_of(tok):
    """'mart' / "mart'ta" / 'kasimda' / 'march' -> (month, stem, suffix) or None."""
    if tok in MONTHS_EN:
        return MONTHS_EN[tok], tok, ""
    for stem, m in MONTHS_TR.items():
        if tok.startswith(stem) and TR_SUFFIX.match(tok[len(stem):]):
            return m, stem, tok[len(stem):]
    return None


def _unit(tok):
    if tok in UNIT:
        return UNIT[tok]
    m = UNIT_TR.match(tok)
    return UNIT[m.group(1)] if m else None


def _ago(now, lo, hi, unit):
    size = {"d": D, "w": 7 * D, "m": 30.44 * D, "y": 365.25 * D}[unit]
    pad = {"d": D, "w": 3 * D, "m": 10 * D, "y": 60 * D}[unit]
    return now - hi * size - max(pad, 0.25 * hi * size), now - lo * size + max(pad, 0.25 * lo * size)


def _yearly(now, m1, d1, dy, m2, d2):
    """Most recent window [d1.m1.Y, d2.m2.(Y+dy)) that started on or before now."""
    y = now.year if datetime(now.year, m1, d1) <= now else now.year - 1
    return datetime(y, m1, d1), datetime(y + dy, m2, d2)


def _season(sm, now, year=None):
    if year is not None:
        s = datetime(year - (sm == 12), sm, 1)
    else:
        s = max(datetime(y, sm, 1) for y in range(now.year - 2, now.year + 1) if datetime(y, sm, 1) <= now)
    return s, _add_months(s, 3)


def parse(query, now):
    f = fold(query)
    toks = [(m.group(0), m.start(), m.end()) for m in TOKEN.finditer(f)]
    w = [t for t, _, _ in toks]
    n = len(w)
    used = [False] * n
    windows, keep = [], []
    today = datetime(now.year, now.month, now.day)

    def free(i):
        return 0 <= i < n and not used[i]

    def take(i, j, s, e, label):
        for k in range(i, j):
            used[k] = True
        windows.append((s, e, label))

    # Year context: "geçen yıl", "bu sene", "last year", or an explicit year 20xx.
    year_ctx, ctx_idx = None, []
    for i in range(n - 1):
        if w[i] in PREV | THIS and _unit(w[i + 1]) == "y":
            year_ctx, ctx_idx = now.year - (w[i] in PREV), [i, i + 1]
            break
    years = [(i, int(t[:4])) for i, t in enumerate(w) if re.match(r"^20\d\d(?:'?[^\W\d_]{0,8})?$", t)
             and int(t[:4]) <= now.year]
    if year_ctx is None and years:
        year_ctx, ctx_idx = years[0][1], [years[0][0]]

    def use_ctx():
        for k in ctx_idx:
            used[k] = True

    # 1) Numeric dates: 2025-03-14, 14.03.2025, 14/03/25, 03.2025
    for i, t in enumerate(w):
        try:
            if re.match(r"^\d{4}-\d{1,2}-\d{1,2}$", t):
                d = datetime(*map(int, t.split("-")))
            elif re.match(r"^\d{1,2}[./]\d{1,2}[./]\d{2,4}$", t):
                dd, mo, yy = map(int, re.split(r"[./]", t))
                d = datetime(yy + 2000 * (yy < 100), mo, dd)
            elif re.match(r"^\d{1,2}[./]\d{4}$", t):
                mo, yy = map(int, re.split(r"[./]", t))
                take(i, i + 1, *_month(yy, mo), f"{mo:02d}.{yy}")
                continue
            else:
                continue
        except ValueError:
            continue
        take(i, i + 1, d, d + D, f"{d:%d.%m.%Y}")

    # 2) Months: [early|end of] [day] month [day] [year] [ayı] [başı|ortası|sonu]
    for i, t in enumerate(w):
        mo = None if used[i] else _month_of(t)
        if not mo:
            continue
        month, stem, suffix = mo
        a, b = i, i + 1
        day = year = part = None
        month_word = False
        if free(i - 1) and DAYNUM.match(w[i - 1]) and 1 <= int(DAYNUM.match(w[i - 1]).group(1)) <= 31:
            day, a = int(DAYNUM.match(w[i - 1]).group(1)), i - 1
        elif b < n and not suffix and DAYNUM.match(w[b]) and 1 <= int(DAYNUM.match(w[b]).group(1)) <= 31:
            day, b = int(DAYNUM.match(w[b]).group(1)), b + 1
        if b < n and re.match(r"^20\d\d", w[b]):
            year, b = int(w[b][:4]), b + 1
        if b < n and w[b] in MONTH_WORD:
            month_word, b = True, b + 1
        if b < n and w[b] in PART:
            part, b = PART[w[b]], b + 1
        if part is None and a >= 2 and w[a - 1] == "of" and w[a - 2] in PART:
            part, a = PART[w[a - 2]], a - 2
            a -= a > 0 and w[a - 1] == "the"
        elif part is None and a >= 1 and w[a - 1] in ("early", "mid", "late"):
            part, a = PART[w[a - 1]], a - 1
        prev = a > 0 and w[a - 1] in PREV
        lead = a > 0 and w[a - 1] in LEAD
        if a > 0 and w[a - 1] in PREV | THIS:
            a -= 1
        context = bool(day or year or part or lead or month_word or LOCATIVE.match(suffix))
        original = query[toks[i][1]:toks[i][2]]
        if stem == "nisan" and ("ş" in original or "Ş" in original):
            continue  # "nişan" (engagement), not "nisan" (April)
        if (stem in AMBIG or re.match(r"^n?[iu]n$", suffix)) and not context:
            continue  # "Martin", "Kasım'ın", "aralığı", "may"
        if stem == "kasim":
            keep.append("kasim")  # can also be a person name
        if year is None:
            if year_ctx is not None:
                year = year_ctx
                use_ctx()
            elif prev and month >= now.month:
                year = now.year - 1
            else:
                year = now.year if month <= now.month else now.year - 1
        if day:
            try:
                d = datetime(year, month, day)
            except ValueError:
                continue
            take(a, b, d, d + D, f"{d:%d.%m.%Y}")
        else:
            s, e = _part(*_month(year, month), part)
            take(a, b, s, e, f"{month:02d}.{year}" + (f" {part}" if part else ""))

    # 3) "3 hafta önce", "a few weeks ago", "son 3 ay", "the past few months"
    for i in range(n):
        if used[i]:
            continue
        a = i
        if w[i] in FEW:
            lo, hi = FEW[w[i]]
            nx = i + 1 + (i + 1 < n and w[i + 1] == "of")
            a -= i > 0 and w[i - 1] == "a"
        elif w[i] == "bir" and i + 1 < n and w[i + 1] == "kac":
            lo, hi, nx = 2, 4, i + 2
        elif w[i].isdigit() or w[i] in NUM:
            lo = hi = int(w[i]) if w[i].isdigit() else NUM[w[i]]
            nx = i + 1
        else:
            continue
        unit = _unit(w[nx]) if nx < n else None
        if unit is None or hi > 400:
            continue
        if nx + 1 < n and w[nx + 1] in AGO:
            s, e = _ago(now, lo, hi, unit)
            take(a, nx + 2, s, e, f"~{lo}{'-' + str(hi) if hi != lo else ''}{unit} ago")
        elif a > 0 and w[a - 1] in ("son", "last", "past", "gecen"):
            a -= 1
            a -= a > 0 and w[a - 1] == "the"
            end = nx + 1 + (nx + 1 < n and w[nx + 1] in ("icinde", "icerisinde", "boyunca", "zarfinda"))
            size = {"d": D, "w": 7 * D, "m": 30.44 * D, "y": 365.25 * D}[unit]
            take(a, end, now - hi * size, now + D, f"last {hi}{unit}")

    # 4) "geçen hafta/ay/yıl [başı|ortası|sonu]", "bu ay", "last week", "this year"
    for i in range(n - 1):
        if not (free(i) and free(i + 1) and w[i] in PREV | THIS):
            continue
        unit = _unit(w[i + 1])
        if unit not in ("w", "m", "y"):
            continue
        prev = w[i] in PREV
        b, part = i + 2, None
        if b < n and w[b] in PART:
            part, b = PART[w[b]], b + 1
        if unit == "w":
            s = today - (today.weekday() + 7 * prev) * D
            e = s + 7 * D if prev else now + D
        elif unit == "m":
            s = _add_months(today, -prev)
            e = _add_months(s, 1) if prev else now + D
            s, e = _part(s, e, part)
        else:
            y = now.year - prev
            s, e = {"early": (datetime(y, 1, 1), datetime(y, 4, 1)), "mid": (datetime(y, 5, 1), datetime(y, 9, 1)),
                    "late": (datetime(y, 10, 1), datetime(y + 1, 1, 1))}.get(part, (datetime(y, 1, 1), datetime(y + 1, 1, 1)))
            e = min(e, now + D)
        take(i, b, s, e, " ".join(w[i:b]))

    # 5) Fixed words and phrases.
    xmas = lambda: _yearly(now, 12, 1, 1, 1, 1)       # December
    newyear = lambda: _yearly(now, 12, 20, 1, 1, 11)  # 20 Dec - 10 Jan
    january = lambda: _yearly(now, 1, 1, 0, 2, 1)
    recent = lambda: (now - 30 * D, now + D)
    long_ago = lambda: (datetime(2000, 1, 1), now - 180 * D)
    phrases = {
        "bugun": lambda: (today, now + D), "today": lambda: (today, now + D), "bu sabah": lambda: (today, now + D),
        "this morning": lambda: (today, now + D), "dun": lambda: (today - D, today), "yesterday": lambda: (today - D, today),
        "evvelsi gun": lambda: (today - 2 * D, today - D), "onceki gun": lambda: (today - 2 * D, today - D),
        "gecenlerde": recent, "gecen gunlerde": recent, "son zamanlarda": recent, "yakin zamanda": recent,
        "yakinlarda": recent, "recently": recent, "lately": recent, "the other day": lambda: (now - 10 * D, now + D),
        "son gunlerde": lambda: (now - 10 * D, now + D), "uzun zaman once": long_ago, "epey once": long_ago,
        "a long time ago": long_ago, "long ago": long_ago, "yillar once": lambda: (datetime(2000, 1, 1), now - 365 * D),
        "aylar once": lambda: (now - 400 * D, now - 90 * D), "months ago": lambda: (now - 400 * D, now - 90 * D),
        "yilbasindan once": xmas, "yilbasi oncesi": xmas, "yilbasi oncesinde": xmas, "yil sonu": xmas,
        "yil sonunda": xmas, "yilsonu": xmas, "yilsonunda": xmas, "before christmas": xmas, "before new year": xmas,
        "before the new year": xmas, "end of the year": xmas, "end of year": xmas, "year end": xmas,
        "yilbasinda": newyear, "yilbasi civarinda": newyear, "noel": newyear, "around christmas": newyear,
        "around new year": newyear, "christmas": newyear,
        "yilbasindan sonra": january, "after new year": january, "after the new year": january,
        "yil basinda": january, "beginning of the year": january, "start of the year": january,
    }
    for phrase in sorted(phrases, key=lambda p: -len(p.split())):
        pw = phrase.split()
        for i in range(n - len(pw) + 1):
            if w[i:i + len(pw)] == pw and all(free(i + k) for k in range(len(pw))):
                take(i, i + len(pw), *phrases[phrase](), phrase)

    # 6) Weekdays: "geçen cuma", "last friday", "salı günü", "on monday".
    for i, t in enumerate(w):
        if used[i] or t not in WEEKDAYS:
            continue
        a, b = i, i + 1
        if free(i - 1) and w[i - 1] in PREV | THIS | {"on"}:
            a = i - 1
        if b < n and w[b] in ("gunu", "gunu'ne"):
            b += 1
        if t in WEEKDAY_AMBIG and (a, b) == (i, i + 1):
            continue
        d = today - ((today.weekday() - WEEKDAYS[t]) % 7 or 7) * D
        take(a, b, d - D, d + 2 * D, t)

    # 7) Seasons: "yazın", "geçen kış", "in the fall", "2025 yazında".
    for i, t in enumerate(w):
        if used[i]:
            continue
        sm = SEASONS.get(t) or SEASONS_CTX.get(t)
        if sm is None and re.match(r"^(yaz|kis|bahar|ilkbahar|sonbahar)(?:in|da|inda)$", t):
            sm = {"yaz": 6, "kis": 12, "bahar": 3, "ilkbahar": 3, "sonbahar": 9}[re.match(r"^[a-z]+?(?=in|da|inda$)", t).group(0)]
        if sm is None:
            continue
        a, b = i, i + 1
        if free(i - 1) and w[i - 1] in LEAD:
            a = i - 1
            a -= a > 0 and w[a - 1] in ("in", "during") and w[i - 1] == "the"
        if t in SEASONS_CTX and a == i:
            continue
        year = None
        if b < n and re.match(r"^20\d\d", w[b]):
            year, b = int(w[b][:4]), b + 1
        elif year_ctx is not None:
            year = year_ctx
            use_ctx()
        take(a, b, *_season(sm, now, year), f"season {sm:02d}")

    # 8) Remaining year context.
    if ctx_idx and not any(used[k] for k in ctx_idx) and len(ctx_idx) == 2:
        take(ctx_idx[0], ctx_idx[1] + 1, datetime(year_ctx, 1, 1), min(datetime(year_ctx + 1, 1, 1), now + D),
             str(year_ctx))
    for i, y in years:
        if not used[i]:
            take(i, i + 1, datetime(y, 1, 1), min(datetime(y + 1, 1, 1), now + D), str(y))
        keep.append(str(y))

    spans = [(toks[i][1], toks[i][2]) for i in range(n) if used[i]]
    return windows, spans, keep


def remove_spans(text, spans):
    out, last = [], 0
    for a, b in sorted(spans):
        out += [text[last:a], " "]
        last = b
    out.append(text[last:])
    return re.sub(r"\s+", " ", "".join(out)).strip()


def date_score(dt, windows):
    """1.0 inside a window; exponential decay outside (scale: half the window, at least 3 days)."""
    best = 0.0
    for s, e, _ in windows:
        if s <= dt < e:
            return 1.0
        dist = (s - dt) if dt < s else (dt - e)
        best = max(best, 2.718281828 ** -(dist / max(3 * D, (e - s) / 2)))
    return best


TR_MONTHS = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


def tr_date(dt):
    return f"{dt.day} {TR_MONTHS[dt.month - 1]} {dt.year}"


def describe(windows):
    return [tr_date(s) if e - s <= D else f"{tr_date(s)} – {tr_date(e - D)}" for s, e, _ in windows]
