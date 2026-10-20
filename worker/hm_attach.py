"""Attachment text extraction (worker only, never inside Outlook).

extract_bytes(data, name, kind) -> (state, text). state: "ok", "empty" (no text found) or "unsupported".
An exception (corrupt or encrypted file) is the caller's business: it records "failed" and goes on.
All readers are local: no network, no external program. OCR uses Windows.Media.Ocr (Windows 10/11 only).

ponytail: plain readers for the common formats: zip + XML with the standard library (xlsx, docx, pptx,
OpenDocument), xlrd (xls), olefile (msg), pypdfium2 (pdf). Not read: rar/7z, legacy doc/ppt, audio/video, CAD.
Upgrade path: one more branch in extract_bytes.
"""
import io
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

MAX_FILE = 20 * 2**20        # bigger attachments are not read (checked with Attachment.Size before saving)
MAX_TEXT = 100_000           # chars kept per attachment
MAX_PDF_PAGES = 50
MAX_OCR_PAGES = 3            # scanned PDF pages sent to OCR
MAX_ZIP_DEPTH = 1            # a zip in a zip is not read
MAX_ZIP_MEMBER = 20 * 2**20  # uncompressed bytes read per zip member (zip bomb guard)
MAX_ZIP_TOTAL = 100 * 2**20  # uncompressed bytes read per zip attachment
BY_VALUE, EMBEDDED_ITEM = 1, 5  # Outlook OlAttachmentType: a file / an attached Outlook item (saved as .msg)

IMAGE = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}
TEXT = {".txt", ".csv", ".tsv", ".log", ".md", ".json"}
MARKUP = {".xml", ".html", ".htm"}
SHEET = {".xlsx", ".xlsm", ".xltx", ".xltm"}
WORD = {".docx", ".docm", ".dotx", ".dotm"}
SLIDES = {".pptx", ".pptm", ".ppsx"}
ODF = {".odt", ".ods", ".odp"}
OTHER = {".pdf", ".xls", ".xlt", ".msg", ".zip"}


class OcrUnavailable(Exception):
    pass


def ocr_available():
    return sys.platform == "win32"


def readable(name, kind=BY_VALUE):
    """True when extract_bytes can read this attachment (Outlook attachment type + file extension)."""
    if kind == EMBEDDED_ITEM:
        return True
    if kind != BY_VALUE:  # 4 = link to a file, 6 = OLE object
        return False
    ext = os.path.splitext(name or "")[1].lower()
    return ext in TEXT | MARKUP | SHEET | WORD | SLIDES | ODF | OTHER or (ext in IMAGE and ocr_available())


def extract_bytes(data, name, kind=BY_VALUE, depth=0):
    ext = ".msg" if kind == EMBEDDED_ITEM else os.path.splitext(name or "")[1].lower()
    if len(data) > MAX_FILE:
        return "unsupported", ""
    if ext in TEXT:
        text = decode(data)
    elif ext in MARKUP:
        text = markup_text(decode(data))
    elif ext in SHEET:
        text = xlsx_text(data)
    elif ext in WORD:
        text = ooxml_text(data, r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml", "p", "t")
    elif ext in SLIDES:
        text = ooxml_text(data, r"ppt/(slides/slide\d+|notesSlides/notesSlide\d+)\.xml", "p", "t")
    elif ext in ODF:
        text = ooxml_text(data, r"content\.xml", ("p", "h"), None)
    elif ext in (".xls", ".xlt"):
        text = xls_text(data)
    elif ext == ".pdf":
        text = pdf_text(data)
    elif ext == ".msg":
        text = msg_text(data)
    elif ext == ".zip" and depth < MAX_ZIP_DEPTH:
        text = zip_text(data, depth)
    elif ext in IMAGE and ocr_available():
        try:
            text = ocr_image(data)
        except OcrUnavailable:
            return "unsupported", ""
    else:
        return "unsupported", ""
    text = tidy(text)[:MAX_TEXT]
    return ("ok" if text else "empty"), text


# --- plain text and markup ---------------------------------------------------------------------

def decode(data):
    for bom, enc in ((b"\xef\xbb\xbf", "utf-8"), (b"\xff\xfe", "utf-16-le"), (b"\xfe\xff", "utf-16-be")):
        if data.startswith(bom):
            return data[len(bom):].decode(enc, "replace")
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("cp1254", "replace")  # Turkish Windows code page


_TAG = re.compile(r"<(script|style)\b.*?</\1\s*>|<[^>]+>", re.S | re.I)


def markup_text(s):
    import html
    return html.unescape(_TAG.sub(" ", s))


def tidy(s):
    s = s.replace("\x00", " ").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t\f\v]+", " ", s)
    s = re.sub(r" ?\n ?", "\n", s)
    return re.sub(r"\n{3,}", "\n\n", s).strip()


# --- zip-based formats: Office Open XML, OpenDocument --------------------------------------------

def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _natural(s):
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", s)]


class _Capped(io.RawIOBase):
    """Reads at most cap bytes, then reports end of file. A zip header can give a false member size."""

    def __init__(self, f, cap):
        self.f, self.left = f, cap

    def readable(self):
        return True

    def readinto(self, b):
        if self.left <= 0:
            return 0
        data = self.f.read(min(len(b), self.left))
        self.left -= len(data)
        b[:len(data)] = data
        return len(data)


def _member(z, info):
    """Member bytes, or None when bigger than MAX_ZIP_MEMBER."""
    with z.open(info) as f:
        data = f.read(MAX_ZIP_MEMBER + 1)
    return None if len(data) > MAX_ZIP_MEMBER else data


def _xml(z, name):
    data = _member(z, z.getinfo(name))
    return None if data is None else ET.fromstring(data)


def _iterparse(z, name, events=("end",)):
    """Streaming parse of one member (large sheets). Stops quietly at MAX_ZIP_MEMBER bytes or at broken XML."""
    with z.open(name) as f:
        try:
            yield from ET.iterparse(io.BufferedReader(_Capped(f, MAX_ZIP_MEMBER)), events=events)
        except ET.ParseError:
            return


def ooxml_text(data, pattern, para_tags, text_tag):
    """One line per paragraph of the members that match pattern. text_tag None: all text in the paragraph."""
    para_tags = (para_tags,) if isinstance(para_tags, str) else para_tags
    out = []
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for name in sorted((n for n in z.namelist() if re.fullmatch(pattern, n)), key=_natural):
            root = _xml(z, name)
            if root is None:
                continue
            for el in root.iter():
                if _local(el.tag) in para_tags:
                    t = "".join(el.itertext()) if text_tag is None else "".join(
                        (x.text or "") if _local(x.tag) == text_tag else " " if _local(x.tag) in ("tab", "br", "cr") else ""
                        for x in el.iter())
                    if t.strip():
                        out.append(t)
    return "\n".join(out)


def xlsx_text(data):
    """[sheet name], then one line per row with tab-separated cell values. Stops at MAX_TEXT."""
    out, size = [], 0
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = set(z.namelist())
        shared = _shared_strings(z, names)
        for member, title in _sheets(z, names):
            out.append("[" + title + "]")
            for cells in _rows(z, member, shared):
                out.append("\t".join(cells))
                size += len(out[-1]) + 1
                if size > MAX_TEXT:
                    return "\n".join(out)
    return "\n".join(out)


def _shared_strings(z, names):
    out = []
    if "xl/sharedStrings.xml" in names:
        root = None
        for ev, el in _iterparse(z, "xl/sharedStrings.xml", ("start", "end")):
            if root is None:
                root = el
            elif ev == "end" and _local(el.tag) == "si":
                out.append("".join(x.text or "" for x in el.iter() if _local(x.tag) == "t"))
                root.clear()  # keep memory flat
    return out


def _sheets(z, names):
    """[(member, sheet name)] in workbook tab order."""
    out = []
    wb = _xml(z, "xl/workbook.xml") if "xl/workbook.xml" in names else None
    rels = _xml(z, "xl/_rels/workbook.xml.rels") if "xl/_rels/workbook.xml.rels" in names else None
    if wb is not None and rels is not None:
        target = {r.get("Id"): r.get("Target") or "" for r in rels}
        for s in wb.iter():
            if _local(s.tag) == "sheet":
                rid = next((v for k, v in s.attrib.items() if _local(k) == "id"), None)
                t = target.get(rid, "")
                member = t.lstrip("/") if t.startswith("/") else "xl/" + t
                if member in names:
                    out.append((member, s.get("name") or member))
    return out or [(n, n.rsplit("/", 1)[-1][:-4]) for n in sorted(names, key=_natural)
                   if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]


def _rows(z, member, shared):
    cells, sheet_data = [], None
    for ev, el in _iterparse(z, member, ("start", "end")):
        tag = _local(el.tag)
        if ev == "start":
            if tag == "sheetData":
                sheet_data = el
        elif tag == "c":
            t = el.get("t")
            if t == "inlineStr":
                v = "".join(x.text or "" for x in el.iter() if _local(x.tag) == "t")
            else:
                v = next((x.text for x in el if _local(x.tag) == "v"), None) or ""
                if t == "s":
                    v = shared[int(v)] if v.isdigit() and int(v) < len(shared) else ""
            if v.strip():
                cells.append(v)
        elif tag == "row":
            if cells:
                yield cells
            cells = []
            if sheet_data is not None:
                sheet_data.clear()  # keep memory flat
    if cells:
        yield cells


def xls_text(data):
    import xlrd
    book = xlrd.open_workbook(file_contents=data, on_demand=True, logfile=io.StringIO())  # no messages on stdout
    out, size = [], 0
    try:
        for sh in book.sheets():
            out.append("[" + sh.name + "]")
            for r in range(sh.nrows):
                vals = [str(v) for v in sh.row_values(r) if v not in ("", None)]
                if vals:
                    out.append("\t".join(vals))
                    size += len(out[-1]) + 1
                    if size > MAX_TEXT:
                        return "\n".join(out)
    finally:
        book.release_resources()
    return "\n".join(out)


# --- PDF ---------------------------------------------------------------------------------------

def pdf_text(data):
    """Text layer of the first MAX_PDF_PAGES pages. A page without text (a scan) goes to OCR, at most
    MAX_OCR_PAGES pages per file."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(data)
    out, ocr_pages = [], 0
    try:
        for i in range(min(len(doc), MAX_PDF_PAGES)):
            page = doc[i]
            try:
                tp = page.get_textpage()
                try:
                    t = tp.get_text_range()
                finally:
                    tp.close()
                if len(t.strip()) < 20 and ocr_available() and ocr_pages < MAX_OCR_PAGES:
                    ocr_pages += 1
                    # 144 DPI (OCR needs letters of about 10 px), at most 4000 px: a poster-size page stays small.
                    bmp = page.render(scale=min(2.0, 4000 / max(1.0, *page.get_size())))
                    try:
                        t = ocr_image(_bmp_file(bmp))
                    except Exception:  # OCR is best effort: keep the other pages
                        pass
                    finally:
                        bmp.close()
                out.append(t)
            finally:
                page.close()
    finally:
        doc.close()
    return "\n".join(out)


def _bmp_file(bmp):
    """A pypdfium2 bitmap (BGR or BGRx, the BMP byte order) as a BMP file. No Pillow needed."""
    import struct
    w, h, stride, n_ch = bmp.width, bmp.height, bmp.stride, bmp.n_channels
    if n_ch not in (3, 4):
        raise ValueError("unexpected bitmap format")
    buf = bytes(bmp.buffer)
    row = ((w * n_ch + 3) // 4) * 4
    pixels = b"".join(buf[y * stride:y * stride + w * n_ch].ljust(row, b"\0") for y in range(h - 1, -1, -1))
    header = struct.pack("<2sIHHI", b"BM", 54 + len(pixels), 0, 0, 54)
    info = struct.pack("<IiiHHIIiiII", 40, w, h, 1, n_ch * 8, 0, len(pixels), 2835, 2835, 0, 0)
    return header + info + pixels


# --- Outlook .msg (attached mail) --------------------------------------------------------------

def msg_text(data):
    """Subject, sender, recipients and body of an attached Outlook item (MAPI properties)."""
    import olefile
    ole = olefile.OleFileIO(data)
    try:
        def prop(code):
            for typ, enc in (("001F", "utf-16-le"), ("001E", "cp1254")):
                name = "__substg1.0_" + code + typ
                if ole.exists(name):
                    return ole.openstream(name).read().decode(enc, "replace").rstrip("\x00")
            return ""
        parts = [prop("0037"), prop("0C1A"), prop("0E04"), prop("1000")]  # subject, sender, to, body
        if not parts[3] and ole.exists("__substg1.0_10130102"):  # no plain text body: HTML body
            parts[3] = markup_text(decode(ole.openstream("__substg1.0_10130102").read()))
        return "\n".join(x for x in parts if x)
    finally:
        ole.close()


# --- zip ---------------------------------------------------------------------------------------

def zip_text(data, depth):
    out, size, read = [], 0, 0
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for info in z.infolist()[:200]:
            if info.is_dir() or info.flag_bits & 0x1 or not readable(info.filename):  # 0x1: encrypted
                continue
            if read >= MAX_ZIP_TOTAL:
                break
            try:
                member = _member(z, info)
                read += len(member or b"")
                st, t = ("unsupported", "") if member is None else extract_bytes(member, info.filename, BY_VALUE, depth + 1)
            except Exception:  # one bad member does not stop the others
                continue
            if st == "ok":
                out.append("[" + info.filename + "]\n" + t)
                size += len(out[-1])
                if size > MAX_TEXT:
                    break
    return "\n".join(out)


# --- OCR: Windows.Media.Ocr (WinRT), the OCR engine of Windows itself -----------------------------

_engine = None


def _load_vc_runtime_first():
    """winrt-runtime 3.2.1 ships its own msvcp140.dll (14.29). When winrt loads first, that old copy stays in the
    process and a newer native library then fails: torch's c10.dll gave WinError 1114 on the Windows test machine;
    onnxruntime.dll needs the newer copy too. Load the System32 copy first: winrt then uses the msvcp140.dll that is
    already loaded (Windows matches a loaded DLL by its name)."""
    import ctypes
    try:
        ctypes.WinDLL(os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "msvcp140.dll"))
    except OSError:  # no VC++ runtime in System32: winrt uses its own copy
        pass


def ocr_image(data):
    """Text in an image file (PNG, JPEG, BMP, GIF, TIFF). Uses the OCR languages of the Windows user profile.
    Each WinRT call is waited for with .get() on this thread: no asyncio event loop (that opens a local socket
    pair on Windows), and the job object of the add-in allows no extra process."""
    global _engine
    if _engine is None:
        _load_vc_runtime_first()
    from winrt.windows.graphics.imaging import (BitmapAlphaMode, BitmapDecoder, BitmapInterpolationMode,
                                                BitmapPixelFormat, BitmapTransform, ColorManagementMode,
                                                ExifOrientationMode)
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    if _engine is None:
        _engine = OcrEngine.try_create_from_user_profile_languages()
        if _engine is None:  # no OCR language installed for the user's languages
            raise OcrUnavailable()
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(data)
    writer.store_async().get()
    writer.detach_stream()
    stream.seek(0)
    dec = BitmapDecoder.create_async(stream).get()
    tf = BitmapTransform()
    w, h, limit = dec.pixel_width, dec.pixel_height, OcrEngine.max_image_dimension
    if max(w, h) > limit:  # the engine refuses bigger images
        k = limit / max(w, h)
        tf.scaled_width, tf.scaled_height = max(1, int(w * k)), max(1, int(h * k))
        tf.interpolation_mode = BitmapInterpolationMode.FANT
    # pywinrt gives each WinRT overload its own name: GetSoftwareBitmapAsync(format, alpha, transform, exif, color)
    bmp = dec.get_software_bitmap_transformed_async(BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED, tf,
                                                    ExifOrientationMode.RESPECT_EXIF_ORIENTATION,
                                                    ColorManagementMode.DO_NOT_COLOR_MANAGE).get()
    res = _engine.recognize_async(bmp).get()
    return "\n".join(line.text for line in res.lines)
