"""Extract text from a PDF to Markdown, Vietnamese-aware (legacy fonts, OCR, Office export bugs).

Usage: py fetch_pdf.py <pdf-path-or-url> <outdir> [--pages 1-20,35] [--ocr | --ocr-all] [--ocr-lang vie]

  --ocr       OCR the pages that have no text layer or whose text is still garbled
  --ocr-all   OCR every page (when the PDF's own text layer is junk)
  --ocr-lang  Tesseract languages, default "vie" (it reads English fine; "vie+eng" makes
              the English model strip Vietnamese tone marks). ~1 s/page.
  OCR uses PyMuPDF's built-in Tesseract; it only needs <lang>.traineddata (tessdata_fast)
  in $TESSDATA_PREFIX, %LOCALAPPDATA%\\tessdata (Windows) or ~/.local/share/tessdata (Mac/Linux),
  or a system Tesseract install. OCR pages are marked `<!-- p. N (OCR) -->`.

Writes to <outdir>:
  source.pdf    - only when the input is a URL (the downloaded copy)
  meta.json     - metadata, table of contents, fonts with their detected encoding
  text.md       - extracted text, one `<!-- p. N -->` marker per page
  text.raw.md   - the same text before any fix (only when a fix changed something)
  summary.json  - the JSON summary printed to stdout

Vietnamese fixes:
  - Legacy 8-bit Vietnamese fonts are decoded span by span:
      tcvn3  .VnTime, VnArial, VnAvantH ... (a name ending in H is an all-caps font)
      vni    VNI-Times, VNI-Helve ...
      vn2    VnPalatino2, VnFujiyama2 ... (base letter + separate tone character)
    The decoder for each font is the one that turns most of its words into valid
    Vietnamese syllables (so renamed/substituted fonts are caught too); the font name only
    breaks ties. Spans that already carry Unicode Vietnamese letters are left alone.
  - U+01A3/U+01A2 (Latin letter "oi") -> u-horn. Office exports of Arial/Times map the
    u-horn glyph to this code point; it never occurs in real Vietnamese text.
  - Eth (Ð/ð) typed for d-stroke, on lines that are otherwise Vietnamese.
  - UTF-8 mojibake ("Viá»‡t" -> "Việt", "Ã©" -> "é").
  - Lost spaces after a toned syllable ("dạdày" -> "dạ dày"), on Vietnamese lines only.
    Common in OCR output and in some Word exports.
  - NFC normalisation.

Prints a JSON summary to stdout. Two lists need a follow-up by the caller:
  image_only_pages - no text layer (scans, slides of pictures): rerun with --ocr
  suspect_pages    - text still looks like legacy-font garble: rerun with --ocr, or check
                     visually before quoting
"""
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path

import fitz  # PyMuPDF

ACUTE, GRAVE, HOOK, TILDE, DOT = "́", "̀", "̉", "̃", "̣"
CIRC, BREVE = "̂", "̆"

# --- TCVN3 (ABC): one precomposed letter per byte ---------------------------------------
TCVN3 = {
    "\xa8": "ă", "\xa9": "â", "\xaa": "ê", "\xab": "ô", "\xac": "ơ", "\xad": "ư", "\xae": "đ",
    "\xa1": "Ă", "\xa2": "Â", "\xa3": "Ê", "\xa4": "Ô", "\xa5": "Ơ", "\xa6": "Ư", "\xa7": "Đ",
}
_TCVN3_TONED = {  # Unicode letters (grave, hook, tilde, acute, dot) -> their TCVN3 bytes
    "àảãáạ": "\xb5\xb6\xb7\xb8\xb9",
    "ằẳẵắặ": "\xbb\xbc\xbd\xbe\xc6",
    "ầẩẫấậ": "\xc7\xc8\xc9\xca\xcb",
    "èẻẽéẹ": "\xcc\xce\xcf\xd0\xd1",
    "ềểễếệ": "\xd2\xd3\xd4\xd5\xd6",
    "ìỉĩíị": "\xd7\xd8\xdc\xdd\xde",
    "òỏõóọ": "\xdf\xe1\xe2\xe3\xe4",
    "ồổỗốộ": "\xe5\xe6\xe7\xe8\xe9",
    "ờởỡớợ": "\xea\xeb\xec\xed\xee",
    "ùủũúụ": "\xef\xf1\xf2\xf3\xf4",
    "ừửữứự": "\xf5\xf6\xf7\xf8\xf9",
    "ỳỷỹýỵ": "\xfa\xfb\xfc\xfd\xfe",
}
for _uni, _codes in _TCVN3_TONED.items():
    TCVN3.update(zip(_codes, unicodedata.normalize("NFC", _uni)))

# --- VNI Windows: base vowel followed by a modifier byte --------------------------------
VNI = {
    "ô": "ơ", "Ô": "Ơ", "ö": "ư", "Ö": "Ư", "ñ": "đ", "Ñ": "Đ",
    "í": "í", "ì": "ì", "æ": "ỉ", "ó": "ĩ", "ò": "ị",
    "Í": "Í", "Ì": "Ì", "Æ": "Ỉ", "Ó": "Ĩ", "Ò": "Ị",
    "ù": ACUTE, "ø": GRAVE, "û": HOOK, "õ": TILDE, "ï": DOT,
    "Ù": ACUTE, "Ø": GRAVE, "Û": HOOK, "Õ": TILDE, "Ï": DOT,
    "â": CIRC, "á": CIRC + ACUTE, "à": CIRC + GRAVE, "å": CIRC + HOOK, "ã": CIRC + TILDE, "ä": CIRC + DOT,
    "Â": CIRC, "Á": CIRC + ACUTE, "À": CIRC + GRAVE, "Å": CIRC + HOOK, "Ã": CIRC + TILDE, "Ä": CIRC + DOT,
    "ê": BREVE, "é": BREVE + ACUTE, "è": BREVE + GRAVE, "ú": BREVE + HOOK, "ü": BREVE + TILDE, "ë": BREVE + DOT,
    "Ê": BREVE, "É": BREVE + ACUTE, "È": BREVE + GRAVE, "Ú": BREVE + HOOK, "Ü": BREVE + TILDE, "Ë": BREVE + DOT,
}

# --- Vn*2 fonts: shaped base letter + separate tone character ---------------------------
# Derived from VnPalatino2 text ("ngûúâi Viïåt" = "người Việt"). Two tone sets: one placed
# after plain vowels, a higher one after â ê ô ă; both decode to the same combining mark.
VN2 = {
    "ê": "â", "ï": "ê", "ö": "ô", "ù": "ă", "ú": "ơ", "û": "ư", "à": "đ",
    "Ê": "Â", "Ï": "Ê", "Ö": "Ô", "Ù": "Ă", "Ú": "Ơ", "Û": "Ư", "À": "Đ",
    "ñ": "í", "ò": "ì", "ó": "ỉ", "ô": "ĩ", "õ": "ị",
    "Ñ": "Í", "Ò": "Ì", "Ó": "Ỉ", "Ô": "Ĩ", "Õ": "Ị",
    "á": ACUTE, "â": GRAVE, "ã": HOOK, "ä": TILDE, "å": DOT,  # after plain vowels
    "æ": ACUTE, "ç": GRAVE, "è": HOOK, "é": TILDE,              # after ă ("răçng" = "rằng")
    "ë": ACUTE, "ì": GRAVE, "í": HOOK, "î": TILDE,              # after â ê ô
    "Á": ACUTE, "Â": GRAVE, "Ã": HOOK, "Ä": TILDE, "Å": DOT,
    "Æ": ACUTE, "Ç": GRAVE, "È": HOOK, "É": TILDE,
    "Ë": ACUTE, "Ì": GRAVE, "Í": HOOK, "Î": TILDE,
}

TABLES = {"tcvn3": str.maketrans(TCVN3), "vni": str.maketrans(VNI), "vn2": str.maketrans(VN2)}

# Letters that only exist in Unicode Vietnamese, never in raw legacy-font output.
UNICODE_VI = re.compile(r"[ăđơưĂĐƠƯẠ-ỹ]")

# --- Vietnamese syllable check, used to pick the right decoder for each font -----------
TONE_MARKS = {ACUTE, GRAVE, HOOK, TILDE, DOT}
SHAPE_MARKS = {CIRC, BREVE, "̛"}  # circumflex, breve, horn
NUCLEI = set(
    "a ă â e ê i o ô ơ u ư y ai ao au ay âu ây eo êu ia iê iu oa oă oe oi ôi ơi oo ua uâ uê ui "
    "uô uơ uy ưa ưi ưu ươ yê iêu yêu oai oay oeo uây uôi ươi ươu uya uyê uyu".split()
)
NEEDS_FINAL = {"iê", "yê", "uô", "ươ", "uyê"}  # "tương", never "tươ"
SYLLABLE = re.compile(
    r"(ngh|ng|nh|ch|gh|gi|kh|ph|qu|th|tr|[bcdđghklmnprstvx])?([aăâeêioôơuưy]{1,3})(ch|ng|nh|[cmnpt])?"
)
EDGE_PUNCT = "\"'“”‘’()[]{}.,;:!?…-–—/*"


def is_vi_syllable(word: str) -> bool:
    w = unicodedata.normalize("NFD", word.lower())
    tones, base = 0, ""
    for c in w:
        if c in TONE_MARKS:
            tones += 1
            if base not in "aeiouy":  # tone on a consonant: ñ, ṭ ... not Vietnamese
                return False
        elif not unicodedata.combining(c):
            base = c
        elif c not in SHAPE_MARKS:
            return False
    if tones > 1:
        return False
    w = "".join(c for c in w if c not in TONE_MARKS)
    m = SYLLABLE.fullmatch(unicodedata.normalize("NFC", w))
    if not m or m.group(2) not in NUCLEI:
        return False
    return bool(m.group(3)) or m.group(2) not in NEEDS_FINAL


def _starts_with_vowel(part: str) -> bool:
    return unicodedata.normalize("NFD", part[0].lower())[0] in "aeiouy"


def split_glued(tok: str) -> list[str]:
    """'dạdày' -> ['dạ', 'dày']: OCR and some Word exports drop the space after a syllable.

    Only for tokens that are not a valid syllable themselves, and only after a part that
    carries a diacritic: that is where the space goes missing, and it keeps foreign words
    ('café', 'René') whole. Picks the split with the fewest parts, then the fewest non-
    Vietnamese parts, then the fewest parts starting with a vowel ('kỹ năng', not 'kỹn ăng').
    A trailing ASCII run of 3+ letters may be one part ('nghềtrading' -> 'nghề trading').
    """
    if tok.isascii() or len(tok) > 40 or is_vi_syllable(tok):
        return [tok]
    n = len(tok)
    best: list = [None] * (n + 1)  # best[i] = (parts, foreign parts, vowel-initial, [parts])
    best[n] = (0, 0, 0, [])
    for i in range(n - 1, -1, -1):
        for j in range(i + 1, n + 1):
            part = tok[i:j]
            if best[j] is None or (j < n and part.isascii()):
                continue
            vi = j - i <= 8 and is_vi_syllable(part)
            ascii_tail = j == n and part.isascii() and len(part) >= 3 and not tok[i - 1].isascii()
            if not (vi or ascii_tail):
                continue
            nxt = best[j]
            cand = (nxt[0] + 1, nxt[1] + (not vi), nxt[2] + _starts_with_vowel(part), [part] + nxt[3])
            if best[i] is None or cand[:3] < best[i][:3]:
                best[i] = cand
    return best[0][3] if best[0] and len(best[0][3]) > 1 else [tok]


WORD_RUN = re.compile(r"[^\W\d_]+")


def unglue(line: str) -> tuple[str, int]:
    """Re-insert lost spaces between Vietnamese syllables on a Vietnamese line."""
    if not UNICODE_VI.search(line):
        return line, 0
    count = 0

    def fix(m: re.Match) -> str:
        nonlocal count
        parts = split_glued(m.group())
        count += len(parts) - 1
        return " ".join(parts)

    return WORD_RUN.sub(fix, line), count


def vi_score(words: list[str], table) -> float:
    """Share of words that become valid Vietnamese syllables after decoding with `table`."""
    ok = sum(is_vi_syllable(w.translate(table) if table else w) for w in words)
    return ok / len(words)


def pick_encoding(words: list[str], name_hint: str | None) -> str | None:
    """Choose the decoder (or None) for one font from a sample of its non-ASCII words."""
    if len(words) < 3:  # too little evidence: trust the font name
        return name_hint
    scores = {enc: vi_score(words, TABLES[enc]) for enc in TABLES}
    plain = vi_score(words, None)
    best = max(scores, key=scores.get)
    if name_hint and scores[name_hint] >= scores[best] - 0.05:
        best = name_hint
    return best if scores[best] >= 0.5 and scores[best] - plain >= 0.15 else None
# Characters that, when dense, mean undecoded legacy-font output. Common Western punctuation
# (« » © °) is left out; French î ç ë stay in but are far below the density threshold.
SUSPECT = set("äåæçëîïðñöøûüþÿÄÅÆÇËÎÏÐÑÖØÛÜÞßƣƢ¨ª¬®¡¢£¤¥¦§µ¶·¸¹¼½¾\xad")
OI_FIX = str.maketrans({"ƣ": "ư", "Ƣ": "Ư"})
ETH_FIX = str.maketrans({"Ð": "Đ", "ð": "đ"})  # Icelandic eth typed for d-stroke in old texts
# UTF-8 bytes shown as cp1252: a lead byte (Ã, Â, á...) then continuation bytes (0x80-0xBF).
_CONT = ("[\u0080-¿ŒœŠšŸŽžƒˆ˜–—"
         "‘-‚“-„†-•…‰‹›€™]")
MOJIBAKE = re.compile(f"[à-ï]{_CONT}{_CONT}|[Â-ß]{_CONT}")


def font_encoding(fontname: str) -> tuple[str | None, bool]:
    """Return (legacy encoding or None, all-caps font?) from a PDF font name."""
    name = fontname.split("+", 1)[-1]  # drop subset tag "ABCDEE+"
    if re.match(r"VNI[-_ ]", name, re.I):
        return "vni", False
    if re.match(r"\.?Vn\w*?2(?:[-,]|$)", name):
        return "vn2", False
    if re.match(r"\.?Vn[A-Z]", name):
        family = re.split(r"[,-]", name)[0]
        return "tcvn3", bool(re.search(r"H(?:Normal|Bold|Italic)*$", family))
    return None, False


def parse_pages(spec: str | None, count: int) -> list[int]:
    if not spec:
        return list(range(count))
    pages: set[int] = set()
    for part in spec.split(","):
        a, _, b = part.strip().partition("-")
        lo, hi = int(a), int(b or a)
        pages.update(range(max(lo, 1) - 1, min(hi, count)))
    return sorted(pages)


def ranges(nums: list[int]) -> str:
    """[1,2,3,7] -> '1-3,7' (1-based page numbers)."""
    out, start, prev = [], None, None
    for n in nums:
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            out.append(f"{start}-{prev}" if prev != start else f"{start}")
            start = prev = n
    if start is not None:
        out.append(f"{start}-{prev}" if prev != start else f"{start}")
    return ",".join(out)


def download(url: str) -> bytes:
    """urllib with certifi's CA bundle; fall back to curl.exe (Windows cert store)."""
    try:
        import ssl
        try:
            import certifi
            ctx = ssl.create_default_context(cafile=certifi.where())
        except ImportError:
            ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
            return resp.read()
    except Exception:
        out = subprocess.run(["curl", "-L", "-sS", "-f", "-A", "Mozilla/5.0", url],
                             capture_output=True, timeout=300)
        if out.returncode != 0:
            raise RuntimeError(f"download failed: {out.stderr.decode(errors='replace')[-300:]}")
        return out.stdout


def _demojibake(m: re.Match) -> str:
    try:  # bytes 0x81 0x8D 0x8F 0x90 0x9D are undefined in cp1252 and arrive as C1 controls
        raw = b"".join(c.encode("latin-1") if "\x80" <= c <= "\x9f" else c.encode("cp1252")
                       for c in m.group())
        return raw.decode("utf-8")
    except UnicodeError:
        return m.group()


def fix_mojibake(line: str) -> str:
    """'Ã©pique' / 'Viá»‡t' (UTF-8 bytes shown as cp1252) -> 'épique' / 'Việt'."""
    return MOJIBAKE.sub(_demojibake, line)


def load_source(src: str, outdir: Path) -> Path:
    if not re.match(r"https?://", src, re.I):
        return Path(src)
    data = download(src)
    if b"%PDF" not in data[:1024]:
        raise ValueError("URL did not return a PDF (no %PDF header) - maybe an HTML landing page")
    path = outdir / "source.pdf"
    path.write_bytes(data)
    return path


DEFAULT_TESSDATA = (Path(os.environ["LOCALAPPDATA"]) / "tessdata" if os.environ.get("LOCALAPPDATA")
                    else Path.home() / ".local" / "share" / "tessdata")


def tessdata_dir(langs: list[str]) -> tuple[str | None, str | None]:
    """(folder holding <lang>.traineddata for every lang, or None; error message or None)."""
    candidates = [os.environ.get("TESSDATA_PREFIX"), str(DEFAULT_TESSDATA)]
    try:
        candidates.append(fitz.get_tessdata())
    except Exception:
        pass
    for cand in filter(None, candidates):
        if all(Path(cand, f"{lang}.traineddata").exists() for lang in langs):
            return cand, None
    need = ", ".join(f"{lang}.traineddata" for lang in langs)
    return None, (f"OCR data missing: put {need} (from github.com/tesseract-ocr/tessdata_fast) "
                  f"in {candidates[1]} or set TESSDATA_PREFIX")


def ocr_page(page: "fitz.Page", lang: str, tessdata: str) -> tuple[str, int]:
    """OCR one page -> (text, spaces restored). Tesseract's vie model often drops the space
    after a syllable with a tone mark ('dạdày'), so the text goes through unglue()."""
    # Keep `page` referenced while the OCR textpage is in use (PyMuPDF holds it weakly).
    tp = page.get_textpage_ocr(language=lang, dpi=300, full=True, tessdata=tessdata)
    lines, restored = [], 0
    for ln in page.get_text(textpage=tp).splitlines():
        ln, n = unglue(unicodedata.normalize("NFC", ln).rstrip())
        lines.append(ln)
        restored += n
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip(), restored


def take_option(args: list[str], name: str, has_value: bool):
    if name not in args:
        return None
    i = args.index(name)
    value = args[i + 1] if has_value else True
    del args[i:i + (2 if has_value else 1)]
    return value


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    args = sys.argv[1:]
    page_spec = take_option(args, "--pages", True)
    ocr_lang = take_option(args, "--ocr-lang", True) or "vie"
    ocr_all = bool(take_option(args, "--ocr-all", False))
    ocr = ocr_all or bool(take_option(args, "--ocr", False))
    if len(args) < 2:
        print(json.dumps({"ok": False, "error": "usage: fetch_pdf.py <pdf-path-or-url> <outdir> "
                          "[--pages 1-20,35] [--ocr | --ocr-all] [--ocr-lang vie]"}))
        sys.exit(1)
    src, outdir = args[0], Path(args[1])
    outdir.mkdir(parents=True, exist_ok=True)

    try:
        pdf_path = load_source(src, outdir)
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(json.dumps({"ok": False, "error": f"{e.__class__.__name__}: {e}"[:500]}, ensure_ascii=False))
        sys.exit(1)
    if doc.needs_pass:
        print(json.dumps({"ok": False, "error": "PDF is password-protected"}))
        sys.exit(1)

    pages = parse_pages(page_spec, doc.page_count)

    # Pass 1: collect spans per page and per-font statistics.
    extracted = []  # (page_no, has_images, [block[line[(font, text)]]])
    font_words: dict[str, list[str]] = defaultdict(list)  # sample of non-ASCII words per font
    for pno in pages:
        page = doc[pno]
        blocks = []
        for b in page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]:
            lines = []
            for line in b.get("lines", []):
                spans = [(s["font"], s["text"]) for s in line["spans"] if s["text"]]
                for font, text in spans:
                    sample = font_words[font]
                    if len(sample) < 400 and not UNICODE_VI.search(text):
                        for tok in text.split():
                            tok = tok.strip(EDGE_PUNCT + "0123456789")
                            if tok and any(ord(c) > 127 for c in tok):
                                sample.append(tok)
                if spans:
                    lines.append(spans)
            if lines:
                blocks.append(lines)
        extracted.append((pno + 1, bool(page.get_images()), blocks))

    fonts = {}
    for font in sorted({f for _, _, bl in extracted for ln in bl for sp in ln for f, _ in sp}):
        hint, caps = font_encoding(font)
        enc = pick_encoding(font_words[font], hint)
        note = None
        if hint and enc != hint:
            note = f"name suggests {hint}, text says {enc or 'Unicode'}"
        elif enc and not hint:
            note = "detected from text (font name gives no hint)"
        fonts[font] = {"encoding": enc, "all_caps": caps, "note": note}

    # Pass 2: decode, fix, assemble.
    fixes = Counter()
    raw_parts, fixed_parts = [], []
    image_only, suspect_pages, samples = [], [], []
    for page_no, has_images, blocks in extracted:
        raw_blocks, fixed_blocks = [], []
        for lines in blocks:
            raw_lines, fixed_lines = [], []
            for spans in lines:
                raw_line, fixed_line = "", ""
                for font, text in spans:
                    raw_line += text
                    info = fonts[font]
                    out = text
                    if info["encoding"] and not UNICODE_VI.search(text):
                        decoded = text.translate(TABLES[info["encoding"]])
                        if info["all_caps"]:
                            decoded = decoded.upper()
                        if decoded != text:
                            fixes[info["encoding"]] += 1
                        out = decoded
                    n_oi = out.count("ƣ") + out.count("Ƣ")
                    if n_oi:
                        fixes["oi_to_u_horn"] += n_oi
                        out = out.translate(OI_FIX)
                    fixed_line += out
                raw_line = raw_line.rstrip()
                if MOJIBAKE.search(fixed_line):
                    repaired = fix_mojibake(fixed_line)
                    if repaired != fixed_line:
                        fixes["utf8_mojibake"] += 1
                        fixed_line = repaired
                fixed_line = unicodedata.normalize("NFC", fixed_line).rstrip()
                if ("Ð" in fixed_line or "ð" in fixed_line) and UNICODE_VI.search(fixed_line):
                    fixes["eth_to_d_stroke"] += fixed_line.count("Ð") + fixed_line.count("ð")
                    fixed_line = fixed_line.translate(ETH_FIX)
                fixed_line, n_spaces = unglue(fixed_line)
                fixes["spaces_restored"] += n_spaces
                if fixed_line.strip():
                    raw_lines.append(raw_line)
                    fixed_lines.append(fixed_line)
            if fixed_lines:
                raw_blocks.append("\n".join(raw_lines))
                fixed_blocks.append("\n".join(fixed_lines))
        page_text = "\n\n".join(fixed_blocks)

        if len(page_text.strip()) < 25 and has_images:
            image_only.append(page_no)
        letters = sum(c.isalpha() for c in page_text)
        bad = sum(c in SUSPECT for c in page_text)
        if bad >= 10 and bad > 0.03 * max(letters, 1):
            suspect_pages.append(page_no)
            if len(samples) < 3:
                line = next(ln for ln in page_text.splitlines() if any(c in SUSPECT for c in ln))
                samples.append(f"p.{page_no}: {line[:120]}")

        raw_parts.append((page_no, "\n\n".join(raw_blocks)))
        fixed_parts.append((page_no, page_text))

    # OCR: pages without a text layer and pages whose text is still garbled (or every page).
    ocr_done, ocr_empty, ocr_error, ocr_seconds = [], [], None, 0.0
    if ocr:
        targets = {p for p, _ in fixed_parts} if ocr_all else set(image_only + suspect_pages)
        tessdata, ocr_error = tessdata_dir(ocr_lang.split("+"))
        if targets and tessdata:
            t0 = time.time()
            for i, (page_no, _) in enumerate(fixed_parts):
                if page_no not in targets:
                    continue
                page = doc[page_no - 1]
                text, n_spaces = ocr_page(page, ocr_lang, tessdata)
                fixes["ocr_spaces_restored"] += n_spaces
                (ocr_done if len(text) >= 25 else ocr_empty).append(page_no)
                fixed_parts[i] = (page_no, text)
            ocr_seconds = round(time.time() - t0, 1)
            image_only = [p for p in image_only if p in ocr_empty]
            suspect_pages = [p for p in suspect_pages if p not in targets]
            samples = [s for s in samples if int(s[2:s.index(":")]) in suspect_pages]
    total_chars = sum(len(t) for _, t in fixed_parts)

    def assemble(parts, mark_ocr: bool) -> str:
        out = []
        for page_no, text in parts:
            tag = " (OCR)" if mark_ocr and page_no in ocr_done else ""
            out.append(f"<!-- p. {page_no}{tag} -->\n{text}")
        return "\n\n".join(out) + "\n"

    meta = {
        "source": src,
        "file": str(pdf_path),
        "page_count": doc.page_count,
        "pages_extracted": ranges([p for p, _, _ in extracted]),
        "metadata": {k: v for k, v in (doc.metadata or {}).items() if v},
        "toc": [{"level": lvl, "title": title, "page": page} for lvl, title, page in doc.get_toc()],
        "fonts": {f: i for f, i in fonts.items() if i["encoding"] or i["note"]} or "all Unicode",
    }
    (outdir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (outdir / "text.md").write_text(assemble(fixed_parts, True), encoding="utf-8")
    fixes = +fixes  # drop zero counts
    raw_written = bool(fixes or ocr_done)
    if raw_written:
        (outdir / "text.raw.md").write_text(assemble(raw_parts, False), encoding="utf-8")

    summary = {
        "ok": True,
        "title": meta["metadata"].get("title"),
        "page_count": doc.page_count,
        "pages_extracted": meta["pages_extracted"],
        "chars": total_chars,
        "toc_entries": len(meta["toc"]),
        "fixes": dict(fixes),
        "image_only_pages": ranges(image_only),
        "suspect_pages": ranges(suspect_pages),
        "suspect_samples": samples,
        "ocr": None if not ocr else {
            "lang": ocr_lang,
            "pages": ranges(ocr_done),
            "no_text_found": ranges(ocr_empty),
            "seconds": ocr_seconds,
            "error": ocr_error,
        },
        "files": {
            "meta": str(outdir / "meta.json"),
            "text": str(outdir / "text.md"),
            "raw": str(outdir / "text.raw.md") if raw_written else None,
            "pdf": str(pdf_path) if pdf_path.parent == outdir else None,
        },
    }
    # Also on disk: a PowerShell console may mis-decode UTF-8 stdout.
    (outdir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
