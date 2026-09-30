"""Copy ORIGINAL content out of the PPTX based on the classification.

    AI decides WHERE the information is.
    This module copies WHAT the source actually says.

No text produced by the LLM ever reaches the Excel file unless it is found
verbatim in the source (management number / model / item are verified).
"""
from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from pathlib import Path

from .classifier import Classification, is_heading_like, section_kind_of_heading
from .pptx_parser import ReportData, SlideData, norm_key, clean_text

LOG = logging.getLogger("report_extractor.extractor")

# Management Number token: YYMMDD + sequence (>= 9 digits), optional "-VOC"-style suffix.
MGMT_RE = re.compile(r"(?<![0-9A-Za-z])(\d{9,12}(?:-[A-Z]{2,6})?)(?![0-9A-Za-z])")
MODEL_SM_RE = re.compile(r"\bSM-([A-Z]\d{3}[A-Z0-9]{0,3})\b")
MODEL_PLAIN_RE = re.compile(r"(?<![A-Za-z0-9])([AMSFGXNJ]\d{3}[A-Z]?)(?![A-Za-z0-9])")
QTY_RE = re.compile(r"\b\d+\s*(?:ea|pcs?|cái|chiếc|con|sp|units?)\b", re.IGNORECASE)
DEFAULT_ITEMS = ["Rear", "Main", "Sub", "PBA", "Front", "Deco", "Bracket", "Window",
                 "Battery", "Camera", "Key", "Cover", "Housing", "Case", "Frame"]


@dataclass
class Section:
    kind: str                 # cause | temporary | improvement | verify | standard | defect | qpn | other
    slide: int
    lines: List[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return clean_text("\n".join(self.lines))


@dataclass
class ExtractedRecord:
    management_number: str = ""
    model: str = ""
    item: str = ""
    defect_content: str = ""
    root_cause: str = ""
    improvement: str = ""
    temporary_excluded: str = ""      # kept for audit only, never written to improvement cell
    cause_sections: List[Section] = field(default_factory=list)
    improvement_sections: List[Section] = field(default_factory=list)
    qpn_slide: Optional[int] = None
    improvement_image_slides: List[int] = field(default_factory=list)
    review_reasons: List[str] = field(default_factory=list)
    blank_fields: List[str] = field(default_factory=list)      # diagnostics: auto fields left blank
    # Vendor: detected from the ORIGINAL improvement text ("Tại công đoạn assy <Vendor>")
    vendor: str = ""
    vendors: List[str] = field(default_factory=list)            # canonical vendors, source order
    vendors: List[str] = field(default_factory=list)            # canonical vendors, source order
    vendor_candidates: List[str] = field(default_factory=list)  # diagnostics: found + unknown names
    # Ngày phát sinh: derived from the Management Number (YYMMDD prefix) as a real date
    occurrence_date: Optional[date] = None
    # WEEK +1..+8: always blank (never invented)
    weeks: Dict[int, str] = field(default_factory=lambda: {i: "" for i in range(1, 9)})

    @property
    def occurrence_date_text(self) -> str:
        return self.occurrence_date.strftime("%d/%m/%Y") if self.occurrence_date else ""


# ----------------------------------------------------------------------------
# Section walker
# ----------------------------------------------------------------------------
_NOISE_LINE = re.compile(r"^\s*(?:\d{1,3}|[\d/.\-]{6,10}|page \d+|trang \d+|\d+\s*/\s*\d+)\s*$", re.IGNORECASE)


def _is_noise(line: str) -> bool:
    return bool(_NOISE_LINE.match(line)) or not line.strip()


def split_sections(slide: SlideData, default_kind: str = "other") -> List[Section]:
    """Walk a slide in reading order and split text into heading-delimited sections."""
    sections: List[Section] = []
    cur = Section(kind=default_kind, slide=slide.number)
    for b in slide.text_blocks:
        lines = b.text.split("\n")
        for i, ln in enumerate(lines):
            if _is_noise(ln):
                if ln.strip() == "" and cur.lines and cur.lines[-1] != "":
                    cur.lines.append("")
                continue
            heading = is_heading_like(ln, b.bold if i == 0 else False, b.size_pt if i == 0 else None)
            kind = section_kind_of_heading(ln) if heading else None
            if kind:
                # every heading starts a new section (even of the same kind) so
                # sub-sections such as "Nguyên nhân trong kiểm tra" stay separate
                if cur.lines:
                    sections.append(cur)
                cur = Section(kind=kind, slide=slide.number, lines=[ln.strip()])
                continue
            cur.lines.append(ln.rstrip())
        # blank line between separate shapes keeps paragraphs apart
        if cur.lines and cur.lines[-1] != "":
            cur.lines.append("")
    if cur.lines:
        sections.append(cur)
    # strip trailing blanks
    for s in sections:
        while s.lines and s.lines[-1] == "":
            s.lines.pop()
    return [s for s in sections if s.lines]


def collect_sections(report: ReportData, cls: Classification) -> List[Section]:
    """Sections of every classified slide, with the classification supplying the default kind."""
    slide_numbers = sorted(set(cls.cause_slides) | set(cls.improvement_slides) |
                           set(cls.temporary_slides) | set(cls.verify_slides))
    out: List[Section] = []
    carry = "other"
    for n in slide_numbers:
        s = report.slide(n)
        if not s:
            continue
        in_cause = n in cls.cause_slides
        in_imp = n in cls.improvement_slides
        in_tmp = n in cls.temporary_slides
        in_ver = n in cls.verify_slides
        if in_tmp and not in_imp and not in_cause:
            default = "temporary"
        elif in_cause and not in_imp:
            default = "cause"
        elif in_imp and not in_cause:
            default = "improvement"
        elif in_cause and in_imp:
            default = carry if carry in ("cause", "improvement") else "cause"
        elif in_ver:
            default = "verify"
        else:
            default = "other"
        secs = split_sections(s, default)
        # heading-less continuation slide: inherit the previous slide's last kind
        out.extend(secs)
        if secs:
            carry = secs[-1].kind
    return out


def join_sections(sections: Iterable[Section]) -> str:
    parts: List[str] = []
    for s in sections:
        t = s.text
        if t and t not in parts:
            parts.append(t)
    return "\n\n".join(parts)


# ----------------------------------------------------------------------------
# Field extraction
# ----------------------------------------------------------------------------
def management_number_from_filename(filename: str) -> str:
    """Management Number token in a report file name, e.g.

        (CTMS)_11107_251119092-VOC_ Đối sách ... (DRT ) 17.11.2025.pptx  ->  251119092-VOC

        (CTMS)_11108_260601017_ Đối sách ... 05.06.2026.pptx                  ->  260601017

    Two valid forms: ``YYMMDD+seq-SUFFIX`` and bare ``YYMMDD+seq``.  A token is accepted
    only when it is delimited by non-alphanumerics, has >= 9 digits and its first six
    digits form a real YYMMDD calendar date.  CTMS prefixes (11107), file-name dates
    (17.11.2025 / 17112025), model numbers (A185) therefore never match.  A suffixed
    token is preferred when both forms occur.  Returns '' when absent.
    """
    name = Path(filename).name if filename else ""
    suffixed, bare = [], []
    for m in MGMT_RE.finditer(name):
        tok = m.group(1)
        if derive_occurrence_date(tok)[0] is None:
            continue                                  # not YYMMDD-based -> not a Management Number
        (suffixed if "-" in tok else bare).append(tok)
    if suffixed:
        return suffixed[0]
    return bare[0] if bare else ""


def extract_management_number(report: ReportData, llm_value: str = "") -> str:
    """Production rule: the Management Number comes from the PPTX FILE NAME only
    (it is the key of the existing Excel row).  LLM/text values are never used."""
    return management_number_from_filename(report.filename)


def derive_occurrence_date(mgmt: str) -> Tuple[Optional[date], str]:
    """First six digits of the Management Number are YYMMDD (2000-based year).

    251119092-VOC -> 19/11/2025, 260601017 -> 01/06/2026 (suffix is irrelevant).
    Strict calendar validation; returns (None, reason) when the prefix is missing or
    is not a real date.
    """
    m = re.match(r"^\s*(\d{6})\d*(?:-|$)", mgmt or "")
    if not m:
        return None, "Management Number không có tiền tố ngày YYMMDD"
    yy, mm, dd = int(m.group(1)[0:2]), int(m.group(1)[2:4]), int(m.group(1)[4:6])
    try:
        return date(2000 + yy, mm, dd), ""
    except ValueError:
        return None, f"Ngày trong Management Number không hợp lệ (YYMMDD={m.group(1)})"


# ----------------------------------------------------------------------------
# Vendor: controlled multi-value field
# ----------------------------------------------------------------------------
# Canonical spellings written to Excel (user-provided controlled list).  Matching is
# case-/accent-insensitive and whitespace tolerant; output always uses these strings.
CANONICAL_VENDORS: List[str] = ["CNC IT", "Assy IT", "Sơn IT", "Nhựa IT", "Mtech", "Mtech VN", "Mtech 4",
                                "Teawon", "Tesung", "Doaltech", "APV", "An Lập", "Yongsong", "HK", "JT Tech Vina"]
# Known source spellings / historical workbook aliases of canonical vendors (NOT arbitrary
# spelling correction: only variants that clearly identify a controlled-list vendor).
# Used for Vendor identification and comparison only – never to rewrite PPTX or Excel text.
VENDOR_ALIASES: Dict[str, str] = {
    # historical aliases used in the production workbook
    "Lắp ráp IT": "Assy IT", "Sơn Intops": "Sơn IT", "CNC Intops": "CNC IT", "NCC JT": "JT Tech Vina",
    "JT Tech": "JT Tech Vina",
    # known report spellings
    "Taewon": "Teawon", "Tae Won": "Teawon", "Daoltech": "Doaltech", "Doal Tech": "Doaltech",
    "Tae Sung": "Tesung", "Yong Song": "Yongsong", "M-tech": "Mtech", "M-tech VN": "Mtech VN", "M-tech 4": "Mtech 4",
}


def _fold(text: str) -> str:
    """Length-preserving fold: lower-case, one base character per input character."""
    out = []
    for ch in text or "":
        if ch in "đĐ":
            out.append("d")
            continue
        d = unicodedata.normalize("NFD", ch)
        out.append(d[0].lower() if d and unicodedata.category(d[0]) != "Mn" else ch.lower())
    return "".join(out)


def _vendor_pattern(name: str) -> "re.Pattern[str]":
    words = [re.escape(_fold(w)) for w in name.split()]
    return re.compile(r"(?<![a-z0-9])" + r"[\s\-]*".join(words) + r"(?![a-z0-9])")


def _vendor_table(vendors: Optional[List[str]] = None) -> List[Tuple["re.Pattern[str]", str]]:
    """(pattern, canonical) pairs; longer names first so 'Mtech VN' wins over 'Mtech'."""
    canon = list(vendors) if vendors else list(CANONICAL_VENDORS)
    pairs: List[Tuple[str, str]] = [(v, v) for v in canon]
    pairs += [(a, c) for a, c in VENDOR_ALIASES.items() if c in canon]
    pairs.sort(key=lambda x: -len(x[0]))
    return [(_vendor_pattern(a), c) for a, c in pairs]


def find_vendors(text: str, vendors: Optional[List[str]] = None) -> List[Tuple[int, str]]:
    """All controlled-list vendors in ``text`` as (position, canonical) in first-occurrence order."""
    return [(a, c) for a, _b, c in _vendor_spans(text, vendors)]


def _vendor_spans(text: str, vendors: Optional[List[str]] = None) -> List[Tuple[int, int, str]]:
    """(start, end, canonical) of every controlled-list vendor / alias occurrence in ``text``.

    Longer names are matched first and their span is masked, so 'Mtech VN' never also
    yields 'Mtech' and 'Assy IT' never yields a bare 'IT' (which is not a vendor).
    """
    folded = _fold(text)
    taken = [False] * len(folded)
    found: List[Tuple[int, int, str]] = []
    for pat, canon in _vendor_table(vendors):
        for m in pat.finditer(folded):
            if any(taken[m.start():m.end()]):
                continue
            for i in range(m.start(), m.end()):
                taken[i] = True
            found.append((m.start(), m.end(), canon))
    found.sort()
    out: List[Tuple[int, int, str]] = []
    for a, b, c in found:
        if c not in [x[2] for x in out]:
            out.append((a, b, c))
    return out


def _is_only_vendor_names(part: str, vendors: Optional[List[str]] = None) -> bool:
    """True when ``part`` consists (up to punctuation/whitespace) of vendor names/aliases only."""
    spans = _vendor_spans(part, vendors)
    if not spans:
        return False
    covered = sum(b - a for a, b, _ in spans)
    significant = len(re.sub(r"[\s.,;:()\-–/|]+", "", part))
    matched = sum(len(re.sub(r"[\s\-]+", "", part[a:b])) for a, b, _ in spans)
    return covered > 0 and significant - matched <= 2


def canonical_vendors(text: str, vendors: Optional[List[str]] = None) -> List[str]:
    """Canonical vendor names found in ``text`` (deduplicated, source order)."""
    return [c for _, c in find_vendors(text, vendors)]


def canonicalize_vendor_value(value: str, vendors: Optional[List[str]] = None) -> Tuple[List[str], List[str]]:
    """Split an Excel vendor cell (line breaks / commas / semicolons / slashes) into
    (canonical names, unrecognised parts)."""
    canon: List[str] = []
    unknown: List[str] = []
    for part in re.split(r"[\n\r,;/|]+", str(value or "")):
        part = part.strip(" \t.-–")
        if not part:
            continue
        # the part must be (essentially) vendor names / aliases only, not a sentence containing one
        if _is_only_vendor_names(part, vendors):
            for h in canonical_vendors(part, vendors):
                if h not in canon:
                    canon.append(h)
        else:
            unknown.append(part)
    return canon, unknown


# Structural phrase naming a production stage: "Tại công đoạn assy Taewon", "Công đoạn Assy Daoltech"
_VENDOR_PHRASE = re.compile(
    r"(?i:c[oô]ng\s+[dđ]o[aạ]n\s+(?:assy|ass\.?y|assembly))\s*[:\-–]?\s*"
    # the name itself is matched case-sensitively: Capitalised words / upper-case codes only
    r"((?:[A-ZÀ-Ỹ][\w&.\-]*|[A-Z0-9&]{2,})(?:[ \t]+(?:[A-ZÀ-Ỹ][\w&.\-]*|[A-Z0-9&]{2,})){0,3})",
    re.UNICODE)
_VENDOR_STOP = {"tai", "cong", "doan", "assy", "ngay", "tu", "ap", "dung", "loi", "va", "cua", "trong", "sau",
                "truoc", "line", "day", "chuyen", "khu", "vuc", "may", "kiem", "tra", "san", "xuat"}


def _clean_vendor(token: str) -> str:
    t = token.strip(" \t.,;:()[]{}–-/\\")
    t = re.sub(r"\s*\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?.*$", "", t)     # trailing dates
    words = []
    for w in t.split():
        if norm_key(w) in _VENDOR_STOP:
            break
        words.append(w)
    return " ".join(words).strip(" .,;:()–-")


def find_vendor_candidates(text: str) -> List[str]:
    """Names following a 'công đoạn assy …' phrase exactly as written (original spelling),
    de-duplicated case-insensitively.  Used only to report UNKNOWN candidates."""
    out: List[str] = []
    for m in _VENDOR_PHRASE.finditer(text or ""):
        v = _clean_vendor(m.group(1))
        if len(v) < 2:
            continue
        if not any(v.lower() == o.lower() for o in out):
            out.append(v)
    return out


def unknown_vendor_candidates(text: str, vendors: Optional[List[str]] = None) -> List[str]:
    """'công đoạn assy X' names that do not resolve to any controlled-list vendor."""
    out = []
    for cand in find_vendor_candidates(text):
        if _is_only_vendor_names(cand, vendors):
            continue                      # e.g. 'IT' of 'Assy IT' is covered by the canonical phrase
        if norm_key(cand) in ("it",):     # 'IT' alone is never a vendor
            continue
        out.append(cand)
    return out


def extract_vendor(rec: "ExtractedRecord", report: ReportData, vendors: Optional[List[str]] = None) -> None:
    """Vendors (controlled list, possibly several) from the improvement content first; the
    rest of the report is only searched when the improvement content names none.
    Never invents names: unknown 'công đoạn assy X' candidates are reported for review."""
    primary = canonical_vendors(rec.improvement, vendors)
    scope_text = rec.improvement if primary else report.all_text()
    found = primary or canonical_vendors(scope_text, vendors)
    rec.vendors = found
    rec.vendor = "\n".join(found)
    unknown = unknown_vendor_candidates(rec.improvement, vendors)
    if not primary:
        for u in unknown_vendor_candidates(report.all_text(), vendors):
            if u not in unknown:
                unknown.append(u)
    rec.vendor_candidates = found + [u for u in unknown if u not in found]
    if unknown:
        rec.review_reasons.append("Phát hiện tên vendor không có trong danh sách chuẩn (không ghi vào Excel): "
                                  + ", ".join(unknown))
    if not found:
        rec.review_reasons.append("Không xác định được Vendor từ nội dung báo cáo")


def normalize_model(model: str) -> str:
    m = (model or "").strip().upper()
    m = re.sub(r"^SM-", "", m)
    m = re.sub(r"[^A-Z0-9\-]", "", m)
    return m


def extract_model(report: ReportData, llm_value: str = "", known_models: Sequence[str] = ()) -> str:
    candidates_text = [report.filename, report.slides[0].text if report.slides else "", report.all_text()]
    # 1) explicit SM-xxxx
    for t in candidates_text:
        m = MODEL_SM_RE.search(t)
        if m:
            return normalize_model(m.group(1))
    # 2) known models from the Excel mapping sheet
    km = [k for k in known_models if k]
    if km:
        for t in candidates_text:
            up = t.upper()
            hits = [k for k in km if re.search(r"(?<![A-Z0-9])" + re.escape(normalize_model(k)) + r"(?![A-Z0-9])", up)]
            if hits:
                return normalize_model(max(hits, key=len))
    # 3) LLM value, only if verbatim in source
    v = normalize_model(llm_value)
    if v:
        for t in candidates_text:
            if re.search(r"(?<![A-Z0-9])(?:SM-)?" + re.escape(v) + r"(?![A-Z0-9])", t.upper()):
                return v
    # 4) plain pattern in filename / title
    for t in candidates_text[:2]:
        m = MODEL_PLAIN_RE.search(t)
        if m:
            return normalize_model(m.group(1))
    return ""


def extract_item(report: ReportData, cls: Classification, item_mapping: Optional[Dict[str, str]] = None) -> str:
    """Item via Excel mapping (keyword -> item) when available, else source text."""
    texts = [report.filename, report.slides[0].text if report.slides else ""]
    if cls.qpn_slide and report.slide(cls.qpn_slide):
        texts.append(report.slide(cls.qpn_slide).text)
    texts.append(report.all_text())
    if item_mapping:
        # longest keyword first, search title/filename before whole text
        keys = sorted(item_mapping.keys(), key=len, reverse=True)
        for t in texts:
            nt = norm_key(t)
            for k in keys:
                nk = norm_key(k)
                if nk and re.search(r"(?<![a-z0-9])" + re.escape(nk) + r"(?![a-z0-9])", nt):
                    return item_mapping[k]
    v = (cls.item or "").strip()
    if v:
        nv = norm_key(v)
        for t in texts:
            if nv and nv in norm_key(t):
                # return the exact source spelling when possible
                m = re.search(re.escape(v), t, re.IGNORECASE)
                return m.group(0) if m else v
    for t in texts[:3]:
        for it in DEFAULT_ITEMS:
            if re.search(r"(?<![A-Za-z])" + it + r"(?![A-Za-z])", t, re.IGNORECASE):
                return it
    return ""


# ---------------------------------------------------------------------------
# Nội dung lỗi from the file name: "... LỖI BONG ATN, MỤN 22.9.2026 SEV.pptx" -> "BONG ATN, MỤN"
# ---------------------------------------------------------------------------
_DEFECT_MARKER_RE = re.compile(r"(?<![A-Za-zÀ-ỹ])[Ll][ỖỗOo][IiỊị](?![A-Za-zÀ-ỹ])[\s:_\-–]*", re.UNICODE)
_TRAILING_DATE_RE = re.compile(r"(?:^|[\s_\-–(])\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?(?=$|[\s_\-–)])")
_TRAILING_META_WORDS = ("sev", "sevt", "sev t", "dreamtech", "drt", "ctms", "final", "rev", "update", "cap nhat", "ver")


def defect_content_from_filename(filename: str) -> str:
    """Defect names written in the report file name (after the marker LỖI), trailing metadata
    (date, SEV, parentheses, extension, underscores) removed.  Returns '' when the marker is
    absent or nothing reliable remains.  Never adds quantities or rewrites the words."""
    stem = Path(filename).stem if filename else ""
    m = None
    for m in _DEFECT_MARKER_RE.finditer(stem):      # the LAST 'LỖI' marker starts the defect list
        pass
    if m is None:
        return ""
    tail = stem[m.end():]
    # cut at the first date token, then drop trailing parentheticals / metadata words
    dm = _TRAILING_DATE_RE.search(tail)
    if dm:
        tail = tail[:dm.start()]
    tail = re.sub(r"\([^)]*\)\s*$", "", tail.strip())
    tail = re.sub(r"\([^)]*\)", " ", tail)
    words = tail.replace("_", " ").split()
    def _meta(w: str) -> bool:
        k = norm_key(w)
        return (k in _TRAILING_META_WORDS or re.fullmatch(r"[\W_]+", w) is not None
                or MODEL_SM_RE.fullmatch(w.upper()) is not None or MODEL_PLAIN_RE.fullmatch(w.upper()) is not None
                or re.fullmatch(r"SM-[A-Z]\d{3}[A-Z0-9]{0,3}", w.upper()) is not None
                or k in {i.lower() for i in DEFAULT_ITEMS} or k in ("model", "item"))
    while words and _meta(words[-1]):          # trailing item / model / metadata words
        words.pop()
    text = " ".join(words).strip(" ,;:-–_")
    text = re.sub(r"\s*,\s*", ", ", text)
    text = re.sub(r"\s+", " ", text)
    if len(text) < 2 or not re.search(r"[A-Za-zÀ-ỹ]", text):
        return ""
    return text


def extract_defect_content(report: ReportData, cls: Classification) -> str:
    """Nội dung lỗi: (1) defect names from the file name; (2) defect text on the QPN/defect slide."""
    from_name = defect_content_from_filename(report.filename)
    if from_name:
        return from_name
    return extract_defect_content_from_slides(report, cls)


def extract_defect_content_from_slides(report: ReportData, cls: Classification) -> str:
    # only the defect / QPN slides are sources; the cover, titles and the file name never are
    slides: List[SlideData] = []
    for n in [cls.defect_slide, cls.qpn_slide]:
        s = report.slide(n) if n else None
        if s and s not in slides:
            slides.append(s)
    # 1) quantity lines ("Xước: 15ea")
    for s in slides:
        lines: List[str] = []
        for b in s.text_blocks:
            if b.kind == "table":
                for row in b.rows:
                    for c in row:
                        if c and QTY_RE.search(c):
                            lines.append(c.strip())
                continue
            for ln in b.text.split("\n"):
                if QTY_RE.search(ln) and len(ln) < 200:
                    lines.append(ln.strip())
        uniq = []
        for ln in lines:
            if ln not in uniq:
                uniq.append(ln)
        if uniq:
            return "\n".join(uniq)
    # 2) heading-delimited defect section (HIỆN TRẠNG / NỘI DUNG LỖI ...), heading line itself dropped
    for s in slides:
        for sec in split_sections(s, "other"):
            if sec.kind == "defect":
                body = [ln for ln in sec.lines[1:]] if section_kind_of_heading(sec.lines[0]) == "defect" else sec.lines
                txt = clean_text("\n".join(body))
                if txt:
                    return txt
    return ""


def defect_only_in_image(report: ReportData, cls: Classification) -> bool:
    """True when the QPN/defect slide carries a picture but its defect region (text before the
    first cause/temporary/improvement heading) has no copyable text.  Text of other sections
    on the same slide (e.g. NGUYÊN NHÂN under the QPN form) does not count as defect text."""
    for n in [cls.defect_slide, cls.qpn_slide]:
        ds = report.slide(n) if n else None
        if ds is None or not ds.pictures:
            continue
        region_text = []
        for sec in split_sections(ds, "defect"):
            if sec.kind in ("cause", "temporary", "improvement", "verify", "standard"):
                break
            for ln in sec.lines:
                if section_kind_of_heading(ln) is None and not _is_noise(ln) and ln.strip():
                    region_text.append(ln)
        if not region_text:
            return True
    return False


def extract_record(report: ReportData, cls: Classification,
                   item_mapping: Optional[Dict[str, str]] = None,
                   known_models: Sequence[str] = (), vendors: Optional[List[str]] = None) -> ExtractedRecord:
    rec = ExtractedRecord()
    rec.qpn_slide = cls.qpn_slide
    rec.management_number = extract_management_number(report, cls.management_number)
    rec.model = extract_model(report, cls.model, known_models)
    rec.item = extract_item(report, cls, item_mapping)
    rec.defect_content = extract_defect_content(report, cls)

    sections = collect_sections(report, cls)
    rec.cause_sections = [s for s in sections if s.kind == "cause"]
    rec.improvement_sections = [s for s in sections if s.kind in ("improvement", "standard")]
    tmp_sections = [s for s in sections if s.kind == "temporary"]
    rec.root_cause = join_sections(rec.cause_sections)
    rec.improvement = join_sections(rec.improvement_sections)
    rec.temporary_excluded = join_sections(tmp_sections)
    rec.improvement_image_slides = list(cls.improvement_image_slides)

    # Vendor (from source text) and Ngày phát sinh (from Management Number)
    extract_vendor(rec, report, vendors)
    if rec.management_number:
        rec.occurrence_date, why = derive_occurrence_date(rec.management_number)
        if why:
            rec.review_reasons.append(why)
    else:
        rec.review_reasons.append("Không xác định được Management Number từ tên file")

    # validation -> review reasons (never invent content)
    if not rec.model:
        rec.review_reasons.append("Không tìm thấy Model trong báo cáo")
    if not rec.qpn_slide:
        rec.review_reasons.append("Không tìm thấy QPN trong báo cáo")
    if not rec.root_cause:
        rec.review_reasons.append("Không tìm thấy Nguyên nhân trong báo cáo")
    if not rec.improvement:
        rec.review_reasons.append("Không tìm thấy Nội dung đối sách cải tiến trong báo cáo")
    if not rec.improvement_image_slides:
        rec.review_reasons.append("Không có hình ảnh cải tiến")
    # ambiguous classification -> manual review rather than guessing
    rec.review_reasons.extend(cls.ambiguities)
    for name, val in (("Management number", rec.management_number), ("Tên vendor", rec.vendor),
                      ("Ngày phát sinh", rec.occurrence_date_text), ("Model", rec.model), ("Item", rec.item),
                      ("Nội dung lỗi", rec.defect_content), ("Nguyên nhân", rec.root_cause),
                      ("Nội dung đối sách cải tiến", rec.improvement)):
        if not val:
            rec.blank_fields.append(name)
    if not rec.defect_content:
        if defect_only_in_image(report, cls):
            rec.review_reasons.append(
                "Nội dung lỗi chỉ có trong hình ảnh QPN, không có text để sao chép – cần bổ sung thủ công")
        else:
            rec.review_reasons.append("Không tìm thấy Nội dung lỗi trong báo cáo")
    return rec
