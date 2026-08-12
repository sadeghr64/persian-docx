# -*- coding: utf-8 -*-
"""fa_docx — تبدیل Markdown و CSV به Word فارسی.
راست‌به‌چپ سه‌لایه · راست‌چین · B Titr / B Nazanin · فهرست خودکار · اصلاح دوجهته.

    python fa_docx.py --root <پوشه>
    python fa_docx.py --file <x.md> [-o y.docx]
    python fa_docx.py --root . --kicker "نام سازمان" --date "۱۴۰۵/۰۶/۰۱"

مرجع قواعد: SKILL.md در پوشه بالاتر.

قواعد کلیدی که در این ابزار رعایت شده و در مهارت «persian-docx» مستند است:
  ۱. اندازه فونت فارسی با w:szCs تعیین می‌شود، نه w:sz — بدون آن اندازه اعمال نمی‌شود.
  ۲. خانواده «B Titr» فقط چهره Bold دارد؛ همیشه با bold=True اعمال شود.
  ۳. راست‌به‌چپ در سه لایه: w:bidi (بخش و پاراگراف)، w:bidiVisual (جدول)، w:rtl (ران).
  ۴. فهرست خودکار با فیلد TOC + w:updateFields تا هنگام باز شدن به‌روز شود.
  ۵. در پاراگراف RTL، w:jc val="left" یعنی بصری راست‌چین؛ "right" یعنی چپ‌چین.
"""
import os, re, csv, sys, io
sys.stdout.reconfigure(encoding="utf-8")

from docx import Document
from docx.shared import Pt, RGBColor, Emu
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = os.getcwd()
KICKER = ""      # نوار بالای کاور — با ‎--kicker
AS_OF = ""       # تاریخ پانویس — با ‎--date

# ---------------------------------------------------------------- فونت و اندازه
F_TITR = "B Titr"        # فقط چهره Bold دارد → همیشه bold=True
F_BODY = "B Nazanin"     # دارای Regular و Bold
F_MONO = "Consolas"

SZ_H1, SZ_H2, SZ_H3, SZ_H4 = 18, 15, 13.5, 12.5
SZ_BODY = 13
SZ_TOC_TITLE = 16
LS_BODY = 1.45           # فاصله خطوط فارسی بیشتر از لاتین لازم دارد

INK     = RGBColor(0x1A, 0x1D, 0x24)
MUTED   = RGBColor(0x5B, 0x62, 0x70)
BRAND   = RGBColor(0xC8, 0x92, 0x2A)
BRAND_D = RGBColor(0x8A, 0x64, 0x10)
BLUE    = RGBColor(0x2C, 0x6E, 0x8F)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
SLATE   = RGBColor(0x2A, 0x2F, 0x3A)

SH_HEAD, SH_BRANDL, SH_RUSTL = "EEF1F5", "FBF3E2", "FDF0EC"
SH_BLUEL, SH_ZEBRA, SH_CODE = "EEF2F7", "FAFBFC", "F4F6F8"

# ---------------------------------------------------------------- نگارش فارسی
_FA = "۰۱۲۳۴۵۶۷۸۹"
ZWNJ = "‌"
LRM = "‎"   # نشانگر چپ‌به‌راست — ترتیب درست ارقام درون متن راست‌به‌چپ
RLM = "‏"   # نشانگر راست‌به‌چپ — جداسازی قطعه لاتین درون متن فارسی

# B Nazanin / B Titr این نویسه‌ها را ندارند؛ جایگزین می‌شوند تا وابسته به فونت پشتیبان نباشیم
_MISSING_MAP = {
    "٫": "/",    # ممیز عربی U+066B → ممیز رایج فارسی (در فونت موجود و خوانا)
    "٬": "،",    # هزارگان عربی U+066C → ویرگول فارسی (U+066C در فونت نیست و شبیه صفر دیده می‌شود)
    "‑": "-",    # خط تیره بدون شکست U+2011
}


def fa_digits(s):
    return "".join(_FA[int(c)] if c.isdigit() else c for c in str(s))


def fa_number(s):
    """عدد لاتین → فارسی؛ هزارگان «٬» و ممیز «/» (قرارداد رایج اسناد فارسی)"""
    s = s.strip().replace(",", "")
    neg = s.startswith("-")
    if neg:
        s = s[1:]
    if not re.fullmatch(r"\d+(\.\d+)?", s):
        return None
    ip, _, fp = s.partition(".")
    ip = f"{int(ip):,}".replace(",", "،")
    out = fa_digits(ip) + (("/" + fa_digits(fp)) if fp else "")
    return ("−" + out) if neg else out


# نویسه‌های عربی که باید فارسی شوند
_ARABIC_MAP = {
    "ي": "ی",  # ي → ی
    "ى": "ی",  # ى → ی
    "ك": "ک",  # ك → ک
    "ڪ": "ک",  # ڪ → ک
    "٠": "۰", "١": "۱", "٢": "۲", "٣": "۳", "٤": "۴",
    "٥": "۵", "٦": "۶", "٧": "۷", "٨": "۸", "٩": "۹",
    "٫": "٫", "٬": "٬",
}
_ARABIC_RE = re.compile("|".join(map(re.escape, _ARABIC_MAP)))

# نیم‌فاصله‌های پرکاربرد و کم‌خطر
_ZWNJ_RULES = [
    (re.compile(r"(?<![؀-ۿ])می\s+(?=[؀-ۿ])"), "می" + ZWNJ),
    (re.compile(r"(?<![؀-ۿ])نمی\s+(?=[؀-ۿ])"), "نمی" + ZWNJ),
    (re.compile(r"(?<=[؀-ۿ])\s+(ها|های|هایی|هایش)(?![؀-ۿ])"), ZWNJ + r"\1"),
    (re.compile(r"(?<=[؀-ۿ])\s+(تر|تری|ترین)(?![؀-ۿ])"), ZWNJ + r"\1"),
]

# فاصله‌گذاری نشانه‌ها: بدون فاصله پیش از نشانه، یک فاصله پس از آن
_PUNC_BEFORE = re.compile(r"\s+([،؛؟!:])")
_PUNC_AFTER = re.compile(r"([،؛؟!])(?=[^\s\d‌])")

# عدد/تاریخ/نسبت فارسی به همراه جداکننده‌های داخلی — باید در LRM محصور شود
_NUM_TOKEN = re.compile(r"[۰-۹]+(?:[٬،/.\-:][۰-۹]+)*٪?")


def wrap_numbers(s):
    """محصورکردن هر عدد در LRM تا ترتیب ارقام در متن راست‌به‌چپ وارونه نشود.
    بدون این کار «۱۴۰۳/۰۹/۰۴» به‌صورت «۰۴/۰۹/۱۴۰۳» و «۱٬۳۹۷/۶» وارونه دیده می‌شود."""
    if LRM in s:
        return s
    return _NUM_TOKEN.sub(lambda m: LRM + m.group(0) + LRM, s)


def fa_text(s, zwnj=True, numbers=True):
    """پاک‌سازی نگارشی متن فارسی. روی کد و پیوند اعمال نمی‌شود."""
    if not s:
        return s
    s = _ARABIC_RE.sub(lambda m: _ARABIC_MAP[m.group()], s)
    for k, v in _MISSING_MAP.items():
        s = s.replace(k, v)
    if zwnj:
        for rx, rep in _ZWNJ_RULES:
            s = rx.sub(rep, s)
    s = _PUNC_BEFORE.sub(r"\1", s)
    s = _PUNC_AFTER.sub(r"\1 ", s)
    s = re.sub(r"[ \t]{2,}", " ", s)
    if numbers:
        s = wrap_numbers(s)
    return s


# ---------------------------------------------------------------- تراز (align)
# تله بحرانی: در پاراگراف راست‌به‌چپ (w:bidi) مقادیر left/right در OOXML
# «آغاز/پایان» معنا می‌دهند، نه چپ/راست بصری. یعنی:
#     w:jc val="left"  → بصری راست‌چین ✅
#     w:jc val="right" → بصری چپ‌چین  ❌
# و python-docx هنگام alignment=RIGHT دقیقاً «right» می‌نویسد → چپ‌چین می‌شود.
# بنابراین هر تراز باید از این نگاشت عبور کند.
_JC_RTL = {"right": "left", "left": "right", "center": "center", "justify": "both"}
_JC_LTR = {"right": "right", "left": "left", "center": "center", "justify": "both"}


def set_align(p, visual="right", rtl=True):
    """تراز بصری. visual یکی از right / left / center / justify."""
    jc = p._p.get_or_add_pPr().get_or_add_jc()
    jc.set(qn("w:val"), (_JC_RTL if rtl else _JC_LTR)[visual])
    return p


def _ltr_para(p):
    """پاراگراف چپ‌به‌راست صریح (برای بلوک کد لاتین) — bidi ارثی از استایل را خنثی می‌کند."""
    e = OxmlElement("w:bidi"); e.set(qn("w:val"), "0")
    p._p.get_or_add_pPr().append(e)


# ---------------------------------------------------------------- XML کمکی
def _rtl_para(p):
    """w:bidi را یک‌بار و در جای درست درج می‌کند (طبق ترتیب طرح‌واره، پیش از w:jc)."""
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is not None:
        return p
    e = OxmlElement("w:bidi")
    jc = pPr.find(qn("w:jc"))
    if jc is not None:
        jc.addprevious(e)
    else:
        pPr.append(e)
    return p


def _shade(el, hexc):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexc)
    el.append(shd)


def _cell_shade(cell, hexc):
    _shade(cell._tc.get_or_add_tcPr(), hexc)


def _para_border(p, sides=("bottom",), sz=6, color="C8922A", space=3):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    for s in sides:
        e = OxmlElement(f"w:{s}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), str(space)); e.set(qn("w:color"), color)
        pbdr.append(e)
    pPr.append(pbdr)


def _cell_border(cell, **sides):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for side, (sz, color) in sides.items():
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), sz)
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), color)
        b.append(e)
    tcPr.append(b)


def run(p, text, size=SZ_BODY, color=INK, bold=False, italic=False, font=F_BODY, rtl=True):
    """ران فارسی: اندازه باید هم روی w:sz و هم w:szCs بنشیند."""
    if font == F_TITR:
        bold = True                      # B Titr فقط چهره Bold دارد
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.color.rgb = color
    r.bold = bold; r.italic = italic
    rpr = r._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a in ("w:cs", "w:ascii", "w:hAnsi", "w:eastAsia"):
        rf.set(qn(a), font)
    szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), str(int(size * 2)))
    rpr.append(szcs)
    if rtl:
        e = OxmlElement("w:rtl"); e.set(qn("w:val"), "1"); rpr.append(e)
    if bold:
        rpr.append(OxmlElement("w:bCs"))
    if italic:
        rpr.append(OxmlElement("w:iCs"))
    return r


def para(host, align="right", after=6, before=0, ls=LS_BODY):
    p = host.add_paragraph()
    _rtl_para(p)
    set_align(p, align)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = ls
    return p


def add_hyperlink(p, url, text, size=SZ_BODY - 1):
    r_id = p.part.relate_to(url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    hl = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), r_id)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("w:cs", "w:ascii", "w:hAnsi"):
        rf.set(qn(a), F_BODY)
    rPr.append(rf)
    c = OxmlElement("w:color"); c.set(qn("w:val"), "2C6E8F"); rPr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    for tag in ("w:sz", "w:szCs"):
        e = OxmlElement(tag); e.set(qn("w:val"), str(int(size * 2))); rPr.append(e)
    r.append(rPr)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); r.append(t)
    hl.append(r); p._p.append(hl)


# ---------------------------------------------------------------- inline
TOKEN = re.compile(r"(\*\*.+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))", re.S)


def emit_inline(p, text, size=SZ_BODY, color=INK, base_bold=False, font=F_BODY):
    for tok in TOKEN.split(text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            run(p, fa_text(tok[2:-2]), size=size, color=color, bold=True, font=font)
        elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
            # قطعه کد با RLM محصور می‌شود تا مرزش با متن فارسی اطراف مخدوش نشود
            run(p, RLM + tok[1:-1] + RLM, size=size - 1.5, color=BLUE, font=F_MONO, rtl=False)
        elif tok.startswith("[") and "](" in tok:
            label, url = tok[1:-1].split("](", 1)
            if url.startswith(("http://", "https://")):
                add_hyperlink(p, url, fa_text(label), size=size - 1)
            else:
                run(p, fa_text(label), size=size, color=BLUE, bold=base_bold, font=font)
        else:
            run(p, fa_text(tok), size=size, color=color, bold=base_bold, font=font)


# ---------------------------------------------------------------- جدول
def page_full(doc):
    s = doc.sections[0]
    return int(s.page_width - s.left_margin - s.right_margin)


def _table_layout(tbl, widths):
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    tblPr.append(OxmlElement("w:bidiVisual"))       # جدول راست‌به‌چپ
    grid = tbl._tbl.find(qn("w:tblGrid"))
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(widths[i] // 635))
    for row in tbl.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Emu(widths[i])


def _table_borders(tbl):
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), "D5DAE2")
        b.append(e)
    tbl._tbl.tblPr.append(b)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))          # تکرار سربرگ در صفحه بعد


def _spacer(doc, pt=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(pt); p.paragraph_format.space_before = Pt(0)


def calc_widths(headers, rows, full):
    n = len(headers)
    ml = []
    for i in range(n):
        L = len(headers[i]) if i < len(headers) else 1
        for r in rows:
            if i < len(r):
                L = max(L, min(len(r[i]), 90))
        ml.append(max(L, 3))
    tot = sum(ml)
    raw = [max(m / tot, 0.055 if n > 6 else 0.08) for m in ml]
    s = sum(raw)
    w = [int(x / s * full) for x in raw]
    w[-1] += full - sum(w)
    return w


def cell_value(v):
    """عدد لاتین در سلول → عدد فارسی با جداکننده"""
    n = fa_number(v)
    return n if n is not None else fa_text(v)


def add_table(doc, headers, rows, full):
    n = max(len(headers), max((len(r) for r in rows), default=0))
    headers = list(headers) + [""] * (n - len(headers))
    rows = [list(r) + [""] * (n - len(r)) for r in rows]
    has_header = any(h.strip() for h in headers)
    fs = 12 if n <= 4 else 11 if n <= 6 else 10 if n <= 8 else 9.5

    tbl = doc.add_table(rows=1 if has_header else 0, cols=n)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    _table_layout(tbl, calc_widths(headers, rows, full))

    if has_header:
        _repeat_header(tbl.rows[0])
        for i, hd in enumerate(headers):
            c = tbl.rows[0].cells[i]
            _cell_shade(c, SH_HEAD); c.vertical_alignment = 1
            p = c.paragraphs[0]; _rtl_para(p)
            set_align(p, "center")
            p.paragraph_format.space_after = Pt(1.5); p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.line_spacing = 1.2
            emit_inline(p, hd, size=fs, color=SLATE, base_bold=True)

    for ri, rowdata in enumerate(rows):
        cells = tbl.add_row().cells
        for ci, val in enumerate(rowdata):
            c = cells[ci]
            if ri % 2 == 1:
                _cell_shade(c, SH_ZEBRA)
            c.vertical_alignment = 1
            p = c.paragraphs[0]; _rtl_para(p)
            txt = cell_value(val)
            set_align(p, "center" if len(txt) <= 14
                      and any(ch in _FA or ch.isdigit() for ch in txt) else "right")
            p.paragraph_format.space_after = Pt(1.5); p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.line_spacing = 1.25
            emit_inline(p, txt, size=fs, base_bold=(not has_header and ci == 0 and n == 2))
    _table_borders(tbl)
    _spacer(doc)
    return tbl


def note_box(doc, lines, full, kind="brand"):
    fill = {"brand": SH_BRANDL, "warn": SH_RUSTL, "info": SH_BLUEL}[kind]
    edge = {"brand": "C8922A", "warn": "C4553B", "info": "2C6E8F"}[kind]
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_layout(tbl, [full])
    c = tbl.cell(0, 0)
    _cell_shade(c, fill); _cell_border(c, right=("18", edge))
    c.paragraphs[0]._p.getparent().remove(c.paragraphs[0]._p)
    for ln in lines:
        p = para(c, after=2, ls=1.35)
        emit_inline(p, ln, size=SZ_BODY - 0.5)
    _spacer(doc, 6)


def code_box(doc, lines, full):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_layout(tbl, [full])
    c = tbl.cell(0, 0)
    _cell_shade(c, SH_CODE); _cell_border(c, right=("12", "C8CFD8"))
    c.paragraphs[0]._p.getparent().remove(c.paragraphs[0]._p)
    # Consolas هم «٬» و «٫» را ندارد. fa_text عمداً روی کد اجرا نمی‌شود (درست است)،
    # پس نگاشت نویسه‌های گمشده باید جداگانه روی سطرهای کد اعمال شود — SKILL.md §۵.
    lines = [ln.replace("٫", "/").replace("٬", "،") for ln in lines]
    for ln in lines:
        p = c.add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        if any("؀" <= ch <= "ۿ" for ch in ln):
            _rtl_para(p); set_align(p, "right")
            run(p, ln or " ", size=11, color=SLATE, font=F_BODY)
        else:
            _ltr_para(p); set_align(p, "left", rtl=False)
            run(p, ln or " ", size=9, color=SLATE, font=F_MONO, rtl=False)
    _spacer(doc, 6)


# ---------------------------------------------------------------- عنوان‌ها
def heading(doc, level, text):
    if level == 1:
        p = para(doc, before=10, after=10, ls=1.3)
        p.style = doc.styles["Heading 1"]
        _rtl_para(p); set_align(p, "right")
        _para_border(p, ("bottom",), sz=18, color="C8922A")
        emit_inline(p, text, size=SZ_H1, color=BRAND_D, font=F_TITR)
        for r in p.runs: r.font.color.rgb = BRAND_D
    elif level == 2:
        p = para(doc, before=16, after=6, ls=1.3)
        p.style = doc.styles["Heading 2"]
        _rtl_para(p); set_align(p, "right")
        _para_border(p, ("bottom",), sz=12, color="C8922A")
        emit_inline(p, text, size=SZ_H2, color=INK, base_bold=True)
        for r in p.runs: r.font.color.rgb = INK
    elif level == 3:
        p = para(doc, before=12, after=4, ls=1.3)
        p.style = doc.styles["Heading 3"]
        _rtl_para(p); set_align(p, "right")
        emit_inline(p, text, size=SZ_H3, color=SLATE, base_bold=True)
        for r in p.runs: r.font.color.rgb = SLATE
    else:
        p = para(doc, before=9, after=3, ls=1.3)
        p.style = doc.styles["Heading 4"]
        _rtl_para(p); set_align(p, "right")
        emit_inline(p, text, size=SZ_H4, color=BRAND_D, base_bold=True)
        for r in p.runs: r.font.color.rgb = BRAND_D
    return p


def bullet(doc, text, size=SZ_BODY, mark="•"):
    p = para(doc, after=3, ls=1.35)
    p.paragraph_format.right_indent = Pt(16)
    run(p, f"{mark}  ", size=size, color=BRAND, bold=True)
    emit_inline(p, text, size=size)
    return p


def numbered(doc, num, text, size=SZ_BODY):
    p = para(doc, after=3, ls=1.35)
    p.paragraph_format.right_indent = Pt(16)
    run(p, f"{fa_digits(num)}. ", size=size, color=BRAND_D, bold=True)
    emit_inline(p, text, size=size)
    return p


def hrule(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(7); p.paragraph_format.space_before = Pt(3)
    _para_border(p, ("bottom",), sz=6, color="DFE3E9", space=1)


# ---------------------------------------------------------------- سند پایه
def _style_font(style, font, size, bold=False, color=None):
    style.font.size = Pt(size)
    style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
    rPr = style.element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rPr.append(rf)
    for a in ("w:cs", "w:ascii", "w:hAnsi", "w:eastAsia"):
        rf.set(qn(a), font)
    for tag in ("w:sz", "w:szCs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = OxmlElement(tag); rPr.append(e)
        e.set(qn("w:val"), str(int(size * 2)))
    if bold:
        for tag in ("w:b", "w:bCs"):
            if rPr.find(qn(tag)) is None:
                rPr.append(OxmlElement(tag))
    e = OxmlElement("w:rtl"); e.set(qn("w:val"), "1"); rPr.append(e)
    pPr = style.element.get_or_add_pPr()
    pPr.append(OxmlElement("w:bidi"))
    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "left"); pPr.append(jc)   # ← بصری راست‌چین


def new_doc(landscape=False):
    doc = Document()
    _style_font(doc.styles["Normal"], F_BODY, SZ_BODY)
    _style_font(doc.styles["Heading 1"], F_TITR, SZ_H1, True, BRAND_D)
    _style_font(doc.styles["Heading 2"], F_BODY, SZ_H2, True, INK)
    _style_font(doc.styles["Heading 3"], F_BODY, SZ_H3, True, SLATE)
    _style_font(doc.styles["Heading 4"], F_BODY, SZ_H4, True, BRAND_D)
    for nm, sz in (("TOC 1", SZ_BODY), ("TOC 2", SZ_BODY - 0.5), ("TOC 3", SZ_BODY - 1)):
        try:
            st = doc.styles[nm]
        except KeyError:
            st = doc.styles.add_style(nm, WD_STYLE_TYPE.PARAGRAPH)
        _style_font(st, F_BODY, sz, bold=(nm == "TOC 1"))
    for sec in doc.sections:
        if landscape:
            sec.page_width, sec.page_height = Emu(int(11.69 * 914400)), Emu(int(8.27 * 914400))
        else:
            sec.page_width, sec.page_height = Emu(int(8.27 * 914400)), Emu(int(11.69 * 914400))
        sec.left_margin = sec.right_margin = Emu(int(0.75 * 914400))
        sec.top_margin = sec.bottom_margin = Emu(int(0.7 * 914400))
        sec._sectPr.append(OxmlElement("w:bidi"))       # بخش راست‌به‌چپ
    return doc


def enable_update_fields(doc):
    """تا Word هنگام باز شدن، فهرست را خودکار به‌روز کند."""
    s = doc.settings.element
    for tag in ("w:updateFields",):
        e = s.find(qn(tag))
        if e is None:
            e = OxmlElement(tag); s.append(e)
        e.set(qn("w:val"), "true")


def add_toc(doc, full, levels="1-3"):
    # عنوان فهرست عمداً سبک Heading نمی‌گیرد تا داخل خودِ فهرست ظاهر نشود
    t = para(doc, before=4, after=8, ls=1.3)
    _para_border(t, ("bottom",), sz=12, color="C8922A")
    run(t, "فهرست مطالب", size=SZ_TOC_TITLE, color=BRAND_D, font=F_TITR)
    p = para(doc, after=4)
    for kind, txt in (("begin", None), (None, f' TOC \\o "{levels}" \\h \\z \\u '),
                      ("separate", None)):
        r = OxmlElement("w:r")
        if kind:
            f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), kind)
            if kind == "begin":
                f.set(qn("w:dirty"), "true")
            r.append(f)
        else:
            t = OxmlElement("w:instrText"); t.set(qn("xml:space"), "preserve"); t.text = txt
            r.append(t)
        p._p.append(r)
    run(p, "برای به‌روزرسانی فهرست: Ctrl+A سپس F9", size=SZ_BODY - 2, color=MUTED, italic=True)
    r = OxmlElement("w:r")
    f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "end"); r.append(f)
    p._p.append(r)
    # شکست صفحه پس از فهرست — با pageBreakBefore، نه پاراگراف خالیِ حاوی break:
    # پاراگراف خالی + WD_BREAK.PAGE یک صفحه کاملاً خالی می‌سازد — SKILL.md §۷-۱.
    bp = doc.add_paragraph()
    pf = bp.paragraph_format
    pf.page_break_before = True
    pf.space_before = pf.space_after = Pt(0); pf.line_spacing = 1
    bp.add_run().font.size = Pt(1)


def cover(doc, title, subtitle, full):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_layout(tbl, [full])
    c = tbl.cell(0, 0)
    _cell_shade(c, "1F2430")
    c.paragraphs[0]._p.getparent().remove(c.paragraphs[0]._p)
    c.add_paragraph().paragraph_format.space_after = Pt(2)
    if KICKER:
        p = para(c, after=5); run(p, KICKER, size=SZ_BODY - 2, color=BRAND, bold=True)
    p = para(c, after=5, ls=1.25)
    emit_inline(p, title, size=SZ_H1 + 2, color=WHITE, font=F_TITR)
    for r in p.runs: r.font.color.rgb = WHITE
    if subtitle:
        p = para(c, after=6)
        emit_inline(p, subtitle, size=SZ_BODY - 1, color=RGBColor(0xC8, 0xCD, 0xD6))
        for r in p.runs: r.font.color.rgb = RGBColor(0xC8, 0xCD, 0xD6)
    c.add_paragraph().paragraph_format.space_after = Pt(2)
    _spacer(doc, 8)


def _page_field(p, size=9.5):
    """شماره صفحه به‌صورت «صفحه X از Y» با ارقام فارسی."""
    def _fa_rpr():
        """زبان ران روی فارسی تا Word ارقام فیلد را فارسی نمایش دهد."""
        rPr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        for a in ("w:cs", "w:ascii", "w:hAnsi"):
            rf.set(qn(a), F_BODY)
        rPr.append(rf)
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag); e.set(qn("w:val"), str(int(size * 2))); rPr.append(e)
        c = OxmlElement("w:color"); c.set(qn("w:val"), "5B6270"); rPr.append(c)
        lang = OxmlElement("w:lang")
        lang.set(qn("w:bidi"), "fa-IR"); lang.set(qn("w:val"), "fa-IR")
        rPr.append(lang)
        e = OxmlElement("w:rtl"); e.set(qn("w:val"), "1"); rPr.append(e)
        return rPr

    def fld(instr):
        for kind, txt in (("begin", None), (None, instr), ("separate", None),
                          (None, None), ("end", None)):
            r = OxmlElement("w:r")
            r.append(_fa_rpr())
            if kind:
                f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), kind); r.append(f)
            elif txt:
                t = OxmlElement("w:instrText"); t.set(qn("xml:space"), "preserve")
                t.text = txt; r.append(t)
            else:
                t = OxmlElement("w:t"); t.text = "۱"; r.append(t)
            p._p.append(r)
    run(p, "صفحه ", size=size, color=MUTED)
    fld(r' PAGE ')
    run(p, " از ", size=size, color=MUTED)
    fld(r' NUMPAGES ')


def add_footer(doc, title):
    for sec in doc.sections:
        p = sec.footer.paragraphs[0]
        _rtl_para(p); set_align(p, "center")
        bits = [b for b in (KICKER.split(" · ")[0] if KICKER else "", title, AS_OF) if b]
        run(p, fa_text(" · ".join(bits)) + "  |  ", size=9.5, color=MUTED)
        _page_field(p)


# ---------------------------------------------------------------- تجزیه Markdown
def split_row(line):
    s = line.strip()
    if s.startswith("|"): s = s[1:]
    if s.endswith("|"): s = s[:-1]
    return [c.strip() for c in s.split("|")]


def is_sep(line):
    s = line.strip().strip("|")
    return bool(s) and all(set(c.strip()) <= set("-: ") and "-" in c for c in s.split("|"))


def render_md(doc, text, full):
    lines = text.split("\n")
    i, n = 0, len(lines)
    skipped_h1 = False
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1; continue

        if s.startswith("```"):
            j = i + 1; buf = []
            while j < n and not lines[j].strip().startswith("```"):
                buf.append(lines[j]); j += 1
            code_box(doc, buf, full); i = j + 1; continue

        if re.fullmatch(r"-{3,}|\*{3,}|_{3,}", s):
            hrule(doc); i += 1; continue

        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            lvl, txt = len(m.group(1)), m.group(2).strip()
            if lvl == 1 and not skipped_h1:
                skipped_h1 = True; i += 1; continue
            heading(doc, min(lvl, 4), txt); i += 1; continue

        if s.startswith("|") and i + 1 < n and is_sep(lines[i + 1]):
            headers = split_row(s)
            j = i + 2; rows = []
            while j < n and lines[j].strip().startswith("|"):
                if not is_sep(lines[j]):
                    rows.append(split_row(lines[j]))
                j += 1
            add_table(doc, headers, rows, full); i = j; continue

        if s.startswith(">"):
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip()); i += 1
            buf = [b for b in buf if b]
            joined = " ".join(buf)
            kind = ("warn" if any(k in joined for k in
                    ("هشدار", "ممنوع", "توقف", "خطا", "بحران", "ریسک انطباق")) else "brand")
            note_box(doc, buf, full, kind=kind); continue

        m = re.match(r"^[-*]\s+\[( |x|X)\]\s+(.*)$", s)
        if m:
            bullet(doc, m.group(2), mark="☐" if m.group(1) == " " else "☑"); i += 1; continue

        m = re.match(r"^[-*]\s+(.*)$", s)
        if m:
            bullet(doc, m.group(1)); i += 1; continue

        m = re.match(r"^(\d+)[.)]\s+(.*)$", s)
        if m:
            numbered(doc, m.group(1), m.group(2)); i += 1; continue

        buf = [s]; i += 1
        while i < n:
            nx = lines[i].strip()
            if (not nx or nx.startswith(("#", ">", "|", "```", "- ", "* "))
                    or re.match(r"^\d+[.)]\s", nx) or re.fullmatch(r"-{3,}", nx)):
                break
            buf.append(nx); i += 1
        emit_inline(para(doc, after=6), " ".join(buf))


def count_headings(text):
    return sum(1 for l in text.split("\n") if re.match(r"^#{2,4}\s+\S", l.strip()))


# ---------------------------------------------------------------- تبدیل‌ها
def convert_md(src, dst):
    text = io.open(src, encoding="utf-8").read()
    m = re.search(r"^#\s+(.*)$", text, re.M)
    title = re.sub(r"\*\*(.+?)\*\*", r"\1", m.group(1).strip()) if m else \
        os.path.splitext(os.path.basename(src))[0]
    sub = ""
    if m:
        for ln in text[m.end():].lstrip("\n").split("\n")[:4]:
            if ln.strip().startswith("**"):
                sub = re.sub(r"\*\*", "", ln.strip()); break
    ncols = max([len(split_row(l)) for l in text.split("\n") if l.strip().startswith("|")] or [0])
    doc = new_doc(landscape=(ncols >= 8))
    full = page_full(doc)
    cover(doc, title, sub, full)
    if count_headings(text) >= 4:
        add_toc(doc, full)
        enable_update_fields(doc)
    render_md(doc, text, full)
    add_footer(doc, title)
    doc.save(dst)


def add_records(doc, headers, data, full):
    for ri, row in enumerate(data):
        row = list(row) + [""] * (len(headers) - len(row))
        label = " — ".join([c for c in row[:2] if c.strip()][:2]) or f"ردیف {ri + 1}"
        heading(doc, 3, label)
        pairs = [(headers[i], row[i]) for i in range(len(headers))
                 if row[i].strip() and row[i].strip() != "—"]
        if not pairs:
            continue
        tbl = doc.add_table(rows=0, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style = "Table Grid"
        _table_layout(tbl, [int(full * 0.26), full - int(full * 0.26)])
        for pi, (k, v) in enumerate(pairs):
            cells = tbl.add_row().cells
            _cell_shade(cells[0], SH_BRANDL)
            if pi % 2 == 1:
                _cell_shade(cells[1], SH_ZEBRA)
            for ci, (txt, bold) in enumerate(((k, True), (cell_value(v), False))):
                c = cells[ci]; c.vertical_alignment = 1
                p = c.paragraphs[0]; _rtl_para(p)
                set_align(p, "right")
                p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.line_spacing = 1.3
                emit_inline(p, txt, size=11.5, color=SLATE if bold else INK, base_bold=bold)
        _table_borders(tbl)
        _spacer(doc, 8)


def convert_csv(src, dst):
    with io.open(src, encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    headers, data = (rows[0], rows[1:]) if rows else ([], [])
    nc = len(headers)
    cells = [c for r in data for c in r]
    avg = sum(len(c) for c in cells) / max(1, len(cells))
    as_records = nc >= 10 or (nc >= 8 and avg > 20)

    doc = new_doc(landscape=(not as_records and nc >= 6))
    full = page_full(doc)
    title = os.path.splitext(os.path.basename(src))[0].replace("_", " ")
    mode = "چیدمان کارتی" if as_records else "جدول داده"
    cover(doc, title,
          f"{mode} · {fa_digits(len(data))} ردیف × {fa_digits(nc)} ستون", full)
    if as_records:
        if len(data) >= 4:
            add_toc(doc, full, levels="3-3")
            enable_update_fields(doc)
        add_records(doc, headers, data, full)
    else:
        add_table(doc, headers, data, full)
    add_footer(doc, title)
    doc.save(dst)


def convert_path(src, dst=None, clobber=True):
    """تبدیل یک فایل md یا csv. خروجی کنار فایل ورودی ساخته می‌شود.

    خروجی هم‌نام موجود بازنویسی می‌شود مگر clobber=False. بازنویسی با پیام هشدار
    گزارش می‌شود تا بی‌سروصدا اتفاق نیفتد."""
    base, ext = os.path.splitext(src)
    dst = dst or base + ".docx"
    if os.path.exists(dst):
        if not clobber:
            print("رد شد (موجود است):", dst)
            return None
        print("هشدار — بازنویسی:", dst)
    (convert_md if ext.lower() == ".md" else convert_csv)(src, dst)
    return dst


def main(argv=None):
    global ROOT, KICKER, AS_OF
    import argparse
    ap = argparse.ArgumentParser(description="تبدیل Markdown/CSV به Word فارسی")
    ap.add_argument("--root", help="پوشه ورودی (بازگشتی)")
    ap.add_argument("--file", help="یک فایل md یا csv")
    ap.add_argument("-o", "--out", help="مسیر خروجی (فقط با ‎--file)")
    ap.add_argument("--kicker", default="", help="نوار بالای کاور")
    ap.add_argument("--date", default="", help="تاریخ پانویس")
    ap.add_argument("--skip", default="_build", help="پوشه‌های نادیده‌گرفته‌شده (با کاما)")
    ap.add_argument("--no-clobber", action="store_true",
                    help="اگر خروجی هم‌نام وجود دارد، بازنویسی نکن و رد شو")
    a = ap.parse_args(argv)
    KICKER, AS_OF = a.kicker, a.date
    clobber = not a.no_clobber
    if a.file:
        out = convert_path(a.file, a.out, clobber)
        if out:
            print("ساخته شد:", out)
        return [out] if out else []
    if not a.root:
        ap.error("یکی از ‎--root یا ‎--file لازم است")
    ROOT = os.path.abspath(a.root)
    skip = {x.strip() for x in a.skip.split(",") if x.strip()}
    made = []
    for dirpath, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in skip]
        for fn in sorted(files):
            base, ext = os.path.splitext(fn)
            if ext.lower() not in (".md", ".csv"):
                continue
            src = os.path.join(dirpath, fn)
            dst = os.path.join(dirpath, base + ".docx")
            rel = os.path.relpath(src, ROOT).replace("\\", "/")
            try:
                if convert_path(src, dst, clobber):
                    made.append(rel)
            except Exception as e:
                print("خطا:", rel, "->", repr(e))
    print(f"ساخته شد: {fa_digits(len(made))} فایل Word")
    return made


if __name__ == "__main__":
    main()
