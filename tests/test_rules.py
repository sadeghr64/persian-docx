# -*- coding: utf-8 -*-
"""Assert that generated output obeys the rules in SKILL.md.

These checks read the raw OOXML, not python-docx's object model, because every trap in §1
survives a python-docx re-read intact and only shows up in the XML or the render.

    python -m pytest tests/ -q
    python tests/test_rules.py        # runs without pytest too
"""
import os
import re
import sys
import zipfile
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "persian-docx", "scripts"))

import fa_docx  # noqa: E402

SAMPLE = os.path.join(ROOT, "examples", "sample.md")

U066C = "٬"   # Arabic thousands separator — reads as a zero in the B fonts
U066B = "٫"   # Arabic decimal separator
LRM = "‎"


def build():
    """Convert the sample into a temp directory and return its document.xml."""
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, "sample.docx")
    fa_docx.convert_path(SAMPLE, out)
    with zipfile.ZipFile(out) as z:
        return {
            "document": z.read("word/document.xml").decode("utf-8"),
            "styles": z.read("word/styles.xml").decode("utf-8"),
            "settings": z.read("word/settings.xml").decode("utf-8"),
        }


XML = build()


def count(pattern, part="document"):
    return len(re.findall(pattern, XML[part]))


# -- §3-1 alignment ---------------------------------------------------------------
def test_no_visually_left_aligned_paragraph():
    """w:jc val="right" inside an RTL paragraph renders LEFT-aligned. There must be none."""
    assert count(r'<w:jc w:val="right"') == 0


def test_right_alignment_is_written_as_left():
    assert count(r'<w:jc w:val="left"') > 0


# -- §2 font size -----------------------------------------------------------------
def test_complex_script_size_is_set():
    """Without w:szCs the size never reaches Persian text."""
    assert count(r"<w:szCs") > 0


# -- §3 three-layer RTL -----------------------------------------------------------
def test_rtl_layers_present():
    assert count(r"<w:bidi/>|<w:bidi w:val=\"1\"") > 0      # section + paragraphs
    assert count(r'<w:rtl w:val="1"') > 0                    # runs
    assert count(r"<w:bidiVisual") > 0                       # tables


def test_styles_are_rtl_too():
    assert count(r"<w:bidi/>", "styles") > 0


# -- §4 bidi numbers --------------------------------------------------------------
def test_numbers_are_lrm_wrapped():
    assert XML["document"].count(LRM) > 0
    # a date must appear wrapped, not bare
    assert re.search(LRM + r"[۰-۹]{4}/[۰-۹]{2}/[۰-۹]{2}" + LRM, XML["document"])


# -- §5 glyph coverage ------------------------------------------------------------
def test_missing_glyphs_absent_everywhere_including_code_blocks():
    """U+066C reads as a zero in the B fonts; Consolas lacks it too, so code blocks count."""
    assert XML["document"].count(U066C) == 0
    assert XML["document"].count(U066B) == 0


# -- §7 / §7-1 table of contents --------------------------------------------------
def test_toc_field_and_update_flag():
    assert "instrText" in XML["document"]
    assert "updateFields" in XML["settings"]


def test_no_empty_page_break_paragraph_after_toc():
    """An empty paragraph carrying a page break produces a fully blank page."""
    assert count(r'<w:br w:type="page"') == 0
    assert count(r"pageBreakBefore") > 0


# -- §6 typography ----------------------------------------------------------------
def test_arabic_characters_normalised():
    for ch in ("ي", "ك", "ى"):   # ي ك ى
        assert XML["document"].count(ch) == 0


if __name__ == "__main__":
    failed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print("PASS", name)
            except AssertionError as e:
                failed += 1
                print("FAIL", name, e)
    print("\n%d failed" % failed)
    sys.exit(1 if failed else 0)
