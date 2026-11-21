import json
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


SCRIPT = Path(__file__).resolve().parents[1] / "persian-docx" / "scripts" / "report_delivery.py"


def invoke(*args):
    return subprocess.run(
        [sys.executable, "-X", "utf8", str(SCRIPT), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8",
    )


def record(review):
    data = json.loads(review.read_text(encoding="utf-8"))
    data["reader"] = "مسئول تولید"
    data["purpose"] = "انتخاب خبر برای انتشار"
    data["passes"] = {
        "structure": "خبرهای داخل بازه پیش از موارد در انتظار تأیید قرار گرفت.",
        "evidence": "عدد ۵۴۲ به مقدار بازیابی‌شده اشاره دارد.",
        "prose": "فعل اصلی در انتهای جمله قرار گرفت.",
        "final_read": "خلاصه و جدول یک دامنه را پوشش می‌دهند.",
    }
    for section in data["sections"]:
        section["finding"] = f"کنترل نقش بخش {section['heading']}"
    for f in data["findings"]:
        data["keep"][f["id"]] = "این صورت در نقل عین لازم است."
    review.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def test_word_build_requires_review_of_this_version(tmp_path):
    src, review, out = [tmp_path / p for p in ("draft.md", "review.json", "out.docx")]
    draft = """# گزارش دوره‌ای

**بازه‌ی آزمون**

## جمع‌بندی
پایش این دوره برای انتخاب خبر انجام شد.

### نتیجه‌ی مالی
ارزش فعلی خالص^[Net Present Value (NPV)] در جدول آمده است.

## شواهد
- گزارش داخلی، وضعیت را ثبت کرد.
- عدد کنار دامنه‌ی آن آمده است.
ارزش فعلی خالص^[Net Present Value (NPV)] در پایان دوباره آمده است.

## اقدام
خبر اول در اولویت قرار گیرد.

## منابع
مرجع: گزارش منتشرشده.
"""
    src.write_text(draft, encoding="utf-8")
    assert invoke("prepare", src, "--review", review).returncode == 0
    assert invoke("build", src, "--review", review, "--out", out).returncode != 0
    assert not out.exists()
    record(review)
    src.write_text(draft.replace("خبر اول", "خبر دوم"), encoding="utf-8")
    assert invoke("build", src, "--review", review, "--out", out).returncode != 0
    assert not out.exists()
    assert invoke("prepare", src, "--review", review, "--refresh").returncode == 0
    record(review)
    assert invoke("build", src, "--review", review, "--out", out).returncode == 0
    doc = Document(out)
    body = doc.element.xpath("//w:t")
    written = " ".join(t.text for t in body if t.text)
    assert written.count("بازه‌ی آزمون") == 1
    assert "Ctrl+A" not in written and "^[" not in written and "خبر دوم" in written
    assert [p.style.name for p in doc.paragraphs if p.text == "جمع‌بندی"] == ["Heading 1"]
    assert [p.style.name for p in doc.paragraphs if p.text == "نتیجه‌ی مالی"] == ["Heading 2"]
    assert doc.styles["Heading 1"].element.xpath(".//w:rFonts")[0].get(qn("w:cs")) == "B Titr"
    assert doc.paragraphs[0].style.name == "Title"
    assert len(doc.element.xpath("//w:pPr/w:numPr")) == 2
    assert len(doc.element.xpath("//w:footnoteReference")) == 1
    assert doc.settings.element.xpath(".//w:numRestart")[0].get(qn("w:val")) == "eachPage"
    assert doc.element.xpath('//w:instrText[contains(text()," TOC ")]')
    assert invoke("build", src, "--review", review, "--out", out).returncode != 0


def test_build_cannot_waive_leaked_source_or_broken_table(tmp_path):
    src, review = tmp_path / "draft.md", tmp_path / "r.json"
    src.write_text(
        "# گزارش\n## دامنه\nمنبع: [داده](file:///C:/work/private.docx)\n"
        "مرجع: [گزارش](https://example.org/released.md)\n"
        "برای به‌روزرسانی فهرست: Ctrl+A سپس F9\n"
        "| شاخص | ارزش |\n| --- | --- |\n| ناقص |\n",
        encoding="utf-8",
    )
    prepared = invoke("prepare", src, "--review", review)
    assert prepared.returncode == 0
    data = record(review)
    assert {x["id"] for x in data["findings"]} == {
        "local-path:3", "internal-citation:3", "editorial-residue:5", "table-width:8",
    }
    checked = invoke("check", src, "--review", review)
    assert checked.returncode == 1
    assert all(code in checked.stderr for code in ("local-path", "editorial-residue", "table-width"))


def test_conversion_only_preserves_bold_opening_sentence(tmp_path):
    import fa_docx

    p = tmp_path / "convert.docx"
    fa_docx.convert_md_text("# عنوان\n\n**نتیجه:** این بند بخشی از متن است.\n", p)
    d = Document(p)
    assert d.paragraphs[-1].text == "نتیجه: این بند بخشی از متن است."
