"""Check the local editorial record before exporting an authored Persian report.

This tool detects a few observable defects and binds the review to the exact draft.
The authoring agent, not this program, judges the meaning and quality of the prose.
"""

import argparse
import hashlib
import json
import re
import sys
import unicodedata
from pathlib import Path


PASSES = ("structure", "evidence", "prose", "final_read")
HEAD = re.compile(r"^(#{1,6})[ \t]+(.+?)\s*$")
SEPARATOR = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$")
URL = re.compile(r"https?://[^\s>)]+")
NOTE = re.compile(r"\^\[([^\]\r\n]*)\]")
FORMAT_CONTROL = str.maketrans("", "", "\u200e\u200f\u061c")
COLLOQUIAL = re.compile(
    r"جورش؟ (?:درنمی‌آید|درنمیاد)|جور نبود|مال (?:این|همین) هفته|این‌?جا جایش نیست"
    r"|قاطی نکنید|حرف زده|لو رفت(?:ن)?|پنجره‌ی فرصت باز بود"
)
VAGUE_TITLE = re.compile(r"^(?:این (?:پرونده|گزارش) برای چیست|سایر موارد|نکات مهم)\s*[؟?]?$")
NEGATION = re.compile(
    r"(?<!\w)(?:نیست(?:ند|م|ید)?|نبود(?:ند|م)?|نشد(?:ه|ند)?|نشده‌اند|"
    r"نکن(?:د|ند|ید)|نگویید|نیاورید|نمی[‌ ]?شود|نخواهد شد)(?!\w)"
)
META = re.compile(r"(?:برای به[‌ ]روزرسانی فهرست\s*:|طبق درخواست شما|توضیح برای کارفرما\s*:)")
LOCAL_PATH = re.compile(r"(?i)(?:\bfile:/|(?<![\w/])[A-Z]:[\\/][^\s]+|(?:^|\s)\\\\[^\s]+|/(?:mnt|Users|home|tmp)/)")
INTERNAL_CITE = re.compile(r"(?i)^(?:[-*]\s+)?(?:منبع|مأخذ|مرجع)\s*[:：].*\.(?:xlsx|docx|csv|md)\b")
TOOL_TOKEN = re.compile(r"\ue200cite\ue202|【\d+†|\[TBD\]|\[TODO\]", re.I)


def plain(text):
    return unicodedata.normalize("NFC", text.translate(FORMAT_CONTROL)).strip()


def scan(text):
    sections, findings = [], []
    seen = {}
    previous = 0

    def flag(line, code, message, blocking=False):
        ident = f"{code}:{line}"
        if not any(x["id"] == ident for x in findings):
            findings.append({"id": ident, "line": line, "message": message, "blocking": blocking})

    lines = text.splitlines()
    fenced = None
    first = next((i for i, line in enumerate(lines, 1) if line.strip()), 1)
    table_cols = None
    for i, line in enumerate(lines, 1):
        raw = plain(line)
        # Safety and leakage checks also cover quoted passages, code blocks and link targets.
        without_web_urls = URL.sub("", raw)
        if LOCAL_PATH.search(without_web_urls):
            flag(i, "local-path", "مسیر یا پیوند داخلی در متن سند", True)
        if INTERNAL_CITE.search(without_web_urls):
            flag(i, "internal-citation", "استناد به فایل کاری به‌جای مأخذ منتشرشده", True)
        if TOOL_TOKEN.search(raw):
            flag(i, "draft-token", "نشانگر ناتمام یا توکن ابزار", True)
        if raw.startswith(("```", "~~~")):
            fence = raw[:3]
            if fenced is None:
                fenced = fence, i
                if fence == "~~~":
                    flag(i, "unsupported-fence", "برای بلوک کد از سه بک‌تیک استفاده کن", True)
            elif fenced[0] == fence:
                fenced = None
            continue
        if fenced is not None:
            continue
        m = HEAD.fullmatch(raw)
        if m:
            level, heading = len(m[1]), m[2]
            if (level == 1 and i != first) or level > 5 or level > previous + 1:
                flag(i, "heading-order", "یک عنوان سند و سطح‌های پیوسته تا چهار سطح عنوان بخش لازم است", True)
            previous = level
            if level > 1:
                sections.append({"line": i, "heading": heading, "finding": ""})
            if VAGUE_TITLE.fullmatch(heading) or heading.endswith(("؟", "?")):
                flag(i, "heading-wording", "عنوان را پس از خواندن بخش به عبارت دقیق تبدیل کن؛ پرسش هدفمند را توضیح بده")
            if "^[" in heading:
                flag(i, "heading-note", "پانویس اصطلاح را به اولین کاربرد در متن اصلی منتقل کن", True)
        if META.search(raw) and not raw.startswith(">"):
            flag(i, "editorial-residue", "گفت‌وگو یا دستور تولید سند در متن نهایی", True)
        if raw.startswith(("• ", "▪ ")):
            flag(i, "typed-list", "برای فهرست واقعی Word، نشانه‌ی Markdown به کار ببر", True)
        if "^[" in raw:
            notes = NOTE.findall(raw)
            if raw.count("^[") != len(notes) or any(not re.search("[A-Za-z]", n) for n in notes):
                flag(i, "term-note", "پانویس اصطلاح باید ^[English term] و تک‌سطر باشد", True)
        if raw.startswith("|"):
            cells = len(raw.strip("|").split("|"))
            if table_cols is None:
                if i == len(lines) or not SEPARATOR.fullmatch(plain(lines[i])):
                    flag(i, "table-format", "جدول بدون ردیف جداکننده", True)
                table_cols = cells
            elif cells != table_cols:
                flag(i, "table-width", "شمار سلول‌ها با سرستون‌های جدول برابر نیست", True)
        else:
            table_cols = None
        if raw and not m and not raw.startswith(("|", ">", "-", "*")) and len(raw) >= 32:
            key = " ".join(raw.split())
            if key in seen:
                flag(i, "duplicate", f"تکرار عین متن در سطر {seen[key]}")
            else:
                seen[key] = i
        if COLLOQUIAL.search(raw):
            flag(i, "register", "لحن محاوره‌ای یا ترجمه‌وار؛ تناسب آن را با مخاطب و نقش بخش بررسی کن")
        if len(NEGATION.findall(raw)) >= 2:
            flag(i, "stacked-negation", "چند فعل منفی؛ شرط یا محدودیت واقعی را نگه دار و بیان تکراری را بازنویسی کن")
    if fenced is not None:
        flag(fenced[1], "unclosed-code", "بلوک کد بدون پایان", True)
    if not lines or not re.match(r"^#[ \t]+\S", plain(lines[first - 1])):
        flag(first, "missing-title", "نخستین خط غیرخالی باید عنوان سند با یک # باشد", True)
    if not sections:
        # Brief notes may fit in a single unheaded body.
        sections.append({"line": first, "heading": "بدنه‌ی گزارش کوتاه", "finding": ""})
    return sections, findings


def packet(data, text):
    sections, findings = scan(text)
    return {
        "version": 1, "sha256": hashlib.sha256(data).hexdigest(),
        "reader": "", "purpose": "",
        "passes": dict.fromkeys(PASSES, ""),
        "sections": sections,
        "findings": findings,
        "keep": {},
    }


def has_text(value):
    return isinstance(value, str) and bool(value.strip())


def validate(current, review):
    if not isinstance(review, dict):
        return ["پرونده‌ی بازبینی باید یک شیء JSON باشد."]
    if review.get("version") != current["version"] or review.get("sha256") != current["sha256"]:
        return ["بازبینی برای این نسخه‌ی پیش‌نویس نیست. پرونده را تازه کن و نثر و معنا را دوباره بسنج."]
    errors = [f"{f['id']}: {f['message']}" for f in current["findings"] if f["blocking"]]
    for name in ("reader", "purpose"):
        if not has_text(review.get(name)):
            errors.append(f"{name}: مخاطب و هدف واقعی گزارش را بنویس.")
    passes = review.get("passes")
    for name in PASSES:
        if not isinstance(passes, dict) or not has_text(passes.get(name)):
            errors.append(f"passes.{name}: یک مشاهده‌ی مشخص از این بازخوانی بنویس.")
    sections = review.get("sections")
    if not isinstance(sections, list) or [
        (s.get("line"), s.get("heading")) if isinstance(s, dict) else None for s in sections
    ] != [(s["line"], s["heading"]) for s in current["sections"]]:
        errors.append("فهرست بخش‌ها باید با متن جاری هم‌خوان باشد.")
    elif any(not has_text(s.get("finding")) for s in sections):
        errors.append("برای همه‌ی بخش‌ها یک یادداشت مشخص درباره‌ی کارکرد و نتیجه‌ی بازخوانی ثبت کن.")
    keep = review.get("keep")
    for f in current["findings"]:
        if not f["blocking"] and (not isinstance(keep, dict) or not has_text(keep.get(f["id"]))):
            errors.append(f"{f['id']}: {f['message']}")
    return errors


def main(argv=None):
    cli = argparse.ArgumentParser(description=__doc__)
    subs = cli.add_subparsers(dest="action", required=True)
    for action in ("prepare", "check", "build"):
        p = subs.add_parser(action)
        p.add_argument("draft", type=Path)
        p.add_argument("--review", type=Path, required=True)
        if action == "prepare":
            p.add_argument("--refresh", action="store_true", help="بازنویسی فرم قدیمی با فرم خالی نسخه‌ی تازه")
        if action == "build":
            p.add_argument("--out", type=Path, required=True)
    a = cli.parse_args(argv)
    try:
        if a.draft.suffix.lower() != ".md" or a.review.suffix.lower() != ".json":
            raise ValueError("پیش‌نویس باید .md و پرونده‌ی بازبینی باید .json باشد.")
        data = a.draft.read_bytes()
        text = data.decode("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
        current = packet(data, text)
        if a.action == "prepare":
            with a.review.open("w" if a.refresh else "x", encoding="utf-8") as f:
                json.dump(current, f, indent=2, ensure_ascii=False)
                f.write("\n")
            for finding in current["findings"]:
                print(f"{finding['id']}: {finding['message']}")
            print(f"فرم بازبینی: {a.review}")
            return 0
        errors = validate(current, json.loads(a.review.read_text(encoding="utf-8-sig")))
        if errors:
            for e in errors:
                print(e, file=sys.stderr)
            return 1
        if a.action == "build":
            if a.out.suffix.lower() != ".docx":
                raise ValueError("پسوند خروجی باید .docx باشد.")
            from fa_docx import convert_md_text
            from tempfile import TemporaryFile
            with TemporaryFile() as tmp:
                convert_md_text(text, tmp, report=True)
                tmp.seek(0)
                with a.out.open("xb") as dest:
                    dest.write(tmp.read())
            print(f"Word ساخته شد؛ رندر و بازبینی دیداری را انجام بده: {a.out}")
        else:
            print("نسخه‌ی جاری با پرونده‌ی بازبینی هم‌خوان است. روانی متن را بازخوان می‌سنجد.")
        return 0
    except (ValueError, OSError) as exc:
        print(exc, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
