# persian-docx

**An Agent Skill for producing Word documents in Persian/Farsi that actually render correctly.**

[![tests](https://github.com/sadeghr64/persian-docx/actions/workflows/test.yml/badge.svg)](https://github.com/sadeghr64/persian-docx/actions/workflows/test.yml)
[![lint](https://github.com/sadeghr64/persian-docx/actions/workflows/lint.yml/badge.svg)](https://github.com/sadeghr64/persian-docx/actions/workflows/lint.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-informational.svg)](LICENSE)

<div dir="rtl">

**یک Agent Skill برای تولید اسناد Word فارسی که واقعاً درست رندر می‌شوند.**

</div>

---

## Why this exists

`python-docx` fails *silently* on Persian text. The file saves fine, re-reading it with
`python-docx` shows correct values — and then Word renders the document left-aligned, with
dates reversed, thousands separators that look like zeros, and table columns starting on the
wrong side.

Every rule in this skill came from a failure observed in real Word output, not from
documentation:

- `font.size` sets only `w:sz` — it has **no effect** on Persian text. `w:szCs` must be set by hand.
- RTL is **three layers** (`w:bidi` on the section, `w:bidi` on the paragraph, `w:rtl` on the
  run) plus `w:bidiVisual` on tables. Each missing layer breaks something different.
- In an RTL paragraph, `w:jc val="right"` renders **left**-aligned. `python-docx`'s
  `alignment = RIGHT` writes exactly that.
- Dates and numbers reverse unless wrapped in LRM (U+200E). RLM does *not* fix them — tested
  and failed.
- The B-family fonts carry only 288 glyphs. `٬` (U+066C) reads as a zero, so `۱٬۷۸۰` is read
  as `۱۰۷۸۰` — one extra digit, in a financial figure.

The skill also records **negative results** — things that were tested and did not work — so the
same experiment isn't repeated. See [`CHANGELOG.md`](CHANGELOG.md).

<div dir="rtl">

## چرا این اسکیل وجود دارد

`python-docx` روی متن فارسی **بی‌سروصدا** شکست می‌خورد: فایل سالم ذخیره می‌شود، بازخوانی با
`python-docx` مقادیر درست نشان می‌دهد، اما Word سند را چپ‌چین رندر می‌کند، تاریخ‌ها وارونه
می‌شوند، جداکنندهٔ هزارگان شبیه صفر دیده می‌شود و ستون اول جدول سمت اشتباه می‌افتد.

هر قاعده در این اسکیل از یک خطای مشاهده‌شده در خروجی واقعی Word به‌دست آمده، نه از مستندات.
نتایج **منفی** (چیزهایی که آزموده شدند و جواب ندادند) هم ثبت شده‌اند تا آزمایش تکراری نشود.

</div>

---

## Persian prose and editorial review

For authoring or prose editing, the skill includes a general-purpose
[Persian prose guide](persian-docx/references/persian-prose.md): audience-appropriate tone,
concise sentences, removal of redundant emphasis, precise vocabulary, consistent terminology,
source fidelity and an independent editorial review. Meaningful restrictions, uncertainty,
negation and exceptions must survive editing; protected quotations are not rewritten for style.
For formatting-only requests, preserve the supplied wording.

For formal reports, the [report workflow](persian-docx/references/report-workflow.md)
starts with the intended reader and decision, keeps project inputs separate from published
sources and incorporates pre-issue feedback into the same report. Design the opening, body and
conclusion with the [report structure guide](persian-docx/references/report-structure.md), then
review the prose separately. The [business plan and feasibility guide](persian-docx/references/report-business-appraisal.md)
matches analytical tools to the decision. Use the [section glossaries](persian-docx/references/report-terminology.md)
for technical, financial and legal writing, and see the [before-and-after examples](persian-docx/references/report-examples.md).
For awkward or formulaic drafts, the [Persian rewrite method](persian-docx/references/persian-rewrite.md)
rebuilds paragraphs from their claims and evidence, then checks readability and meaning separately.
It also records which optional Persian text tools to use for typography and why they are not required.

For authoring or rewriting a report, start with the [built-in report delivery workflow](persian-docx/references/report-delivery.md).
The same agent plans, edits and rereads the report. The local `report_delivery.py` command
checks observable issues and requires an editorial record for the exact draft before exporting
Word. It needs no separate proofreading service or model. A successful check does not certify
the prose or verify evidence. Use `fa_docx.py` for conversion only.

<div dir="rtl">

### نگارش و ویرایش نثر

راهنمای نثر برای نامه، گزارش، راهنما، صورت‌جلسه و دیگر اسناد فارسی کاربرد دارد. لحن با نوع
سند تنظیم می‌شود؛ عبارت زائد حذف می‌شود، اما شرط، استثنا، نفی و میزان قطعیت حفظ می‌شود.
واژه‌گزینی بر دقت و تناسب با مخاطب استوار است، نه دشوارنویسی یا تغییر اصطلاحات ثابت.
بازبینی نثر از کنترل صفحه‌آرایی جداست و اسکریپت تبدیل، ویرایش معنایی انجام نمی‌دهد.

در گزارش رسمی، متن برای خواننده‌ی نهایی و تصمیم مورد انتظار نوشته می‌شود. هر فصل تخصصی
واژه‌نامه‌ی حوزه‌ی خودش را دارد؛ داده‌ی داخلی مبنای تحلیل است و در کتاب‌نامه نمی‌آید.
اصلاحات پیش از صدور در همان گزارش ادغام می‌شوند. رسم‌الخط گزارش، از جمله «نقشه‌ی» با
نیم‌فاصله، در بازخوانی نثر و سپس در رندر Word بررسی می‌شود.

ابتدا بر اساس [نوع گزارش](persian-docx/references/report-structure.md)، ساختار، محتوای هر بخش و لحن آن
را طرح‌ریزی کنید. سپس متن را از نظر روانی و دقت بازنویسی کنید. [راهنمای طرح کسب‌وکار و امکان‌سنجی](persian-docx/references/report-business-appraisal.md)
به انتخاب روش تحلیل بر اساس پرسش گزارش کمک می‌کند.

برای تألیف یا بازنویسی گزارش، [مسیر داخلی تولید و تحویل](persian-docx/references/report-delivery.md)
را اجرا کنید. فایل Word از همان پیش‌نویسی ساخته می‌شود که بازخوانی و ثبت شده است. برای تبدیل
صرف، مبدل مستقیم مناسب است؛ متن را بازنویسی نمی‌کند.

</div>

## Repository layout

| Path | Purpose |
|---|---|
| `persian-docx/SKILL.md` | The skill itself — 12 sections: the six traps, fonts and sizes, three-layer RTL, alignment, bidi/LRM rules, glyph coverage, Persian typography, auto TOC, tables, the mandatory verification workflow, page setup, lessons index |
| `persian-docx/references/persian-prose.md` | General Persian prose guidance, neutral examples and an editorial review checklist |
| `persian-docx/references/report-workflow.md` | Audience, evidence, pre-issue revisions and the report review workflow |
| `persian-docx/references/report-delivery.md` | Internal authoring and editorial checks before export |
| `persian-docx/references/report-structure.md` | Opening, body, conclusion and section tone across ten report types |
| `persian-docx/references/report-business-appraisal.md` | Business plans, feasibility studies and selection of appraisal tools |
| `persian-docx/references/report-source-review.md` | Published samples and guidance used to design the report structures |
| `persian-docx/references/report-terminology.md` | Section-specific terminology selection and starter glossaries |
| `persian-docx/references/report-examples.md` | Before-and-after report examples and semantic checks |
| `persian-docx/references/persian-rewrite.md` | Meaning-first Persian rewriting and review criteria |
| `persian-docx/scripts/report_delivery.py` | Prepare, check and export the exact reviewed report draft locally |
| `persian-docx/scripts/fa_docx.py` | Reference converter — Markdown/CSV → Persian Word layout and typography; no semantic rewriting. Standard library plus `python-docx`, nothing else |
| `examples/sample.md` → `sample.docx` | End-to-end sample: headings, tables, code blocks, callouts, lists, links, dates, numbers |
| `tests/test_rules.py` | Asserts the generated OOXML obeys the skill. Runs in CI on Python 3.9–3.13 |

---

## Quick start

```bash
pip install -r requirements.txt

# Convert existing text without editorial review
python persian-docx/scripts/fa_docx.py --file report.md

# a whole tree of .md and .csv files
python persian-docx/scripts/fa_docx.py --root ./docs \
       --kicker "نام سازمان" --date "۱۴۰۵/۰۶/۰۱"
```

Output is written next to each input as `<name>.docx`. An existing file of the same name is
overwritten, with a warning printed; pass `--no-clobber` to skip instead. See
[`SECURITY.md`](SECURITY.md) for exactly what the script touches on disk.

---

## Installation as a skill

### Claude Code

```bash
git clone https://github.com/sadeghr64/persian-docx.git
mkdir -p ~/.claude/skills/persian-docx
cp -R persian-docx/persian-docx/. ~/.claude/skills/persian-docx/
```

Then `/persian-docx` in any session. For a single project, put it in the project's
`.claude/skills/` instead.

### Claude (web / desktop / Cowork)

```bash
cd persian-docx && zip -r persian-docx.zip persian-docx
```

Then **Customize → Skills** in the Claude app, and add the zip. The skill folder must be the
**root** of the archive, not nested inside another folder.

<div dir="rtl">

**توجه:** اسکیل‌های Claude Code فقط روی همان دستگاه در دسترس‌اند. برای اینکه در Cowork هم کار
کند، باید همان بستهٔ zip را در **Customize → Skills** آپلود کنی.

</div>

### Other agents and tools

A repository URL is a pointer; it does not load or activate the skill in another assistant.
Choose the route based on what the destination can actually read and run:

| Destination | What to provide |
|---|---|
| An agent that can load skills and read local files | Install the entire inner `persian-docx/` folder, including `SKILL.md`, `WRITE-PERSIAN.md`, `references/` and `scripts/`. Replace an older installed copy when updating. |
| A chat without skill installation or shell access | Attach [`WRITE-PERSIAN.md`](persian-docx/WRITE-PERSIAN.md), or paste its contents before the brief. It includes the complete prose workflow for chat and does not depend on scripts or linked files. |
| A chat with browser access only | Point it directly to the [standalone guide](persian-docx/WRITE-PERSIAN.md) and explicitly ask it to read the full file and use it. If it cannot read the full text, attach or paste it instead. |

<div dir="rtl">

برای استفاده در چت، این متن را همراه فایل راهنما بفرستید:

> «فایل پیوست `WRITE-PERSIAN.md` را کامل بخوان و به‌عنوان دستور نگارش اجرا کن. گزارش را برای [مخاطب] و با هدف [هدف] بنویس. پیش از تحویل، روانی نثر و وفاداری به شواهد را جداگانه بررسی کن. بیرون گزارش، شناسه‌ی راهنما، کنترل‌های واقعاً انجام‌شده و یک نمونه از اصلاح پیش‌نویس یا متن مبنا را بنویس.»

اگر چت فقط متن تولید می‌کند، این مسیر بازخوانی نثر را پوشش می‌دهد؛ برای ساخت و کنترل Word از نصب کامل در محیط دارای ابزار استفاده کنید. اعلام شناسه به‌تنهایی گواه اجرای دستورها نیست.

</div>

---

## Verification — the part people skip

> **It is not done until you have looked at the rendered output with your own eyes.**

Re-reading the `.docx` with `python-docx` is not verification: every trap survives a re-read
intact and only appears in the render. §9 of the skill is a mandatory render → image → checklist
loop, with a machine-checkable XML pass first.

```bash
# Windows + Word gives the reference render (and updates the TOC field).
# Anywhere else, LibreOffice is enough for layout, direction, fonts, tables and numbers:
soffice --headless --convert-to pdf report.docx
```

The committed sample is checked automatically:

| Assertion | Result |
|---|---|
| `w:jc val="right"` on any paragraph — would render **left**-aligned | 0 |
| `w:szCs` on runs · `w:bidi` on section and paragraphs · `w:rtl` on runs · `w:bidiVisual` on tables | present |
| Dates wrapped in LRM, not reversed | ✅ |
| `٬` / `٫` surviving anywhere — including inside code blocks | 0 |
| Empty paragraph carrying a page break after the TOC (produces a blank page) | 0 |
| Arabic `ي` `ك` `ى` left unnormalised | 0 |

```bash
python tests/test_rules.py     # or: pytest tests/ -q
```

---

## Requirements

- Python 3.9+
- `python-docx` — required
- `pymupdf`, `fonttools` — optional, for the verification workflow
- Persian B-family fonts installed (B Nazanin, B Titr) for a faithful render
- Microsoft Word for the full workflow; LibreOffice covers everything except field updates
  (the TOC stays empty and page numbers keep Latin digits)

---

## Contributing

Rules enter this skill only with evidence — a reproducing input and a rendered image.
**Negative results count as contributions.** See [`CONTRIBUTING.md`](CONTRIBUTING.md) and
[`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md).

## License

MIT — see [LICENSE](LICENSE).
