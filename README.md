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

## Repository layout

| Path | Purpose |
|---|---|
| `persian-docx/SKILL.md` | The skill itself — 12 sections: the six traps, fonts and sizes, three-layer RTL, alignment, bidi/LRM rules, glyph coverage, Persian typography, auto TOC, tables, the mandatory verification workflow, page setup, lessons index |
| `persian-docx/scripts/fa_docx.py` | Reference implementation — Markdown/CSV → Persian Word, applying every rule. Standard library plus `python-docx`, nothing else |
| `examples/sample.md` → `sample.docx` | End-to-end sample: headings, tables, code blocks, callouts, lists, links, dates, numbers |
| `tests/test_rules.py` | Asserts the generated OOXML obeys the skill. Runs in CI on Python 3.9–3.13 |

---

## Quick start

```bash
pip install -r requirements.txt

# one file
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
cp -r persian-docx/persian-docx ~/.claude/skills/persian-docx
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

`SKILL.md` follows the Agent Skills open format — plain Markdown with YAML frontmatter
(`name`, `description`). Nothing in the body depends on any particular vendor, so it is
portable to anything that reads instruction files:

| Tool | How |
|---|---|
| Cursor | Copy `SKILL.md` into `.cursor/rules/persian-docx.mdc` |
| GitHub Copilot | Copy the body into `.github/copilot-instructions.md` |
| Gemini CLI / Codex / OpenCode / Cline | Point the agent at `SKILL.md`, or paste it into the project instruction file |
| Any LLM chat | Paste the body as context before asking for a Persian Word document |
| No agent at all | Read it as a reference and use `fa_docx.py` directly |

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
