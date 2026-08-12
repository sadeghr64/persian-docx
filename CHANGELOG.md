# Changelog

All notable changes to this skill are documented here. Format loosely follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[Semantic Versioning](https://semver.org/).

Each entry records a rule and the observation behind it. **Rules that were tested and did not
work are recorded too** — a negative result saves the next person the same experiment.

## [1.0.0] — 2026-08-12 — first public release

### Added — the rule set (`SKILL.md`)

- **Font size (§2).** `font.size` sets only `w:sz`, which has no effect on Persian text.
  `w:szCs` must be written by hand, in half-points.
- **B Titr (§2).** The face ships Bold only; without `bold=True` Word substitutes another font.
- **Three-layer RTL (§3).** `w:bidi` on the section, `w:bidi` on the paragraph, `w:rtl` on the
  run, `w:bidiVisual` on tables. Each missing layer produces a different failure.
- **`w:bidi` ordering (§3).** Must be inserted once, and before `w:jc`, per the schema order.
- **Alignment (§3-1).** In an RTL paragraph `w:jc val="right"` renders **left**-aligned and
  `left` renders right-aligned; `python-docx` writes `right` for `alignment = RIGHT`. The same
  inversion applies to `w:ind`.
- **Number and date order (§4).** Wrapping numeric tokens in LRM (U+200E) preserves their order.
- **Glyph coverage (§5).** The B font family carries 288 glyphs. `٬` (U+066C) reads as a zero,
  which silently adds a digit to any figure. Replace with `،`, and `٫` with `/`.
- **Code blocks (§5).** `Consolas` lacks those glyphs too, so the replacement must be applied
  to code lines as well as prose.
- **Persian typography (§6).** Arabic→Persian character mapping, conservative ZWNJ rules,
  punctuation spacing, Persian quotation marks.
- **Table of contents (§7).** TOC field plus `w:updateFields` so Word populates it on open.
- **Blank page after the TOC (§7-1).** An empty paragraph carrying `WD_BREAK.PAGE` produces a
  fully blank page; `pageBreakBefore` on a 1 pt paragraph does not.
- **Tables (§8).** Column widths from content length, landscape past six columns, card layout
  past ten, repeated header rows.
- **Verification workflow (§9).** Render → image → visual checklist, plus XML assertions.
  Mandatory: re-reading the `.docx` is not verification.
- **Page setup (§11).** A4, right-aligned rather than justified.

### Added — tested and failed

- `w:rtl val="0"` on a run does **not** fix reversed dates. Only LRM/RLM control characters do.
- RLM does not fix reversed dates either — it is for isolating Latin fragments.
- Latin vs Persian digit *shape* cannot be controlled from inside the document. Four
  combinations of `w:rtl` and `w:lang` were tested; all four rendered Persian digits. Word's
  `Numeral` setting is application-level, so the same file renders differently on two machines.
  The only guaranteed way to keep Latin digits in an identifier is to insert it as an image.
- PDF text extraction is not a valid measurement for RTL: it drops the leading word of a line
  and fragments coordinates. The rendered image is the evidence.

### Added — reference implementation (`scripts/fa_docx.py`)

Markdown and CSV to Persian Word, applying every rule above: four heading levels, tables,
bullets, numbered lists, checklists, callout boxes, code blocks, hyperlinks, cover page,
automatic TOC, footer with page numbers.

### Notes for this release

- The skill body was generalised for public use: examples tied to specific document sets were
  replaced with generic ones.
- `--no-clobber` was added, and overwrites are now reported rather than silent.
- The verification workflow gained a LibreOffice path so it is no longer Windows-only.
- `ruff.toml` selects its lint rules explicitly instead of inheriting ruff's default set,
  which widened from roughly 40 rules to 400 in ruff 0.16. The CI job pins the ruff version
  to match, so the badge reflects the code rather than the age of the toolchain.
- `.gitattributes` normalises line endings to LF and marks `.docx` and other binaries as
  binary, so contributors on different platforms don't produce whole-file diffs.
