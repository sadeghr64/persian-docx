# Contributing

Thanks for wanting to improve this skill. The rules here were learned the hard way — from
Word output that looked fine on re-read and wrong on screen. Contributions are welcome on
exactly those terms: **evidence, not documentation**.

## The evidence standard

A rule enters `SKILL.md` only if you have *rendered* the document and looked at it. Re-reading
a `.docx` with `python-docx` is not evidence — every trap in §1 survives a re-read intact.

Before opening a PR that changes a rule, please include:

1. A minimal Markdown or CSV input that reproduces the problem.
2. The rendered output, before and after — a PNG at ≥200 dpi, or the page number in
   `examples/sample.docx` where it is visible.
3. Which Word version (or LibreOffice version) you observed it in. Behaviour differs.

## Negative results are contributions

If you tested something and it did **not** work, that belongs in the skill too. Knowing that
RLM does not fix reversed dates saves the next person the same experiment. Add it as a
"tested and failed" row, with the same evidence standard.

## Checklist for a rule change

- [ ] The rule is in the right section of `SKILL.md`
- [ ] A line was added to the visual checklist in §9
- [ ] `CHANGELOG.md` has an entry describing what changed and why
- [ ] `examples/sample.md` exercises the new case, if it can be exercised
- [ ] `python persian-docx/scripts/fa_docx.py --file examples/sample.md` still runs clean
- [ ] The XML assertions in §9 still pass on the regenerated sample

## Keep it general

This skill is published for everyone. Please do not add content tied to a specific
organisation, client, or document set: no internal document codes, project names, client
names, or counts from a private corpus. Describe the *class* of document instead
("a financial report", "a technical specification").

## Code style

`fa_docx.py` is deliberately a single dependency-light file. Please keep it that way:
standard library plus `python-docx`. No network access, no subprocess calls, no writing
outside the output path the user asked for.

## Language

`SKILL.md` is written in Persian because its audience is people producing Persian documents.
`README.md` is bilingual. Issues and PRs are welcome in either Persian or English.
