# Security

## What this project does

`persian-docx` is a documentation skill plus one Python script. It converts Markdown and CSV
files into Word `.docx` files. It is offline and self-contained:

- **No network access.** The script makes no HTTP requests and imports no networking module.
- **No subprocess execution.** It does not shell out, and it does not call `eval`, `exec`,
  or `pickle`.
- **No deletion.** It never removes files from disk.
- **Dependencies:** `python-docx` only. `pymupdf` and `fonttools` are optional and used only
  by the verification workflow, which you run yourself.

## What it writes

For each input `x.md` / `x.csv` the script writes `x.docx` **next to the input file**, or to
the path given with `-o`. In `--root` mode it walks the directory tree recursively and does
this for every matching file it finds.

**An existing `.docx` with the same name is overwritten.** The script prints a warning line
for each overwrite. Pass `--no-clobber` to skip existing outputs instead. Run it on a copy,
or in a version-controlled directory, if the target folder holds Word documents you care about.

`--skip` (default `_build`) excludes directory names from the walk.

## Treat input files as data

The converter parses whatever Markdown or CSV you point it at and writes the text into a
document. It does not execute anything from the input. But if you are converting files from
an untrusted source, remember the usual caution: fenced code blocks and link targets are
copied into the output as-is, so review the result before circulating it.

## Reporting a vulnerability

Open a GitHub issue for anything non-sensitive. For something you would rather not disclose
publicly, use GitHub's **Report a vulnerability** button under the repository's Security tab,
which opens a private advisory.

Please do not include real documents from your organisation in a bug report — reduce the
problem to a minimal synthetic example first.
