---
name: document-format-skills
description: Chinese Word document formatting toolkit for .docx/.doc/.wps workflows. Use when Codex needs to diagnose document formatting, fix mixed Chinese/English punctuation and spacing, apply official/academic/legal/custom presets, normalize tables and page numbers, preserve or output Word revision marks, convert plain text or Markdown into formatted DOCX, or batch/script document cleanup for Chinese official documents.
---

# Document Format Skills

Use these scripts to clean and format Chinese Word documents from the command line. Prefer `scripts/process.py` for normal work because it mirrors the desktop app's core pipeline without the GUI.

## Quick Workflow

Run one smart pass when the user wants the document cleaned end to end:

```bash
uv run --with python-docx python scripts/process.py smart input.docx output.docx --preset official
```

Run diagnostics only:

```bash
uv run --with python-docx python scripts/process.py analyze input.docx
uv run --with python-docx python scripts/process.py analyze input.docx --json
```

Run only punctuation/spacing cleanup:

```bash
uv run --with python-docx python scripts/process.py punctuation input.docx output.docx --space-mode keep_en_boundary
```

Run only formatting:

```bash
uv run --with python-docx python scripts/process.py format input.docx output.docx --preset official
```

On Windows, `.doc` and `.wps` input/output are supported through WPS or Microsoft Word COM automation:

```bash
uv run --with python-docx --with pywin32 python scripts/process.py smart input.wps output.wps
```

## Scripts

| Script | Use |
| --- | --- |
| `scripts/process.py` | One-shot CLI for `smart`, `analyze`, `punctuation`, and `format`; handles `.doc/.wps` conversion on Windows. |
| `scripts/formatter.py` | Apply formatting presets, custom JSON settings, page numbers, table cleanup, revision marks, macOS font fallback. |
| `scripts/punctuation.py` | Fix punctuation while preserving run formatting; supports spacing strategies. |
| `scripts/from_text.py` | Create a DOCX from `.txt` or Markdown, then optionally run smart formatting. |
| `scripts/analyzer.py` | Lower-level diagnostic script. |
| `scripts/converter.py` | Windows-only `.doc/.wps` conversion helpers. |

## Formatting Options

Built-in presets:

- `official`: GB/T 9704-2012 style official document formatting.
- `academic`: academic paper formatting.
- `legal`: legal document formatting.
- `custom`: read the active desktop custom preset when available.

Useful flags:

```bash
--custom-settings path.json
--revision
--deep-clean
--smart-table-align
--no-page-number
--page-number-style dash|plain|page_text|page_total
--page-number-position outside|left|center|right
--page-number-offset-mm 7
--no-bold-serial
```

`--custom-settings` accepts desktop schema v2 config files, exported preset files shaped as `{"preset": {...}}`, or plain preset/override JSON. For non-custom presets, the JSON is merged over the selected preset.

## Punctuation And Spacing

Punctuation cleanup protects URLs, email addresses, Windows paths, time values like `9:30`, and standards like `ISO 9001:2015`. It fixes brackets, colons, semicolons, question/exclamation marks, Chinese comma/period contexts, ellipses, dashes, and paired quotes.

Spacing modes:

- `remove_all`: delete half-width and full-width spaces.
- `keep_en_boundary`: remove Chinese-to-Chinese spaces but keep exactly one space between Chinese and English/digits.
- `keep_all`: leave spaces unchanged.

## Text Or Markdown To DOCX

Generate and format a document from text:

```bash
uv run --with python-docx python scripts/from_text.py input.md output.docx --title "工作方案"
```

Markdown mode detects headings, bold spans, ordered/unordered lists, quotes, and fenced code blocks. `#` becomes the main title, `##` becomes `一、`, `###` becomes `（一）`, and deeper headings become numbered lower-level headings.

Use `--no-process` to only create the raw DOCX.

## Implementation Notes

- `.docx` processing needs only `python-docx`.
- `.doc/.wps` conversion needs Windows plus WPS Office or Microsoft Word and `pywin32`.
- Page number handling avoids overwriting non-page footer content and can replace existing page-number footers when requested.
- Default table formatting preserves original alignment; use `--smart-table-align` for numeric/right and short-text/center alignment.
- macOS font handling keeps installed official fonts when present and falls back to compatible system fonts only when detection confirms the original is missing.
