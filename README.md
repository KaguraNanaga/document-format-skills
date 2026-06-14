# document-format-skills

Command-line skills for cleaning and formatting Chinese Word documents. This repo mirrors the core processing logic from [Document Format GUI](https://github.com/KaguraNanaga/docformat-gui) v1.8.7 so agents such as Codex, Claude Code, and OpenCode can run the same document pipeline without the desktop UI.

Chinese docs: [README_CN.md](./README_CN.md)

## Features

- Smart one-shot processing: punctuation/spacing cleanup plus formatting.
- Format diagnosis for punctuation, numbering, paragraph, and font issues.
- Official, academic, legal, and custom presets.
- GB/T 9704-2012 style page margins, fonts, line spacing, headings, signatures, dates, and page numbers.
- Safer page-number handling with styles, positions, offsets, replacement control, and non-page footer protection.
- Table normalization with optional smart alignment.
- Custom settings compatible with the desktop app schema v2 and exported preset JSON.
- Word revision marks for supported formatting changes.
- macOS font fallback for common Chinese official-document fonts.
- `.doc` / `.wps` conversion on Windows when WPS Office or Microsoft Word is installed.
- Plain text or Markdown to formatted DOCX.

## Requirements

- Python 3.8+
- `python-docx`
- `pywin32` only for `.doc/.wps` conversion on Windows

Use `uv` for ad-hoc runs:

```bash
uv run --with python-docx python scripts/process.py --help
```

For Windows `.doc/.wps` conversion:

```bash
uv run --with python-docx --with pywin32 python scripts/process.py --help
```

## Quick Start

Smart cleanup:

```bash
uv run --with python-docx python scripts/process.py smart input.docx output.docx --preset official
```

Analyze only:

```bash
uv run --with python-docx python scripts/process.py analyze input.docx
uv run --with python-docx python scripts/process.py analyze input.docx --json
```

Punctuation and spacing only:

```bash
uv run --with python-docx python scripts/process.py punctuation input.docx output.docx --space-mode keep_en_boundary
```

Formatting only:

```bash
uv run --with python-docx python scripts/process.py format input.docx output.docx --preset official
```

Create a formatted DOCX from Markdown or text:

```bash
uv run --with python-docx python scripts/from_text.py input.md output.docx --title "Work Plan"
```

## Useful Options

```bash
--preset official|academic|legal|custom
--custom-settings path.json
--revision
--deep-clean
--smart-table-align
--no-page-number
--page-number-style dash|plain|page_text|page_total
--page-number-position outside|left|center|right
--space-mode remove_all|keep_en_boundary|keep_all
```

`--custom-settings` accepts desktop schema v2 config files, exported preset files like `{"preset": {...}}`, or plain preset/override JSON.

## Scripts

| Script | Purpose |
| --- | --- |
| `scripts/process.py` | Main CLI pipeline: `smart`, `analyze`, `punctuation`, `format`. |
| `scripts/formatter.py` | Formatting engine and preset handling. |
| `scripts/punctuation.py` | Punctuation and spacing fixer. |
| `scripts/from_text.py` | Text/Markdown to DOCX generator. |
| `scripts/analyzer.py` | Diagnostic helpers. |
| `scripts/converter.py` | Windows `.doc/.wps` conversion helpers. |

## Notes

- `.docx` is the most reliable format.
- `.doc/.wps` requires Windows plus WPS Office or Microsoft Word.
- Keep a backup of important documents before running automated formatting.

## License

MIT
