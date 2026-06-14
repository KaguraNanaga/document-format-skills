# 📄 document-format-skills 公文格式助手

> **[中文版本 README / Chinese Version](./README_CN.md)**
>
> 看不懂英文也没关系，请点击上面的链接查看中文版说明。

> 💡 **想要无需联网、一键运行修复格式的桌面应用版本？**
>
> 现已推出 **[Document Format GUI 公文格式助手](https://github.com/KaguraNanaga/docformat-gui)** —— 无需联网、一键修复公文格式的桌面应用，小白也能轻松上手！

A Word document formatting toolkit for Chinese official documents for the government, CCP (China Communist Party), and the state-owned enterprises. Diagnose formatting issues, fix punctuation, and apply standardized styles with one command. Available for Claude Code, Codex, and OpenCode.

公文格式助手是一个面向中文公文排版的 Skill 工具包，用于诊断 Word 文档格式问题、修复中英文标点和空格混用、统一公文/论文/法律文书格式、处理页码和表格，并支持从纯文本或 Markdown 生成规范 DOCX。

## Features

- 中文用户如果不熟悉命令行，可以优先使用 [Document Format GUI 公文格式助手](https://github.com/KaguraNanaga/docformat-gui)：无需联网、一键修复公文格式。
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
