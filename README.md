# 📄 document-format-skills

> **[中文版本 README / Chinese Version](./README_CN.md)**

> 💡 **想要图形界面版？**  
> 现已推出 **[Document Format GUI](https://github.com/KaguraNanaga/docformat-gui)** —— 无需联网、一键修复公文格式的桌面应用，小白也能轻松上手！

A Word document formatting toolkit for Chinese documents (docx). Diagnose formatting issues, fix punctuation, and apply standardized styles with one command. Available for Claude Code, Codex, OpenCode.

...

A Word document formatting toolkit for Chinese documents (docx). Diagnose formatting issues, fix punctuation, and apply standardized styles with one command. Available for Claude Code, Codex, OpenCode.

## ✨ Features

| Module | Description | Script |
|--------|-------------|--------|
| **Format Analyzer** | Detect formatting issues in documents | `analyzer.py` |
| **Punctuation Fixer** | Fix mixed Chinese/English punctuation | `punctuation.py` |
| **Style Formatter** | Apply preset formatting standards | `formatter.py` |

## 🚀 Quick Start

### Prerequisites

- Python 3.8+
- [uv](https://github.com/astral-sh/uv) (recommended) or pip

### Installation

```bash
git clone https://github.com/yourusername/document-format-skills.git
cd document-format-skills
```

### Usage

**1. Diagnose formatting issues**

```bash
uv run --with python-docx python scripts/analyzer.py input.docx
```

**2. Fix punctuation**

```bash
uv run --with python-docx python scripts/punctuation.py input.docx output.docx
```

**3. Apply formatting preset**

```bash
# Official document format (GB/T 9704-2012)
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset official

# Academic paper format
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset academic

# Legal document format
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset legal
```

## 📋 What It Fixes

### Punctuation Issues

The toolkit intelligently converts punctuation based on context:

| Type | Incorrect | Chinese | English |
|------|-----------|---------|---------|
| Parentheses | Mixed usage | （） | () |
| Quotes | Straight `"` | "" '' | "" '' |
| Colon | Mixed usage | ： | : |
| Comma | Mixed usage | ， | , |
| Period | Mixed usage | 。 | . |
| Semicolon | Mixed usage | ； | ; |
| Ellipsis | `...` | …… | ... |
| Dash | `--` | —— | -- |

### Format Issues

- **Paragraph indentation** — Detects missing first-line indents
- **Line spacing** — Identifies inconsistent spacing
- **Font usage** — Flags mixed fonts and sizes
- **Numbering** — Catches inconsistent numbering styles (e.g., mixing `1.` with `1、`)

## 📐 Formatting Presets

### Official Document (GB/T 9704-2012)

Chinese government document standard:

- **Page**: A4, margins: top 37mm, bottom 35mm, left 28mm, right 26mm
- **Title**: FangZheng XiaoBiaoSong, 22pt, centered
- **Body**: FangSong_GB2312, 16pt, 2-character indent, 28pt line spacing
- **Headings**: Structured with 一、/ （一）/ 1. / （1）

### Academic Paper

Standard academic formatting:

- **Page**: A4, 25mm margins
- **Title**: SimHei, 18pt, bold, centered
- **Body**: SimSun/Times New Roman, 12pt, 1.5x line spacing

### Legal Document

Legal document formatting:

- **Page**: A4, margins: top 30mm, bottom 25mm, left 30mm, right 25mm
- **Title**: SimSun bold, 22pt, centered
- **Body**: SimSun, 14pt, 1.5x line spacing

## 📁 Project Structure

```
document-format-skills/
├── README.md           # English documentation
├── README_CN.md        # Chinese documentation
├── SKILL.md            # Skill definition file
└── scripts/
    ├── analyzer.py     # Format diagnostics
    ├── punctuation.py  # Punctuation fixer
    └── formatter.py    # Style formatter
```

## 🔧 Dependencies

- [python-docx](https://python-docx.readthedocs.io/)

Automatically installed when using `uv run --with python-docx`.

## ⚠️ Notes

1. **Only supports `.docx`** — Legacy `.doc` format is not supported
2. **Backup your files** — Always keep a backup before processing
3. **Font requirements** — Output files require corresponding fonts installed on the system to display correctly
4. **Table content** — Text within tables is also processed

## 📄 License

MIT License

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
