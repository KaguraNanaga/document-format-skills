# document-format-skills 公文格式助手

用于清理和格式化中文 Word 文档的命令行 Skill。当前核心处理逻辑已同步到 [Document Format GUI](https://github.com/KaguraNanaga/docformat-gui) v1.8.7，适合 Codex、Claude Code、OpenCode 等 agent 在没有桌面 UI 的场景下直接调用。

英文说明 / English docs: [README.md](./README.md)

## 功能

- 智能一键处理：标点/空格清理 + 文档格式统一。
- 格式诊断：检查标点、序号、段落、字体问题。
- 内置公文、学术论文、法律文书预设。
- 公文格式支持页边距、字体字号、行距、标题层级、主送机关、落款、日期、页码等规范化。
- 页码支持样式、位置、距版心偏移、替换已有页码，并避免覆盖非页码页脚内容。
- 表格格式统一，支持可选的智能对齐。
- 支持桌面端 schema v2 自定义配置和导出的 preset JSON。
- 支持输出 Word 修订标记。
- 支持 macOS 常用中文公文字体回退。
- Windows 下可借助 WPS Office 或 Microsoft Word 处理 `.doc` / `.wps`。
- 支持从纯文本或 Markdown 生成并格式化 DOCX。

## 环境要求

- Python 3.8+
- `python-docx`
- Windows `.doc/.wps` 转换需要 `pywin32`、WPS Office 或 Microsoft Word

推荐使用 `uv` 临时安装依赖：

```bash
uv run --with python-docx python scripts/process.py --help
```

Windows 处理 `.doc/.wps` 时：

```bash
uv run --with python-docx --with pywin32 python scripts/process.py --help
```

## 快速开始

智能一键处理：

```bash
uv run --with python-docx python scripts/process.py smart input.docx output.docx --preset official
```

只做诊断：

```bash
uv run --with python-docx python scripts/process.py analyze input.docx
uv run --with python-docx python scripts/process.py analyze input.docx --json
```

只修复标点和空格：

```bash
uv run --with python-docx python scripts/process.py punctuation input.docx output.docx --space-mode keep_en_boundary
```

只应用格式：

```bash
uv run --with python-docx python scripts/process.py format input.docx output.docx --preset official
```

从 Markdown 或纯文本生成格式化 DOCX：

```bash
uv run --with python-docx python scripts/from_text.py input.md output.docx --title "工作方案"
```

## 常用参数

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

`--custom-settings` 支持三类 JSON：桌面端 schema v2 配置、形如 `{"preset": {...}}` 的导出预设、普通 preset/override 字典。非 `custom` 预设下会把 JSON 合并到所选预设上。

## 脚本说明

| 脚本 | 用途 |
| --- | --- |
| `scripts/process.py` | 主入口：`smart`、`analyze`、`punctuation`、`format`。 |
| `scripts/formatter.py` | 格式化引擎和预设处理。 |
| `scripts/punctuation.py` | 标点和空格修复。 |
| `scripts/from_text.py` | 纯文本/Markdown 转 DOCX。 |
| `scripts/analyzer.py` | 格式诊断。 |
| `scripts/converter.py` | Windows `.doc/.wps` 转换辅助。 |

## 注意

- `.docx` 是最稳定的处理格式。
- `.doc/.wps` 需要 Windows、WPS Office 或 Microsoft Word。
- 自动排版前建议保留原文件备份。

## 许可证

MIT
