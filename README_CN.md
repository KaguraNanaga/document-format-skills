# 📄 document-format-skills

> **[English Documentation / 英文文档](./README.md)**

专业的 Word 文档格式（如DOCX格式）处理工具包。一键诊断格式问题、修复标点符号、统一文档样式。可用于Claude Code, Codex, OpenCode等。

## ✨ 功能概览

| 模块 | 说明 | 脚本 |
|------|------|------|
| **格式诊断** | 分析文档存在的格式问题 | `analyzer.py` |
| **Markdown 工作流** | Markdown 转换、标点修复、套用任意格式预设并运行最终诊断 | `md_to_docx.py` |
| **标点修复** | 修复中英文标点混用 | `punctuation.py` |
| **格式统一** | 应用预设格式规范 | `formatter.py` |

## 🚀 快速开始

### 环境要求

- Python 3.8+
- [uv](https://github.com/astral-sh/uv)（推荐）或 pip

### 安装

```bash
git clone https://github.com/yourusername/document-format-skills.git
cd document-format-skills
```

### 使用方法

**1. Markdown 转 DOCX**

```bash
python scripts/md_to_docx.py input.md output.docx
```

默认流程会完成 Markdown 转换、标点修复和最终诊断。

可使用 `--title` 指定 Word 主标题，使用 `--overwrite` 覆盖已有输出文件。

**2. Markdown 转指定格式预设**

同一个 Markdown 入口可以使用所有 formatter 预设：

```bash
python scripts/md_to_docx.py input.md --preset official
python scripts/md_to_docx.py input.md --preset academic
python scripts/md_to_docx.py input.md --preset legal
python scripts/md_to_docx.py input.md --preset custom
```

省略输出路径时，预设输出默认生成 `input_official.docx`、`input_academic.docx` 等。

常用选项：`--title`、`--overwrite`、`--draft-only`、`--no-punctuation`、`--skip-diagnose`、`--keep-temp`、`--no-page-number`。

**3. Markdown 仅转换不套格式**

```bash
python scripts/md_to_docx.py input.md

# 只转换 Markdown，然后运行诊断
python scripts/md_to_docx.py input.md --draft-only
```

**4. 格式诊断**

```bash
uv run --with python-docx python scripts/analyzer.py input.docx
```

输出示例：
```
=== 格式诊断报告 ===

【标点问题】共 5 处
  - 英文括号: 第2、3、5段
  - 英文引号: 第3段

【序号问题】共 2 处
  - 序号格式不统一: 同时存在 arabic_dot, arabic_comma

【段落问题】共 3 处
  - 缺少首行缩进: 第2、4、7段
  - 行距不统一: 存在 3 种不同行距

【字体问题】共 2 处
  - 字号不统一: 检测到 5 种字号
```

**5. 修复标点**

```bash
uv run --with python-docx python scripts/punctuation.py input.docx output.docx
```

**6. 应用格式预设**

```bash
# 公文格式（GB/T 9704-2012）
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset official

# 学术论文格式
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset academic

# 法律文书格式
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset legal
```

**7. 组合使用**

```bash
# 先诊断
uv run --with python-docx python scripts/analyzer.py messy.docx

# 修复标点 + 应用格式
uv run --with python-docx python scripts/punctuation.py messy.docx temp.docx
uv run --with python-docx python scripts/formatter.py temp.docx clean.docx --preset official
```

## 📋 修复内容

### 标点符号

智能根据上下文转换标点：

| 类型 | 错误示例 | 中文标点 | 英文标点 |
|------|----------|----------|----------|
| 括号 | 中英混用 | （） | () |
| 引号 | 直引号 `"` | "" '' | "" '' |
| 冒号 | 中英混用 | ： | : |
| 逗号 | 中英混用 | ， | , |
| 句号 | 中英混用 | 。 | . |
| 分号 | 中英混用 | ； | ; |
| 省略号 | `...` | …… | ... |
| 破折号 | `--` | —— | -- |

**智能判断逻辑：**
- 中文环境（前后都是中文字符）→ 使用中文标点
- 英文环境（前后都是英文/数字）→ 使用英文标点
- 混合环境 → 默认使用中文标点

### 格式问题

- **段落缩进** — 检测缺少首行缩进的段落
- **行距** — 识别不统一的行距设置
- **字体** — 标记混用的字体和字号
- **序号** — 发现不一致的序号风格（如 `1.` 和 `1、` 混用）

## 📐 格式预设

### 公文格式（GB/T 9704-2012）

符合国家标准的公文格式：

```
页面：A4，上边距37mm，下边距35mm，左边距28mm，右边距26mm
主标题：方正小标宋简体，二号（22pt），居中
一级标题：黑体，三号（16pt），"一、"
二级标题：楷体_GB2312，三号（16pt），"（一）"
三级标题：仿宋_GB2312，三号（16pt），"1."
四级标题：仿宋_GB2312，三号（16pt），"（1）"
正文：仿宋_GB2312，三号（16pt），首行缩进2字符，行距固定值28pt
```

### 学术论文格式

标准学术论文格式：

```
页面：A4，边距25mm
标题：黑体，小二（18pt），加粗，居中
一级标题：黑体，小三（15pt），"1"
二级标题：黑体，四号（14pt），"1.1"
正文：宋体/Times New Roman，小四（12pt），首行缩进2字符，行距1.5倍
```

### 法律文书格式

法律文书专用格式：

```
页面：A4，上边距30mm，下边距25mm，左边距30mm，右边距25mm
标题：宋体加粗，二号（22pt），居中
条款标题：黑体，四号（14pt），"第一条"
正文：宋体，四号（14pt），首行缩进2字符，行距1.5倍
```

## 📁 项目结构

```
document-format-skills/
├── README.md           # 英文文档
├── README_CN.md        # 中文文档
├── SKILL.md            # 技能定义文件
└── scripts/
    ├── analyzer.py              # 格式诊断
    ├── converter.py             # 旧格式文档转换辅助
    ├── fix_spacing.py           # 行距修复辅助
    ├── fix_spacing_simple.py    # 简化行距修复辅助
    ├── formatter.py             # 格式统一
    ├── md_to_docx.py            # Markdown 转 DOCX 工作流
    └── punctuation.py           # 标点修复
```

## 🔧 依赖

- [python-docx](https://python-docx.readthedocs.io/)

使用 `uv run --with python-docx` 时会自动安装。

## ⚠️ 注意事项

1. **核心处理对象为 .docx** — `.doc`/`.wps` 等旧格式转换依赖 `converter.py` 及特定平台的 Office/WPS 支持
2. **备份原文件** — 修改前建议备份
3. **字体依赖** — 输出文件需要系统安装对应字体才能正确显示
4. **表格内容** — 会自动处理表格内的文字
5. **Markdown 转换保持基础能力** — 支持常见标题、列表、引用、代码块和加粗文本，最终版式由诊断结果和格式预设继续处理
6. **Markdown 默认运行诊断** — 批处理或静默输出时可使用 `--skip-diagnose`

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交 Pull Request！
