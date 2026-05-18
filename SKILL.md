---
name: document-format-skills
description: 文档格式处理工具。支持格式诊断、标点符号修复、格式统一。输入杂乱的文档，输出规范整洁的docx。
---

# 文档格式处理工具

处理文档格式问题：诊断格式错误、修复标点符号、统一文档样式。

## 功能概览

| 功能 | 说明 | 脚本 |
|------|------|------|
| 格式诊断 | 分析文档存在的格式问题 | `analyzer.py` |
| Markdown 工作流 | Markdown 转换、标点修复、套用任意格式预设并运行最终诊断 | `md_to_docx.py` |
| 标点修复 | 修复中英文标点混用 | `punctuation.py` |
| 格式统一 | 应用预设格式规范 | `formatter.py` |
| 表格自动调整 | 自动调整表格布局与对齐 | `formatter.py` |
| 页码规范 | 统一页码格式与位置 | `formatter.py` |

## 使用方法

### Markdown 转 docx

将 Markdown 文件转换为 Word，并默认执行标点修复和最终诊断：

```bash
python3 scripts/md_to_docx.py input.md output.docx
```

支持 `#`、`##`、`###`、`####` 标题层级，支持列表、引用、代码块和基础加粗。可使用 `--title` 指定 Word 主标题，使用 `--overwrite` 覆盖已有输出文件。

### Markdown 转指定格式预设

Markdown 入口可直接接入现有格式预设：

```bash
python3 scripts/md_to_docx.py input.md --preset official
python3 scripts/md_to_docx.py input.md --preset academic
python3 scripts/md_to_docx.py input.md --preset legal
python3 scripts/md_to_docx.py input.md --preset custom
```

省略输出路径时，预设输出默认生成 `input_official.docx`、`input_academic.docx` 等。

### Markdown 仅转换不套格式

```bash
python3 scripts/md_to_docx.py input.md

# 只转换 Markdown，然后运行诊断
python3 scripts/md_to_docx.py input.md --draft-only
```

常用选项：`--title`、`--overwrite`、`--draft-only`、`--no-punctuation`、`--skip-diagnose`、`--keep-temp`、`--no-page-number`。

### 格式诊断

分析文档存在的问题，输出诊断报告：

```bash
uv run --with python-docx python3 scripts/analyzer.py input.docx
```

输出示例：
```
=== 格式诊断报告 ===

【标点问题】共 5 处
  - 第2段: 英文括号 () 建议改为 （）
  - 第3段: 英文引号 "" 建议改为 ""

【序号问题】共 2 处
  - 序号格式不统一: 同时存在 "1、" 和 "1." 
  - 第5段: 层级跳跃，从 "一、" 直接到 "1."

【段落问题】共 3 处
  - 第2、4、7段: 缺少首行缩进
  - 行距不统一: 存在单倍、1.5倍混用

【字体问题】共 2 处
  - 正文字号不统一: 12pt、14pt 混用
  - 检测到 4 种字体混用
```

### 标点符号修复

```bash
uv run --with python-docx python3 scripts/punctuation.py input.docx output.docx
```

### 格式统一

```bash
# 应用公文格式
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx --preset official

# 应用学术论文格式
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx --preset academic

# 应用法律文书格式
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx --preset legal
```

### 组合使用

```bash
# 先诊断
uv run --with python-docx python3 scripts/analyzer.py messy.docx

# 修复标点 + 应用格式
uv run --with python-docx python3 scripts/punctuation.py messy.docx temp.docx
uv run --with python-docx python3 scripts/formatter.py temp.docx clean.docx --preset official
```

---

## 标点符号处理规则

### 修复范围

| 类型 | 错误 | 正确（中文） | 正确（英文） |
|------|------|-------------|-------------|
| 括号 | 中英混用 | （） | () |
| 引号 | 直引号 "" | ""'' | "" '' |
| 冒号 | 中英混用 | ： | : |
| 逗号 | 中英混用 | ， | , |
| 句号 | 中英混用 | 。 | . |
| 分号 | 中英混用 | ； | ; |
| 问号 | 中英混用 | ？ | ? |
| 叹号 | 中英混用 | ！ | ! |
| 省略号 | ... | …… | ... |
| 破折号 | -- 或 — | —— | -- |

### 智能判断逻辑

1. **中文环境**：前后都是中文字符 → 用中文标点
2. **英文环境**：前后都是英文/数字 → 用英文标点
3. **混合环境**：默认用中文标点（可配置）

### 特殊处理

- 数字与单位之间：`100%` 保持英文
- 英文缩写：`e.g.` `i.e.` 保持英文句点
- 网址邮箱：保持原样不处理
- 代码块：跳过不处理

---

## 格式预设

### 公文格式（GB/T 9704-2012）

```
页面：A4，上边距37mm，下边距35mm，左边距28mm，右边距26mm
标题：方正小标宋简体，二号（22pt），居中
一级标题：黑体，三号（16pt），顶格，"一、"
二级标题：楷体_GB2312，三号（16pt），顶格，"（一）"
三级标题：仿宋_GB2312，三号（16pt），首行缩进，"1."
正文：仿宋_GB2312，三号（16pt），首行缩进2字符，行距固定值28pt
```

### 学术论文格式

```
页面：A4，边距25mm
标题：黑体，小二（18pt），居中
一级标题：黑体，小三（15pt），"1"
二级标题：黑体，四号（14pt），"1.1"
正文：宋体/Times New Roman，小四（12pt），首行缩进2字符，行距1.5倍
```

### 法律文书格式

```
页面：A4，上边距30mm，下边距25mm，左边距30mm，右边距25mm
标题：宋体加粗，二号（22pt），居中
条款标题：黑体，四号（14pt），"第一条"
正文：宋体，四号（14pt），首行缩进2字符，行距1.5倍
```

---

## 文件结构

```
document-format-skills/
├── SKILL.md
├── README.md
├── scripts/
│   ├── analyzer.py      # 格式诊断
│   ├── converter.py     # 旧格式转换辅助
│   ├── fix_spacing.py   # 行距修复辅助
│   ├── fix_spacing_simple.py # 简化行距修复辅助
│   ├── formatter.py     # 格式统一
│   ├── md_to_docx.py    # Markdown 转 docx 工作流
│   └── punctuation.py   # 标点修复
```

---

## 依赖

- python-docx

使用 `uv run --with python-docx` 自动安装。

---

## 注意事项

1. **核心处理对象为 .docx**：`.doc`/`.wps` 等旧格式转换依赖 `converter.py` 及特定平台的 Office/WPS 支持
2. **备份原文件**：修改前建议备份
3. **字体依赖**：输出文件需要系统安装对应字体才能正确显示
4. **表格内容**：会自动处理表格内的文字
5. **Markdown 转换保持基础能力**：支持常见标题、列表、引用、代码块和加粗文本，最终版式由诊断结果和格式预设继续处理
6. **Markdown 默认运行诊断**：批处理或静默输出时可使用 `--skip-diagnose`
