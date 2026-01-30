# 📄 document-format-skills

> **[English Documentation / 英文文档](./README.md)**

专业的 Word 文档格式处理工具包。一键诊断格式问题、修复标点符号、统一文档样式。

## ✨ 功能概览

| 模块 | 说明 | 脚本 |
|------|------|------|
| **格式诊断** | 分析文档存在的格式问题 | `analyzer.py` |
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

**1. 格式诊断**

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

**2. 修复标点**

```bash
uv run --with python-docx python scripts/punctuation.py input.docx output.docx
```

**3. 应用格式预设**

```bash
# 公文格式（GB/T 9704-2012）
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset official

# 学术论文格式
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset academic

# 法律文书格式
uv run --with python-docx python scripts/formatter.py input.docx output.docx --preset legal
```

**4. 组合使用**

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
    ├── analyzer.py     # 格式诊断
    ├── punctuation.py  # 标点修复
    └── formatter.py    # 格式统一
```

## 🔧 依赖

- [python-docx](https://python-docx.readthedocs.io/)

使用 `uv run --with python-docx` 时会自动安装。

## ⚠️ 注意事项

1. **只支持 .docx** — 不支持旧版 .doc 格式
2. **备份原文件** — 修改前建议备份
3. **字体依赖** — 输出文件需要系统安装对应字体才能正确显示
4. **表格内容** — 会自动处理表格内的文字

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交 Pull Request！
