# DOCX Editor / DOCX 编辑器

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)
[![CI](https://github.com/chuanyueaiqing/docx-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/chuanyueaiqing/docx-editor/actions/workflows/ci.yml)

**English** | [中文](#chinese)

A Python library for reading, editing, and creating DOCX files with **chapter-based editing**, **markdown conversion**, **format analysis**, and **tracked changes** support.

Built on top of [python-docx](https://github.com/python-openxml/python-docx) with extended capabilities for real-world document workflows.

---

## Features

- **Chapter-based editing** — Get, replace, or delete document chapters by number (`doc.get_chapter('3.1')`)
- **Markdown-to-DOCX conversion** — Full markdown support: headings, tables, lists, images, code blocks, **mermaid diagrams**
- **Table merging** — Custom `>` (merge right) and `v` (merge down) markers in markdown tables
- **Format extraction** — Read all OOXML formatting attributes via direct XML access
- **Format consistency analysis** — Detect cross-chapter formatting inconsistencies
- **Tracked changes** — Via Microsoft Word COM automation (`win32com`)
- **WPS Office support** — Alternative COM automation for WPS Office
- **Chinese heading detection** — Recognizes "第一章", "第1节", "第3.2条" etc.
- **Formula/equation support** — LaTeX-to-DOCX formula rendering via UnicodeMath and OMML
- **CLI tools** — Analyze, create, and batch-replace documents from the command line

## Installation

```bash
pip install python-docx
```

**Optional dependencies:**

```bash
pip install pywin32          # Tracked changes via Microsoft Word (Windows only)
pip install Pillow           # Image processing
npm install -g @mermaid-js/mermaid-cli  # Mermaid diagram rendering
```

## Quick Start

```python
from docx_editor import DocxDocument, DocxCreator

# Load and edit an existing document
doc = DocxDocument('document.docx')
chapter = doc.get_chapter('3.1')         # Get chapter content as markdown
doc.replace_chapter('3.1', '# New Section\n\nContent...')
doc.save('output.docx')

# Create a new document from markdown
DocxCreator.create('new.docx', '# Title\n\nBody content', format_spec={
    'font_name': 'Times New Roman',
    'font_name_east_asia': '宋体',
    'font_size': 12,
})
```

## CLI Usage

```bash
# Show chapter tree
python scripts/analyze_docx.py tree document.docx

# Replace a chapter
python scripts/analyze_docx.py replace document.docx "2.1" --content "# New" --output out.docx

# Create document from markdown (basic)
python scripts/create_docx.py --content "# Title\n\nBody" --output out.docx

# Create document with full formatting (font, spacing, images)
python scripts/create_docx.py --input doc.md --output report.docx \
    --font-name "Times New Roman" --font-ea "宋体" --font-size 12 \
    --line-spacing-rule EXACTLY --line-spacing 20 \
    --alignment JUSTIFY --first-line-indent 24 --space-after 6 \
    --image-max-width 15 --image-max-height 12 --image-align CENTER \
    --math-font "Cambria Math"

# Format analysis report
python scripts/analyze_docx.py format-report document.docx

# Batch replace with tracked changes (requires Word)
python scripts/batch_replace.py input.docx output.docx --sections changes.json --track-changes
```

> **Note:** On Windows with Chinese content, set `PYTHONIOENCODING=utf-8` to avoid GBK encoding errors.

### Mermaid Diagram Rendering

Mermaid diagrams in DOCX require `@mermaid-js/mermaid-cli` (mmdc) with a Chrome/Chromium browser.

```bash
# Install mmdc
npm install -g @mermaid-js/mermaid-cli
```

On first install, if Puppeteer fails to download its bundled Chrome (`chrome.exe` missing in cache), configure mmdc to use your system Chrome:

```bash
# Create ~/.mmdc.json pointing to your system Chrome
echo '{"executablePath": "C:/Program Files/Google/Chrome/Application/chrome.exe"}' > ~/.mmdc.json

# Then render with -p flag
mmdc -p ~/.mmdc.json -i diagram.mmd -o diagram.png
```

If this file is present, `create_docx.py` will automatically use it when rendering mermaid blocks.

## Claude Code Skill Integration

This repository includes a **Claude Code skill** (`skills/docx-mode/`) that enables AI-assisted DOCX editing through natural language commands.

### Installation

```bash
# Clone this repo (if you haven't already)
git clone https://github.com/chuanyueaiqing/docx-editor.git

# Install the skill for Claude Code
cp -r skills/docx-mode ~/.claude/skills/
```

Or, if you want to keep it in sync with the repo:

```bash
# Symlink instead of copy (Linux/Mac)
ln -s $(pwd)/skills/docx-mode ~/.claude/skills/docx-mode
```

### Usage

Once installed, tell Claude Code you want to work with a DOCX file. The skill will be auto-invoked when you mention `.docx` files, Word documents, comments, chapter edits, or report formatting.

Available commands inside the skill:

| Task | Command |
|------|---------|
| Create DOCX from markdown | `/docx-mode create output.docx --content "# Title"` |
| View chapter tree | `/docx-mode analyze tree document.docx` |
| Read chapter content | `/docx-mode analyze contents document.docx 3` |
| Replace a chapter | `/docx-mode analyze replace document.docx 2.1 --content "# New"` |
| Format analysis | `/docx-mode analyze format-report document.docx` |
| Batch replace | `/docx-mode batch input.docx output.docx --sections changes.json` |

The skill works in both English and Chinese — you can say things like "帮我替换第三章内容" or "replace chapter 3.1 with this markdown".

## Architecture

```
docx_editor/
├── __init__.py           # Public API: DocxDocument, DocxCreator
├── document.py           # Main document facade
├── creator.py            # Document creation from markdown
├── chapter_parser.py     # Chapter tree builder
├── markdown_processor.py # Markdown-to-DOCX conversion engine
├── format_extractor.py   # OOXML format extraction
├── format_analyzer.py    # Cross-chapter format comparison
├── table_builder.py      # Table with merge markers
├── mermaid_renderer.py   # Mermaid diagram rendering
├── models.py             # Data classes and type definitions
├── utils.py              # Utilities and custom exceptions
├── win32_ops.py          # Microsoft Word COM automation
├── wps_ops.py            # WPS Office COM automation
├── equation_ops.py       # Formula/equation insertion via win32com
├── latex_to_unicodemath.py  # LaTeX → UnicodeMath conversion
├── omml_builder.py       # OMML (Office Math Markup Language) builder
├── pandoc_omml.py        # Pandoc OMML post-processing
├── font_utils.py         # Font name and style utilities
└── py.typed              # Type checker marker
```

## Dependencies

| Package       | Required | Purpose                              |
|---------------|----------|--------------------------------------|
| python-docx   | Yes      | Core DOCX reading/writing            |
| lxml          | Yes*     | XML parsing (via python-docx)        |
| pywin32       | No       | Microsoft Word COM automation        |
| Pillow        | No       | Image processing                     |
| @mermaid-js/mermaid-cli | No | Mermaid diagram rendering       |

## Development

```bash
make install-dev    # Install dev dependencies
make test           # Run tests (skip Windows-only tests)
make test-all       # Run all tests
make coverage       # Run tests with coverage report
make build          # Build package
```

## Python Version Support

Python 3.8+ (tested on 3.9–3.12, Windows and Linux).

## License

[MIT](LICENSE)

---

<h2 id="chinese">中文说明</h2>

基于 python-docx 的 DOCX 编辑库，提供**章节级编辑**、**Markdown 转 DOCX**、**格式一致性分析**、**修订模式**、**WPS 集成**等功能。

### 功能特性

- **章节级编辑** — 按编号获取/替换/删除章节内容（支持 "3.1"、"第2章" 等格式）
- **Markdown 转 DOCX** — 完整支持标题、表格、列表、图片、代码块、**Mermaid 图表**
- **表格合并** — 在 Markdown 表格中使用 `>`（向右合并）和 `v`（向下合并）标记
- **格式提取** — 通过直接读取 OOXML XML 获取完整格式属性
- **格式一致性分析** — 检测跨章节的格式不一致问题
- **修订模式** — 通过 Microsoft Word COM 自动化实现
- **WPS 支持** — 通过 WPS COM 自动化实现修订
- **中文标题识别** — 支持"第一章"、"第1节"、"第3.2条"等中文章节编号
- **公式支持** — LaTeX 公式渲染为 Word 可编辑的 OMML/UnicodeMath 公式
- **命令行工具** — 分析、创建、批量替换文档

### 快速开始

```python
from docx_editor import DocxDocument, DocxCreator

# 加载并编辑
doc = DocxDocument('document.docx')
chapter = doc.get_chapter('3.1')          # 获取章节内容（Markdown 格式）
doc.replace_chapter('3.1', '# 新章节\n\n新内容...')
doc.save('output.docx')

# 创建新文档
creator = DocxCreator()
creator.set_default_format({'font_name': '宋体', 'font_size': 12})
creator.add_markdown('# 标题\n\n正文内容')
creator.save('new.docx')
```

### Claude Code 技能集成

本仓库包含一个 **Claude Code 技能**（`skills/docx-mode/`），让你可以通过自然语言与 Claude Code 对话来完成 DOCX 编辑操作。

#### 安装方法

```bash
# 克隆仓库
git clone https://github.com/chuanyueaiqing/docx-editor.git

# 安装技能到 Claude Code
cp -r skills/docx-mode ~/.claude/skills/
```

安装后，当你在 Claude Code 中提及 .docx 文件、Word 文档、批注、章节修改等关键词时，技能会自动激活。

#### 常用操作

| 任务 | 自然语言指令 |
|------|-------------|
| 从 Markdown 创建 DOCX | "帮我用这些 Markdown 内容创建一个 DOCX 文档" |
| 查看章节结构 | "帮我看看这个文档的章节树" |
| 替换章节内容 | "把第3.1节替换成以下内容" |
| 格式一致性检查 | "检查这个文档的格式是否一致" |
| 处理批注 | "帮我看看文档里有哪些批注" |
| 批量替换 | "按这个 JSON 配置批量替换章节" |

### 架构

```
docx_editor/
├── __init__.py           # 公开 API：DocxDocument, DocxCreator
├── document.py           # 文档门面
├── creator.py            # 从 Markdown 创建文档
├── chapter_parser.py     # 章节树构建
├── markdown_processor.py # Markdown 转 DOCX 引擎
├── format_extractor.py   # OOXML 格式提取
├── format_analyzer.py    # 跨章节格式对比
├── table_builder.py      # 表格合并标记处理
├── mermaid_renderer.py   # Mermaid 图表渲染
├── models.py             # 数据类与类型定义
├── utils.py              # 工具函数与异常
├── win32_ops.py          # Word COM 自动化
├── wps_ops.py            # WPS COM 自动化
├── equation_ops.py       # 公式插入（通过 win32com）
├── latex_to_unicodemath.py  # LaTeX → UnicodeMath 转换
├── omml_builder.py       # OMML 公式构建器
├── pandoc_omml.py        # Pandoc OMML 后处理
├── font_utils.py         # 字体工具
└── py.typed              # 类型检查标记
```

### 相关项目

- [python-docx](https://github.com/python-openxml/python-docx) — 底层的 DOCX 操作库
