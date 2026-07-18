# DOCX Editor / DOCX 编辑器

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)
[![CI](https://github.com/chuanyuepingfang/docx-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/chuanyuepingfang/docx-editor/actions/workflows/ci.yml)

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

# Create document from markdown
python scripts/create_docx.py --content "# Title\n\nBody" --output out.docx

# Format analysis report
python scripts/analyze_docx.py format-report document.docx

# Batch replace with tracked changes (requires Word)
python scripts/batch_replace.py input.docx output.docx --sections changes.json --track-changes
```

> **Note:** On Windows with Chinese content, set `PYTHONIOENCODING=utf-8` to avoid GBK encoding errors.

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
└── py.typed              # 类型检查标记
```

### 相关项目

- [python-docx](https://github.com/python-openxml/python-docx) — 底层的 DOCX 操作库
