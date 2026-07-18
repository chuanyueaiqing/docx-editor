#!/usr/bin/env python
"""批量替换 DOCX 中多个章节，支持修订模式。

用法:
  # 从 JSON 替换（直接替换）
  py scripts/batch_replace.py input.docx output.docx --sections content.json

  # 从 Markdown 替换（用 ## 3.1 等标题分隔章节）
  py scripts/batch_replace.py input.docx output.docx --sections content.md

  # 启用修订模式（一次 COM 对比，保留所有修订标记）
  py scripts/batch_replace.py input.docx output.docx --sections content.json --track-changes

  # 从现有文档生成模板
  py scripts/batch_replace.py input.docx --create-template template.json

JSON 格式:
  {
    "3.1": "# 3.1 标题\\n\\n正文...",
    "3.2": "# 3.2 标题\\n\\n正文..."
  }

Markdown 格式（用 ## 3.1 等标题作为章节分隔）:
  ## 3.1  标题
  正文...

  ## 3.2  标题
  正文...
"""
import argparse
import json
import os
import re
import sys
import tempfile

_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _PROJECT)

from docx_editor.document import DocxDocument
from docx_editor.win32_ops import Win32Ops


# ======================== Template generation ========================

def create_template(path: str, output: str):
    """从现有文档生成模板 JSON 文件，包含所有章节号和标题。"""
    doc = DocxDocument(path, use_track_changes=False)
    chapters = doc.get_chapter_tree()
    template = {}

    def walk(nodes, prefix=''):
        for node in nodes:
            num = node.to_string()
            heading = node.heading_text or ''
            key = num if num else prefix
            # 去掉章节号前缀，只保留纯标题文本（如 "DE-工具需求与约束"）
            pure_title = re.sub(r'^[\d.]+\s*', '', heading).strip()
            template[key] = f'# {num}  {pure_title}\n\n（请在此填写{num}的新内容）'
            if node.children:
                walk(node.children, prefix=key)

    walk(chapters)

    with open(output, 'w', encoding='utf-8') as f:
        json.dump(template, f, ensure_ascii=False, indent=2)
    print(f"模板已生成：{output}")
    print(f"共 {len(template)} 个章节，请编辑后执行替换命令。")


# ======================== Sections loading ========================

def load_sections(path: str) -> dict:
    """从 JSON 或 Markdown 文件加载章节内容映射。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.json':
        return _load_json(path)
    elif ext in ('.md', '.markdown'):
        return _load_markdown(path)
    else:
        raise ValueError(f"不支持的文件格式：{ext}，请使用 .json 或 .md")


def _load_json(path: str) -> dict:
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError("JSON 文件须为对象格式：{ \"3.1\": \"内容...\", ... }")
    return data


def _load_markdown(path: str) -> dict:
    """解析 Markdown 文件，用 ## 3.1 等标题作为章节分隔。"""
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配 ## 3.1 / # 3.1 或 ## 3.1.1 等（章节号 + 空格 + 可选标题）
    pattern = r'^(#{1,3})\s+([\d]+(?:\.[\d]+)*)\s*(.*?)$'
    lines = content.split('\n')

    sections = {}
    current_num = None
    current_lines = []

    for line in lines:
        m = re.match(pattern, line.strip())
        if m:
            # 保存上一节
            if current_num is not None and current_lines:
                sections[current_num] = '\n'.join(current_lines).strip()

            # 开始新节
            current_num = m.group(2)
            heading_level = m.group(1)
            heading_text = (m.group(3) or '').strip()
            # 重建 heading（保持原始级别和格式）
            heading = f'{heading_level} {current_num}'
            if heading_text:
                heading += f'  {heading_text}'
            current_lines = [heading]
        else:
            if current_num is not None:
                current_lines.append(line)

    # 最后一节
    if current_num is not None and current_lines:
        sections[current_num] = '\n'.join(current_lines).strip()

    return sections


# ======================== Batch replace ========================

def accept_all_tracked_changes(input_path: str) -> str:
    """用 Word COM 接受文档中所有既有修订，返回干净基线文件路径。"""
    clean_path = tempfile.NamedTemporaryFile(suffix='.docx', delete=False).name
    print("正在通过 Word 接受文档中所有既有修订...")
    with Win32Ops() as ops:
        ops.open_document(input_path)
        ops.word.ActiveDocument.AcceptAll()
        ops.wd_doc = ops.word.ActiveDocument
        ops.save_document(clean_path)
    print("  既有修订已全部接受，保存为干净基线版本。")
    return clean_path


def check_existing_tracked_changes(path: str) -> bool:
    """检查文档中是否存在既有修订标记。"""
    import zipfile, xml.etree.ElementTree as ET
    try:
        with zipfile.ZipFile(path) as z:
            if 'word/document.xml' not in z.namelist():
                return False
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            ins = root.findall('.//w:ins', ns)
            dell = root.findall('.//w:del', ns)
            return len(ins) > 0 or len(dell) > 0
    except Exception:
        return False


def batch_replace(
    input_path: str,
    output_path: str,
    sections: dict,
    track_changes: bool,
):
    """批量替换章节。

    - track_changes=False：直接替换并保存
    - track_changes=True：内存中替换 → 一次 COM 对比，保留完整修订标记
    """
    # Step 0: 检查既有修订
    has_existing = check_existing_tracked_changes(input_path)
    working_input = input_path

    if track_changes and has_existing:
        print(f"检测到原文档有既有修订，正在接受...")
        working_input = accept_all_tracked_changes(input_path)

    # Step 1: 加载文档（不使用修订模式，避免 python-docx 的 COM 对比）
    print("正在加载文档...")
    doc = DocxDocument(working_input, use_track_changes=False)

    # Step 2: 逐个替换章节
    chapter_nums = sorted(sections.keys(), key=lambda x: [int(p) for p in x.split('.')])
    for ch_num in chapter_nums:
        content = sections[ch_num]
        print(f"  替换章节 {ch_num} ...")
        doc.replace_chapter(ch_num, content)

    if track_changes:
        # Step 3a: 保存修改版到临时文件
        modified_tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
        modified_tmp.close()
        doc.document.save(modified_tmp.name)
        print("修改版已保存到临时文件。")

        # Step 3b: Word COM 对比
        print("正在启动 Word 进行修订对比...")
        with Win32Ops() as ops:
            ops.open_document(working_input)
            ops.word.ActiveDocument.Compare(
                Name=modified_tmp.name,
                CompareTarget=2,          # wdCompareTargetNew
                IgnoreAllComparisonWarnings=True,
            )
            ops.wd_doc = ops.word.ActiveDocument
            ops.save_document(output_path)

        # Step 4: 清理
        os.unlink(modified_tmp.name)
        if has_existing and working_input != input_path:
            os.unlink(working_input)
    else:
        # Step 3a（直接模式）：直接保存
        doc.save(output_path)

    print(f"\n完成！输出文件：{output_path}")


# ======================== CLI ========================

def main():
    parser = argparse.ArgumentParser(
        description='批量替换 DOCX 章节内容',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('output', nargs='?', help='输出 .docx 文件路径（不指定则覆盖输入文件）')
    parser.add_argument('--sections', help='章节内容文件（.json 或 .md）')
    parser.add_argument('--track-changes', action='store_true',
                        help='启用修订模式（Word COM 对比）')
    parser.add_argument('--create-template', metavar='OUTPUT',
                        help='从文档生成模板文件（JSON），不执行替换')

    args = parser.parse_args()

    if args.create_template:
        create_template(args.input, args.create_template)
        return

    if not args.sections:
        parser.error("请指定 --sections 或 --create-template")

    sections = load_sections(args.sections)
    if not sections:
        print("错误：未找到任何章节内容，请检查文件格式。")
        sys.exit(1)

    output = args.output or args.input
    batch_replace(args.input, output, sections, track_changes=args.track_changes)


if __name__ == '__main__':
    main()
