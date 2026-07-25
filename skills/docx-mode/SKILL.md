---
name: docx-mode
description: |
  对 .docx 文件进行读取、创建、编辑和批注审核的专业工具。支持三种操作模式：
  (1) 将 Markdown 渲染为全新的 DOCX（创建前先询问用户格式偏好）；
  (2) 替换现有 DOCX 中的章节内容（自动检测跨章节格式不一致并让用户确认）；
  (3) 逐条读取和处理文档中的批注（展示章节上下文 + 批注意见，让用户选择应用、跳过或删除）。
  当用户提到 .docx 文件、Word 文档、批注、章节修改、报告排版、通篇改格式时，都应该使用这个 skill。
---

# DOCX Mode Skill

基于 `docx_editor` 库（E:\demo\docx\docx_editor），提供完整的 DOCX 文档创建、修改和批注审核能力。

---

## Quick Reference

| 任务 | 命令 |
|------|------|
| 从 Markdown 创建 DOCX | `PYTHONIOENCODING=utf-8 py scripts/create_docx.py --content "..." --output out.docx [格式参数] [--math-font "Times New Roman"]` |
| 查看章节树 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py tree path.docx` |
| 读取章节内容 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py contents path.docx 3` |
| 格式一致性报告 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py format-report path.docx` |
| 替换章节 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py replace path 2.1 --content "..." [--track-changes] --output out.docx` |
| 批量替换章节 | `PYTHONIOENCODING=utf-8 py scripts/batch_replace.py in.docx out.docx --sections sections.json [--track-changes]` |
| 从文档生成模板 | `PYTHONIOENCODING=utf-8 py scripts/batch_replace.py in.docx --create-template template.json` |
| 查看批注上下文 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py comments-with-context path.docx` |
| 删除单条批注 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py delete-comment path.docx <comment_id>` |
| 删除所有批注 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py delete-all-comments path.docx` |
| **一次性文档验证** | `PYTHONIOENCODING=utf-8 py scripts/verify_docx.py path.docx` |
| 通篇改正文格式 | `PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx --target body --font-name 黑体 --font-size 14` |
| 插入目录 | `py -c "from docx_editor import DocxDocument; d=DocxDocument('path.docx'); d.insert_toc(); d.save()"` |
| 刷新目录 | `py -c "from docx_editor import DocxDocument; d=DocxDocument('path.docx'); d.refresh_toc(); d.save()"` |
| 检查是否有目录 | `py -c "from docx_editor import DocxDocument; d=DocxDocument('path.docx'); print(d.has_toc())"` |
| 删除目录 | `py -c "from docx_editor import DocxDocument; d=DocxDocument('path.docx'); d.remove_toc(); d.save()"` |

---

## Core Concepts

- **docx_editor 库**：位于 `E:\demo\docx\docx_editor\`，所有操作都通过 `scripts/` 下的 CLI 脚本调用，不需要手写 Python
- **章节机制**：文档按标题（Heading 样式或中文"第一章"等）自动分割为章节。`tree` 命令查看完整结构
- **目录（TOC）自动刷新**：一旦通过 `insert_toc()` 插入了目录，后续对章节进行 `replace_chapter()` 或 `delete_chapter()` 操作时，目录会自动刷新——删除旧条目、基于新章节树重新生成。无需手动干预。
  - `refresh_toc()`：手动触发刷新
  - `remove_toc()`：删除目录
  - `has_toc()`：检查是否存在目录
- **⚠️ 强制纪律：Markdown 内容中禁止使用 `1.` `2.` 等数字编号列表**：章节树解析器会将行首的 `数字+.+空格` 模式（如 `1. 良率分析`）误识别为标题，导致产生大量假章节，后续 `replace` 命令会打错位置，破坏文档结构。
  - ✅ **正确做法**：一律使用 `- ` 破折号列表（如 `- 良率分析`）或 `（1）` `（2）` 中文括号编号
  - ❌ **禁止使用的格式**：`1. ` `2. ` `3. ` 等数字句点开头的列表
  - 如果用户提供的 Markdown 内容包含这种格式，**必须**先转换为合规格式再执行替换
- **修订模式**：`--track-changes` 会依次尝试 Microsoft Word → WPS Office → python-docx 直接保存。
  - Word 不可用或 WPS 冲突时，自动切换到 WPS 的 COM 接口（需要 WPS Office 已安装）
  - 两者都不可用时回退到直接替换（无修订标记）
- **Mermaid 图**：如果在创建或替换时用到 mermaid 代码块，需要 `mmdc` CLI 来渲染为图片。不可用时回退为代码块
- **⚠️ 谨慎使用 `（1）` `（2）` 等括号编号**：在修订模式的 COM 对比中，括号编号 + 大段新增内容的组合偶发内容截断（如"加盐哈希"丢失、"HTTPS/TLS"在斜杠处截断）。优先使用中文序号（`第一` `第二` `第三`）或破折号列表（`- 条目`）。如果必须使用括号编号，替换后务必用 Word 打开确认内容完整性
- **表格合并（自定义 Markdown 扩展）**：标准 Markdown 不支持单元格合并，本工具通过特殊的单元格标记实现：
  - `>` 表示**向右合并**——该单元格与左边的单元格合并
  - `v` 表示**向下合并**——该单元格与上方的单元格合并
  - 被合并的单元格内容必须留空（用 `>` 或 `v` 作为单元格内容）
  - `merge_map` 自动从表格解析中生成，与 `rows` 一一对应
  - 示例：
    ```markdown
    | 姓名 | 项目 | 得分 |
    |------|------|------|
    | 张三 | 语文 | 95   |
    | v    | 数学 | 88   |
    | v    | 英语 | >    |
    ```
    效果：第一列"张三"向下合并 3 行，第三列"95"和"88"分别独立，"88"向右与空单元格合并。
  - 合并顺序：先处理横向合并（`>`），再处理纵向合并（`v`）。这与 Word OOXML 的 gridSpan / vMerge 机制一致。
- **修订模式**：`--track-changes` 会依次尝试 Microsoft Word → WPS Office → python-docx 直接保存。
  - Word 不可用或 WPS 冲突时，自动切换到 WPS 的 COM 接口（需要 WPS Office 已安装）
  - 两者都不可用时回退到直接替换（无修订标记）
- **Mermaid 图**：如果在创建或替换时用到 mermaid 代码块，需要 `mmdc` CLI 来渲染为图片。不可用时回退为代码块
- **数学公式（$...$ / $$...$$）**：文档中的 LaTeX 数学公式自动转为 Word 原生可编辑公式（OMML 格式）
  - 转换引擎使用 **Pandoc**（需要系统中安装 pandoc，`pip install pypandoc`），优先将 LaTeX 转为标准 OMML
  - Pandoc 不可用时自动回退到内置 OMML builder（支持基础公式：分式、根式、上下标、希腊字母等）
  - **公式字体配置**：通过 `--math-font "Times New Roman"` 或 `set_default_format({'math_font': 'Times New Roman'})` 指定公式专用字体，默认 Cambria Math
  - 14/15 的常见复杂公式可通过 Pandoc 正确转换（极限：5 层以上嵌套连分数会回退到内置 builder）
- **工作目录**：所有命令应在 `E:\demo\docx\` 下执行
- **Windows 编码注意**：在 Git Bash 终端下运行 `py scripts/` 命令时，如果文档包含中文，必须添加 `PYTHONIOENCODING=utf-8` 环境变量前缀，避免 GBK 编码报错。例如：`PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py comments-with-context path.docx`

---

## Mode 1: 从 Markdown 创建新 DOCX

当用户想创建一个全新的 Word 文档时使用此模式。

### 工作流

**Step 1: 询问格式偏好**

在创建文档前，**必须先**向用户询问以下格式偏好。不要使用默认值而不问用户：

| 项目 | 典型值 | 说明 |
|------|--------|------|
| 中文字体 | 宋体 / 黑体 / 微软雅黑 | 正文主要字体 |
| 西文字体 | Times New Roman / Arial / Calibri | 英文/数字字体 |
| 正文字号 | 12pt (小四) / 10.5pt (五号) | 按需选择 |
| 行距 | 1.5 倍 / 1.25 倍 / 单倍 | 推荐 1.5 倍 |
| 对齐方式 | 两端对齐 / 左对齐 | 默认两端对齐 |
| 首行缩进 | 2 字符 (约 24pt) | 中文正文常见 |
| 段后间距 | 6pt / 0pt | 段落之间间距 |
| 公式字体 | Cambria Math / Times New Roman | 公式专用字体（默认 Cambria Math） |

如果用户表示"按默认就行"，使用：宋体 / Times New Roman / 12pt / 1.5 行距 / 两端对齐 / 首行缩进 24pt / 段后 6pt。公式字体默认 Cambria Math。

**Step 2: 用户提供需求或 Markdown 内容**

可以是描述需求的自然语言、直接输入的 Markdown、或 .md 文件路径。

- 如果用户给的是需求（如"帮我写一份关于 XX 的报告"），由我来生成 Markdown 内容
- 如果用户给的是现成的 Markdown，直接使用

**Step 3: 展示 Markdown 内容给用户确认**

在运行任何创建命令之前，**必须**先将将要写入的 Markdown 内容展示给用户审阅。

```
用户确认格式后，展示以下内容等待确认：

─────────────────────────────────
# 标题

正文内容...

$$ 公式 $$

- 条目
─────────────────────────────────

以上内容是否确认？[Y/N]
```

- 用户 **确认后** 才进入下一步
- 用户如果要求修改，修改后重新展示，直到用户确认为止
- **禁止**在用户确认之前写入文件

**Step 4: 运行创建命令**

用户确认 Markdown 内容后，执行创建：

```bash
PYTHONIOENCODING=utf-8 py scripts/create_docx.py \
  --content "# 标题\n\n正文内容..." \
  --output output.docx \
  --font-name "Times New Roman" \
  --font-ea "宋体" \
  --font-size 12 \
  --line-spacing 1.5 \
  --alignment JUSTIFY \
  --first-line-indent 24 \
  --space-after 6 \
  --math-font "Cambria Math"
```

如果用户有 .md 文件，用 `--input file.md` 代替 `--content`。

**Step 5: 询问是否需要校验（非自动）**

创建完成后，**询问用户**是否需要运行校验：

```
文档已写入。是否需要运行格式校验（verify_docx.py）来检查字体/行距/内容？[Y/N]
```

- 用户说 **Y** → 运行 `verify_docx.py`，输出结果给用户
- 用户说 **N** → 跳过，直接进入下一步
- **不允许**在用户未确认的情况下自动运行校验

校验命令：

```bash
# 单条命令输出全部信息：格式、字数、章节、公式数、统计
PYTHONIOENCODING=utf-8 py scripts/verify_docx.py output.docx

# JSON 模式（程序化检查）
PYTHONIOENCODING=utf-8 py scripts/verify_docx.py output.docx --json
```

输出涵盖：默认格式（字体/字号/行距/对齐/缩进）、内容统计（段落/标题/表格/图片/公式/字数）、章节结构、错误/警告。

**Step 6 (可选): 插入目录**

如果用户希望文档包含目录，在创建完成后询问并执行：

```bash
# 插入默认目录（标题"目录"，最多 3 级标题）
py -c "
from docx_editor import DocxDocument
d = DocxDocument('output.docx')
d.insert_toc(max_level=3)   # max_level 控制最大标题层级
d.save()
"

# 或使用 Word TOC 域（在 Word 中按 Ctrl+A → F9 刷新）
py -c "
from docx_editor import DocxDocument
d = DocxDocument('output.docx')
d.insert_toc(max_level=3, use_word_field=True)
d.save()
"
```

**副作用**：插入目录后，后续对文档做章节替换或删除时，目录会自动刷新。

---

## Mode 2: 修改现有 DOCX 中的章节

当用户想修改文档中特定章节的内容时使用此模式。

### ⚠️ 重要：确认修订模式

**在开始任何修改工作流之前，必须先询问用户是否要启用修订模式（track-changes）。**

修改时涉及的章节可能较多，应让用户提前决定，避免中途切换。

```
用户：帮我把第三章内容改一下
你：  是否需要启用修订模式（track-changes）来保留修改痕迹？
      如果启用，需要 Microsoft Word 且不能有 WPS 干扰保存流程。
```

如果用户选择启用，但 Word 不可用或 WPS 冲突时，fallback 到直接修改并告知用户。

---

### 工作流

**Step 1: 查看章节结构**

```bash
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py tree path.docx
```

输出示例：
```
1.1 项目背景 (3 body para)
1.2 项目目标 (2 body para)
  1.2.1 短期目标 (2 body para, 1 children)
  1.2.2 长期目标 (1 body para)
2 技术方案 (5 body para, 2 children)
```

**Step 2: 检测格式一致性（重要）**

**必须**先运行格式分析：

```bash
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py format-report path.docx
```

- 如果所有章节的正文格式一致，输出会显示"未发现格式不一致"——可以继续
- **如果检测到不一致**（例如第 1 章正文是 12pt 宋体，第 2 章是 10.5pt 黑体），**必须**：
  1. 将差异展示给用户
  2. 让用户确认本次修改应该使用哪种格式
  3. 如果用户选择了统一格式，先执行 `apply-format`：

```bash
# 统一正文格式为 12pt 宋体
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx \
  --target body --font-name "Times New Roman" --font-ea "宋体" --font-size 12
```

**Step 3: 展示替换内容给用户确认**

根据用户的需求生成章节的新 Markdown 内容后，**必须**展示给用户审阅：

```
以下是将要替换 2.1 节的内容：
─────────────────────────────────
# 新标题

新正文内容...

$$ 公式 $$

- 条目
─────────────────────────────────
确认替换？[Y/N]
```

- 用户 **确认后** 才执行替换
- 用户要求修改则修改后重新展示，直到确认

**Step 4: 执行替换**

用户确认后，询问是否启用修订模式：

- **Word 修订模式**（需 Microsoft Word + pywin32） → 加 `--track-changes`
  - 自动降级：如果 Word 不可用或 WPS 冲突，会自动尝试 WPS Office COM 接口
  - 两者都失败时自动回退到直接替换（无修订标记）
- 直接替换（不可追踪）：

```bash
# 替换第 2.1 章，启用修订模式
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py replace path.docx "2.1" \
  --content "# 新标题\n\n新正文" \
  --track-changes \
  --output output.docx

# 或从文件读取 Markdown
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py replace path.docx "2.1" \
  --input new_content.md \
  --output output.docx
```

**Step 5: 询问是否需要校验（非自动）**

替换完成后，询问用户是否需要运行校验：

```
替换完成。需要运行格式校验来确认结果吗？[Y/N]
```

- 用户说 **Y** → 运行 `verify_docx.py`
- 用户说 **N** → 跳过

```bash
# 一次性检查结构和格式
PYTHONIOENCODING=utf-8 py scripts/verify_docx.py output.docx

# 如需查看指定章节的具体文本
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py contents output.docx "2.1"
```

**⚠️ 目录（TOC）自动刷新**：如果文档之前已经插入了目录，`replace_chapter` 和 `delete_chapter` 会自动刷新目录——删除旧的条目，基于最新的章节树重新生成。无需手动操作。

如果需要手动刷新（例如批量修改后确认目录是最新的）：

```bash
py -c "
from docx_editor import DocxDocument
d = DocxDocument('output.docx')
d.refresh_toc()
d.save()
"
```

如果不需要目录了：

```bash
py -c "
from docx_editor import DocxDocument
d = DocxDocument('output.docx')
d.remove_toc()
d.save()
"
```

### 批量替换多个章节（含修订模式）

#### 场景 A：不需要修订模式（直接替换）

可以逐个执行 `replace` 命令，每个替换独立执行。无特殊注意事项。

#### 场景 B：需要修订模式（--track-changes）

**⚠️ 禁止逐个替换！** 原因如下：

`replace --track-changes` 的工作流程是：python-docx 替换 → Word COM 对比 → 保存到文件 → **python-docx 重载（吃掉所有既有修订）** → 下一轮。所以只有最后一轮的修订标记会保留，前面几轮的全丢了。

**正确做法：一次替换全部内容，一次性 COM 对比。**

具体方案：
1. 用 `use_track_changes=False` 加载文档
2. 在 python-docx 内存中完成所有章节的替换（不触发 COM 对比）
3. 保存修改版到临时文件
4. **一次** Word COM 对比：原始文档 vs 修改版
5. 输出文件包含全部章节的修订标记

参考脚本：`E:\demo\docx\scripts\batch_replace.py`（通用命令行工具）

```bash
# 1. 从文档生成模板（自动提取所有章节号和标题）
cd E:/demo/docx && PYTHONIOENCODING=utf-8 py scripts/batch_replace.py input.docx --create-template template.json

# 2. 编辑 template.json 填写新内容（JSON 格式，key=章节号，value=Markdown）

# 3. 直接替换（无修订标记）
cd E:/demo/docx && PYTHONIOENCODING=utf-8 py scripts/batch_replace.py input.docx output.docx --sections template.json

# 4. 修订模式（一次 COM 对比，保留全部修订标记）
cd E:/demo/docx && PYTHONIOENCODING=utf-8 py scripts/batch_replace.py input.docx output.docx --sections template.json --track-changes
```

也可用 Markdown 模板（用 `## 3.1` 等标题分隔章节）：

```bash
cd E:/demo/docx && PYTHONIOENCODING=utf-8 py scripts/batch_replace.py input.docx output.docx --sections content.md [--track-changes]
```

**额外注意**：如果原文档本身有未接受的修订，python-docx 加载时会自动接受它们，导致信息丢失。应在 python-docx 介入之前，先用 Word COM 打开文档执行 `AcceptAll()`，将结果保存为干净基线。

```python
with Win32Ops() as ops:
    ops.open_document(input_path)
    ops.word.ActiveDocument.AcceptAll()  # 接受所有既有修订
    ops.wd_doc = ops.word.ActiveDocument
    ops.save_document(clean_path)
```

---

### ⚠️ 修订模式验证陷阱：python-docx 无法读取修订标记内容

**经验教训**：`--track-changes`（Word COM Compare）生成的修订版 DOCX 中，新增内容存放在 `<w:ins>` XML 元素内。python-docx 的 `r.text` / `p.text` 以及 `analyze_docx.py contents` **无法读取 `<w:ins>` 中的文本**，导致回读时看起来像"内容被截断丢失"。

**这通常是误报**——实际文件在 Word 中打开是完整正确的。以下是正确的验证方法：

#### ✅ 正确验证方法：XML 文本提取

```bash
cd E:/demo/docx && PYTHONIOENCODING=utf-8 py -c "
import zipfile, re
z = zipfile.ZipFile('path/to/output.docx')
xml = z.read('word/document.xml').decode('utf-8')
z.close()

# 提取所有 <w:t> 文本（含 ins 内的）和 <w:delText>
texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
full_text = ''.join(texts)

# 关键术语检查（需空白归一化，避免跨元素边界漏报）
normalized = re.sub(r'\s+', '', full_text)
terms = ['30秒', 'Dashboard', 'DataManager', 'SLA', 'TLS1.2']
for t in terms:
    if t in normalized:
        print(f'✅ {t}')
    else:
        print(f'❌ 可能丢失: {t}')
"
```

#### ❌ 错误验证方法（不要用）

- `analyze_docx.py contents` → 无法读取 `<w:ins>` 中的文本
- python-docx 的 `paragraph.text` / `r.text` → 同样无法读取
- `python -m markitdown` → 同样无法处理修订标记
- 原始 XML 中简单 `"keyword" in xml` → 可能因跨元素边界漏报

#### 最可靠的方法

**用 Word 打开文件**，打开"审阅"窗格，肉眼确认修订标记完整。或者先接受所有修订后另存为新文件，再用 `contents` 命令验证。

---

## Mode 3: 批注审核

当用户想查看和处理文档中的批注（评论）时使用此模式。

### 工作流

**Step 1: 获取批注列表**

```bash
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py comments-with-context path.docx
```

输出采用**线程化展示**，自动识别批注之间的回复关系：

- 同一段落的批约会归为同一个**对话线程**，回复排在发起者下方
- 每条批注显示序号（`[#n]`）、作者、日期和批注内容
- 仅线程发起者显示上下文段落（`>>>` 标记批注所在行）
- 未映射到章节的批注单独列出

示例输出：
```
19 条批注，5 个对话线程

── 线程 1 ── 3.5 配置管理与质量保证约束 （2 条）──
[#3] 真 [2]  2026-07-04 17:45
研制总结报告中写的是集成测试与系统测试两级体系...
    3.5 配置管理与质量保证约束
>>> 关键功能须经过集成测试和合格性测试两级验证。
    ... (共 7 段)

[#4] 您有一份快递到了，请签收  2026-07-05 10:01
舒冲统一看一下

── 单独批注 ── (未映射到章节) ──
[#19] 您有一份快递到了，请签收  2026-07-05 10:06
这个是预防阶段，提前调研，应该没啥问题吧？
```

**Step 2: 逐条与用户交互**

对每条批注（用 `[#n]` 序号引用），Claude 应该：

1. **展示完整信息给用户**（可参照上面线程化输出的格式）

2. **解释批注含义**：用自然语言告诉用户这个批注在说什么、建议做什么修改

3. **提供操作选项**：
   - **[Y] 应用修改** → 用户提供替换用的 Markdown 内容 → 执行 `replace` 命令替换对应章节
   - **[S] 跳过** → 保留批注，进入下一条
   - **[A] 应用并删除批注** → 先替换章节，再执行 `delete-comment` 删除已处理的批注
   - **[D] 仅删除批注** → 直接执行 `delete-comment` 删除批注但不改内容
   - **[Q] 退出** → 保存当前进度，结束审核

**Step 3: 应用修改时的格式处理**

当用户选择 Y 或 A 时，Claude 应该参考该章节现有的格式（从 format-report 获取），如果用户不指定新格式，用该章节当前的格式替换：

```bash
# 先查看该章节的格式
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py format-report path.docx

# 替换章节（保持现有格式）
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py replace path.docx "2.1" \
  --content "更新后的内容" \
  --output output.docx
```

**Step 4: 删除批注**（选择 A 或 D 时）

```bash
# 删除单条
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py delete-comment path.docx "3"

# 或一次性删除所有（审核完成后）
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py delete-all-comments path.docx
```

**Step 5: 全部批注处理完成后**

询问用户是否保存并结束审核。如果需要保留文档，确保保存。

---

## 通篇修改格式

当用户要求"统一格式"、"通篇改格式"、"所有正文改成 X" 时使用。

### 命令格式

```bash
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx \
  --target <目标> \
  --font-name <字体> \
  --font-ea <中文字体> \
  --font-size <字号> \
  --bold \
  --alignment <对齐> \
  --line-spacing <行距> \
  --first-line-indent <缩进> \
  --space-after <段后>
```

### target 选项

| 值 | 含义 |
|------|--------|
| `all` | 全部段落（正文 + 标题） |
| `body` | 仅正文段落（非标题） |
| `heading` | 仅标题段落 |
| `"2.1"` / `"1"` | 指定章节号 |

### 示例

```bash
# 全文正文改为黑体 14pt
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx \
  --target body --font-name 黑体 --font-size 14

# 所有标题加粗居中
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx \
  --target heading --bold --alignment CENTER

# 指定章节 2.1 改行距
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx \
  --target "2.1" --line-spacing 1.5 --font-size 12

# 全文统一
PYTHONIOENCODING=utf-8 py scripts/analyze_docx.py apply-format path.docx \
  --target all --font-name "Times New Roman" --font-ea 宋体 --font-size 12
```

---

## Dependencies

### Python 包
- `python-docx` — 核心 DOCX 读写能力（已安装）
- `pypandoc` — LaTeX 公式转 OMML（可选，`pip install pypandoc`，需系统安装 pandoc）
- `pywin32` — Win32 COM 自动化（可选，仅修订模式需要）
- `Pillow` — 图片处理（可选）
- `markitdown` — 文本验证（可选，`pip install markitdown`）

### 外部工具
- `@mermaid-js/mermaid-cli`（`mmdc`）— Mermaid 图渲染（可选，`npm install -g @mermaid-js/mermaid-cli`）
- Microsoft Word — 修订模式需要（可选）

### 环境检查
```bash
# 检查 Word 是否可用（返回 True/False）
py -c "import sys; sys.path.insert(0, 'E:/demo/docx'); from docx_editor.win32_ops import Win32Ops; print(Win32Ops.is_word_available())"
```

---

## Triggers

当用户提到以下内容时使用此 skill：
- `.docx` 文件、Word 文档、报告排版
- "帮我写一个文档" / "新建一个 word 文档" / "帮我生成报告"
- "修改第 X 章" / "更新文档中的某个章节" / "替换章节内容"
- "用修订模式修改" / "保留修改痕迹" / "审阅模式"
- "看看这个文档的批注" / "处理批注" / "审核批注" / "一条一条处理意见"
- "统一格式" / "通篇改成 X" / "批量改格式" / "把所有正文改成 X"
- 任何涉及 .docx 文件内容的操作
