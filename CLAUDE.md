# E:\demo\docx — docx 编辑器项目

## 🚨 项目规则：所有代码操作必须使用 CodeGraph

本项目已配置 CodeGraph MCP 服务器（`codegraph_*` 工具）。**任何代码搜索、代码读取、代码分析操作都必须通过 CodeGraph 完成。**

### 禁止做法 vs 正确做法

| 你需要的操作 | ✅ 正确工具 | ❌ 禁止做法 |
|-------------|-----------|-----------|
| 查找符号定义 | `codegraph_search` | Grep / Bash(grep) |
| 理解模块上下文 | `codegraph_context` | 先 search 再 node 链式调用 |
| 查看调用关系 | `codegraph_callers` / `codegraph_callees` | Grep -r 搜函数名 |
| 看多个相关符号源码 | `codegraph_explore` | 一个一个 codegraph_node / Read |
| 看单个符号签名 | `codegraph_node` | Read 整个文件 |
| 看项目文件结构 | `codegraph_files` | Glob / Bash(ls) |
| 确认索引状态 | `codegraph_status` | ls .codegraph/ |

### 例外（仅限以下 3 种）

1. **纯文本搜索** — 搜注释、日志、字符串字面量、Markdown 文档中**不在 AST 中的文本**
2. **已读文件回看** — 当前对话中已经 Read 过的文件，需要再看一次
3. **CodeGraph 索引不存在** — 提示用户初始化

### 搜索结果信任

CodeGraph 基于完整 AST 解析，结果必须信任。**禁止再用 Read / Grep 去验证 CodeGraph 的结果。** 不信任就升级给更高级模型判断，不要自己浪费 token 验证。

---

## 项目结构

```
E:/demo/docx/
├── docx_editor/       # 核心库包
│   ├── creator.py     # DOCX 创建
│   ├── document.py    # 文档操作
│   ├── models.py      # 数据模型
│   ├── utils.py       # 工具函数
│   ├── win32_ops.py   # Windows Word 操作
│   ├── wps_ops.py     # WPS 操作
│   ├── equation_ops.py    # 公式操作
│   ├── font_utils.py      # 字体工具
│   ├── latex_to_unicodemath.py  # LaTeX→UnicodeMath 转换
│   ├── omml_builder.py    # OMML 构建器
│   └── pandoc_omml.py    # Pandoc OMML 处理
├── tests/             # 测试
├── scripts/           # 脚本
├── skills/            # Claude Code skills
│   └── docx-mode/     # docx-mode skill
└── output/            # 输出文件
```

---

---

## Bash 执行规则：必须合并，禁止拆分

连续执行的 Bash 命令必须合并为一次调用，禁止拆开。

### ❌ 错误做法

```
cd "E:/demo/docx"       ← 一次 API 调用
py -c "print('hello')"  ← 又一次 API 调用
```

### ✅ 正确做法

```
cd "E:/demo/docx" && py -c "print('hello')"   ← 一次搞定
```

### 合并三原则

1. **路径合并** — 用 `cd X && cmd`，不单独 cd
2. **脚本化** — 3+ 步 Python → 写临时 .py 文件 `py script.py` 一次性执行
3. **&& 串联** — `cmd1 && cmd2 && cmd3` 全连起来

**合并 Bash = 直接减少 API 调用次数。** 今天 73 次单独的 cd = 73 次浪费的调用。

---

---

## 上下文成本意识——何时继续、何时 clear

当对话过长或你提到要 `/clear` 时，我会自动运行成本分析：

```
py ~/.claude/scripts/context_advisor.py
```

判断标准：**当前上下文 < 系统提示词×10 → 继续；>= 系统提示词×10 → clear。**
系统提示词大小由脚本动态检测。**没到阈值就别 clear，继续在当前对话里问。**

---

## 平台注意事项

- 运行在 **Windows 10 + Git Bash + Python 3.14**
- 始终用 `py` 命令代替 `python`（`python` 不可用）
- Windows 终端 GBK 编码，CLI 脚本输出避免 emoji
- 文件路径用正斜杠 `E:/demo/docx/`，勿用反斜杠
