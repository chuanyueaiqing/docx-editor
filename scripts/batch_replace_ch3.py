#!/usr/bin/env python
"""批量替换第三章全部小节，一次性生成修订标记。"""
import os, sys, tempfile
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _PROJECT)

from docx_editor.document import DocxDocument
from docx_editor.win32_ops import Win32Ops

INPUT  = r"D:\360MoveData\Users\1\Desktop\小川\11.开发计划-0703.docx"
OUTPUT = r"D:\360MoveData\Users\1\Desktop\小川\11.开发计划-0703-修订版.docx"

CHAPTERS = ["3.1", "3.2", "3.3", "3.4", "3.5", "3.6", "3.7"]

sections = {
    "3.1": """# 3.1  DE-工具需求与约束

DE-系统面向射频/微波器件生产测试领域，其需求与约束从功能、性能和适配三个维度进行定义。

功能需求方面，系统须提供以下十项核心功能模块（各模块的详细定义及验收标准见需求规格说明书第4.3章）：

- **良率分析（GN-YA-001）** — 支持CPK/PPK等统计过程控制分析，具备按批次、机台、产品型号等多维度交叉分析能力，以图表形式展示良率趋势与分布状况。
- **数据过滤与清洗（GN-DC-002）** — 提供去重、表变换、数据合并、异常值替换和删除五种数据清洗手段，支持清洗规则的组合配置与批量执行。
- **数据分析（GN-DH-001）** — 支持直流参数与射频参数等多维数据的交互式分析，涵盖关联分析、统计分析和回归分析，分析结果可导出为报告。
- **晶圆缺陷检测（GN-WM-001）** — 支持Wafer Map整体及局部放大显示，提供良率热力图和缺陷模式的自动分类与标记功能。
- **脚本管理（GN-SC-001）** — 支持Parse、De-Embedding、Extract、Transform四类Python脚本的在线编辑、语法检查和版本管理，按角色控制脚本的查看和修改权限。
- **数据上传（GN-DC-001）** — 支持测试机台数据的实时采集与批量导入，兼容Excel、CSV、SNP等常见文件格式，同时提供手动录入界面作为补充。
- **数据可视化（GN-DV-001）** — 提供Dashboard综合看板，支持Smith圆图、极坐标图、散点图、箱型图等专业图表类型，图表配置支持保存为模板。
- **用户权限管理（GN-AUTH-001）** — 基于统一认证系统，设置工程师和管理员两种角色，实现功能权限和数据权限的精细化控制。
- **外部接口（GN-IF-001）** — 提供测试设备控制接口和标准REST API接口，满足与产线MES系统的数据对接需求。
- **DataManager数据管理（GN-DM-001）** — 覆盖Wafer级和DUT级数据的上传、加工、下载和删除全生命周期管理，操作记录可追溯。

性能约束方面：系统界面操作响应时间不超过30秒；支持至少50个并发用户同时访问；支持百万级测试数据的高效存储和查询，其中良率分析等核心查询响应时间不超过10秒。

适配约束方面：系统须适配麒麟操作系统和统信操作系统；须适配海光四代及以上和飞腾CPU；系统数据存储容量须支持1万片晶圆测试数据（约10TB至100TB）。""",

    "3.2": """# 3.2  项目文档需求与约束

本项目在开发过程中须按阶段产出相应的项目文档。各文档的编制应遵循GB/T 8567-2006《计算机软件文档编制规范》的相关要求，并在对应的阶段评审节点前完成评审归档。

本项目须产出的文档清单如下：

| 文档类别 | 文档名称 | 计划完成时间 |
|---------|---------|------------|
| 策划类 | 开发计划 | 需求分析阶段完成 |
| 需求类 | 需求规格说明书 | 已完成（V4.0） |
| 设计类 | 概要设计说明书 | 设计阶段完成 |
| 设计类 | 详细设计说明书 | 设计阶段完成 |
| 测试类 | 软件测试计划 | 设计阶段完成 |
| 测试类 | 软件测试报告 | 测试阶段完成 |
| 交付类 | 用户手册 | 安装移交阶段完成 |
| 交付类 | 维护手册 | 安装移交阶段完成 |""",

    "3.3": """# 3.3  工程策略与约束

为确保项目高效推进并满足国产化部署要求，本项目的工程实施遵循以下策略：

- 采用**混合开发模式**，以迭代方式推进功能开发。每个迭代周期约4周，结合阶段评审节点进行里程碑验收，确保开发过程的可控性和透明度。迭代计划应根据优先级动态调整，关键路径上的任务优先排期。
- **优先保障核心功能**的稳定实现。核心功能涵盖数据上传、脚本管理和数据分析可视化三个关键领域，在开发排期和资源分配上予以优先保障，非核心功能可在后续迭代中逐步完善。
- **国产化适配须尽早开展**。针对麒麟/统信操作系统和海光/飞腾CPU的适配工作应在开发环境建立阶段完成基础验证，降低后期集成的技术风险。适配过程中如遇到平台差异，应及时记录并在架构层面预留兼容方案。""",

    "3.4": """# 3.4  进度安排与资源约束

项目的进度安排与资源配置如下：

- 项目整体开发周期约**34个月**，计划自2025年2月启动，至2027年12月完成验收交付。
- 核心研发团队共计**27人**，涵盖前端开发、后端开发、算法工程、软件测试、质量保证和配置管理各专业岗位，各阶段人员投入可根据实际工作需要动态调整。
- 关键资源需求包括：支持国产CPU的开发测试服务器（麒麟/统信双操作系统环境）、StarRocks分布式数据库集群（不少于3节点）、Redis缓存服务器，以及配套的CI/CD持续集成环境。上述资源应在开发环境建立阶段落实到位。""",

    "3.5": """# 3.5  配置管理与质量保证约束

项目的配置管理与质量保证工作遵循以下要求：

- 所有代码和文档须纳入配置管理，使用Git进行版本控制。代码库按模块实行分支管理（main/develop/feature/hotfix），文档按版本归档。分支合并须经过代码评审，确保代码质量。
- 每个里程碑节点须通过内部评审后方可进入下一阶段。评审内容涵盖工作产品的完整性、一致性和规范性，评审发现的问题应记录并跟踪至关闭，评审记录须归档备查。
- 关键功能须经过**集成测试**和**合格性测试**两级验证。集成测试验证模块间接口和功能交互的正确性；合格性测试在国产化目标环境下完整执行，确保系统满足全部功能和性能需求。测试发现的问题按照严重等级分级管理，明确修复时限和验证标准。""",

    "3.6": """# 3.6  安装与移交约束

系统的安装部署与移交工作须满足以下要求：

- 系统须在用户生产现场完成安装部署和数据迁移。部署方案应包含数据迁移策略、回退预案和验收标准，确保生产环境的平稳过渡。部署前应完成环境检查和预演，降低现场风险。
- 移交前须完成用户培训，提供用户手册和维护手册。培训应覆盖工程师和管理员两类角色，内容包含日常操作、常见故障处理和基本维护操作，确保用户能够独立完成日常工作。
- 移交后提供不少于**3年**的技术保障服务。服务内容包括系统故障处理、版本升级、安全补丁和技术咨询，保障系统长期稳定运行。服务响应时间应按故障等级设立明确标准。""",

    "3.7": """# 3.7  保密性约束

系统须满足以下保密性要求：

- 用户密码须采用加密方式存储，不得以明文形式保存在数据库或日志文件中。密码策略应包含最小长度、复杂度要求和定期更换机制。
- 敏感数据在传输过程中须使用加密协议（如HTTPS/TLS），敏感数据在存储时宜采用加密存储，防止数据泄露。
- 系统应具备访问日志审计功能，记录所有敏感数据的访问和操作行为（包括操作人、时间、操作类型和结果），日志应不可篡改并保留不少于6个月，确保可追溯。""",
}


def accept_all_tracked_changes(input_path: str) -> str:
    clean_path = tempfile.NamedTemporaryFile(suffix='.docx', delete=False).name
    print("正在通过 Word 接受文档中所有既有修订...")
    with Win32Ops() as ops:
        ops.open_document(input_path)
        ops.word.ActiveDocument.AcceptAll()
        ops.wd_doc = ops.word.ActiveDocument
        ops.save_document(clean_path)
    print("  既有修订已全部接受，保存为干净基线版本。")
    return clean_path


# ===== Main =====

# Step 0: Handle existing tracked changes
import zipfile, xml.etree.ElementTree as ET
has_existing_tracked = False
with zipfile.ZipFile(INPUT) as z:
    if 'word/document.xml' in z.namelist():
        xml_content = z.read('word/document.xml')
        root = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        ins_count = len(root.findall('.//w:ins', ns))
        del_count = len(root.findall('.//w:del', ns))
        if ins_count > 0 or del_count > 0:
            has_existing_tracked = True
            print(f"检测到原文档有 {ins_count} 处插入和 {del_count} 处删除的既有修订。")

working_input = INPUT
if has_existing_tracked:
    working_input = accept_all_tracked_changes(INPUT)
else:
    print("原文档无既有修订，直接使用原文档。")

# Step 1: Load document and replace all chapters
print("正在加载文档...")
doc = DocxDocument(working_input, use_track_changes=False)

for ch_num in CHAPTERS:
    print(f"  替换章节 {ch_num} ...")
    doc.replace_chapter(ch_num, sections[ch_num])

# Step 2: Save modified version
modified_tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
modified_tmp.close()
doc.document.save(modified_tmp.name)
print("修改版已保存到临时文件。")

# Step 3: Word COM comparison
print("正在启动 Word 进行修订对比...")
with Win32Ops() as ops:
    ops.open_document(working_input)
    ops.word.ActiveDocument.Compare(
        Name=modified_tmp.name,
        CompareTarget=2,
        IgnoreAllComparisonWarnings=True,
    )
    ops.wd_doc = ops.word.ActiveDocument
    ops.save_document(OUTPUT)

# Step 4: Cleanup
os.unlink(modified_tmp.name)
if has_existing_tracked and working_input != INPUT:
    os.unlink(working_input)

print(f"\n完成！输出文件：{OUTPUT}")
print("请用 Word 打开查看修订标记。")
