#!/usr/bin/env python
"""批量替换第三章全部小节（优化版），一次性生成修订标记。"""
import os, sys, tempfile, zipfile, xml.etree.ElementTree as ET
_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, _PROJECT)

from docx_editor.document import DocxDocument
from docx_editor.win32_ops import Win32Ops

INPUT  = r"D:\360MoveData\Users\1\Desktop\小川\11.开发计划-0703.docx"
OUTPUT = r"D:\360MoveData\Users\1\Desktop\小川\11.开发计划-0703-修订版-优化.docx"

CHAPTERS = ["3.1", "3.2", "3.3", "3.4", "3.5", "3.6", "3.7"]

sections = {
    "3.1": """# 3.1  DE-工具需求与约束

DE-系统面向射频/微波器件生产测试领域，其需求与约束从功能、性能和适配三个维度进行定义。

功能需求方面，系统须提供以下十一项核心功能模块（各模块的详细定义及验收标准见需求规格说明书第 4.3 章）：

- 良率分析（GN-YA-001）— 支持 CPK/PPK 等统计过程控制分析，具备按批次、机台、产品型号等多维度交叉分析能力，以图表形式展示良率趋势和分布状况。
- 数据过滤与清洗（GN-DC-002）— 提供去重、表变换、数据合并、异常值替换和删除五种数据清洗手段，支持清洗规则的组合配置与批量执行。
- 数据分析（GN-DH-001）— 支持直流参数与射频参数等多维数据的交互式分析，涵盖关联分析、统计分析和回归分析，分析结果可导出为报告。
- 晶圆缺陷检测（GN-WM-001）— 支持 Wafer Map 整体及局部放大显示，提供良率热力图和缺陷模式的自动分类与标记功能。
- 脚本管理（GN-SC-001）— 支持 Parse、De-Embedding、Extract、Transform 四类 Python 脚本的在线编辑、语法检查和版本管理，按角色控制脚本的查看和修改权限。
- 数据上传（GN-DC-001）— 支持测试机台数据的实时采集与批量导入，兼容 Excel、CSV、SNP 等常见文件格式，同时提供手动录入界面作为补充。
- 数据可视化（GN-DV-001）— 提供 Dashboard 综合看板，支持 Smith 圆图、极坐标图、散点图、箱型图等专业图表类型，图表配置支持保存为模板。
- 用户权限管理（GN-AUTH-001）— 基于统一认证系统，设置工程师和管理员两种角色，实现功能权限和数据权限的精细化控制。
- 外部接口（GN-IF-001）— 提供测试设备控制接口和标准 REST API 接口，满足与产线 MES 系统的数据对接需求。
- DataManager 数据管理（GN-DM-001）— 覆盖 Wafer 级和 DUT 级数据的上传、加工、下载和删除全生命周期管理，操作记录可追溯。
- 系统配置管理（GN-CFG-001）— 提供系统运行参数的集中配置界面，支持测试项目配置、模板管理和系统运行参数的动态调整。

上述十一项功能模块中，前九项为需求规格说明书已定义的核心功能，后两项为本次策划新增补充模块，旨在完善数据生命周期管理和系统可维护性。

性能约束方面：系统界面操作响应时间不超过 30 秒；支持至少 50 个并发用户同时访问；支持百万级测试数据的高效存储和查询，其中良率分析等核心查询响应时间不超过 10 秒；数据批量导入速率不低于 10MB/s。

适配约束方面：系统须适配麒麟操作系统和统信操作系统；须适配海光四代及以上和飞腾 CPU；系统数据存储容量须支持 1 万片晶圆测试数据（约 10TB ～ 100TB），且须支持在线扩容。""",

    "3.2": """# 3.2  项目文档需求与约束

本项目在开发过程中须按阶段产出相应的项目文档。各文档的编制应遵循 GB/T 8567-2006《计算机软件文档编制规范》的相关要求，并在对应的阶段评审节点完成评审归档。

本项目须产出的文档如下表所示：

| 文档类别 | 文档名称 | 计划完成阶段 |
|---------|---------|------------|
| 策划类 | 开发计划 | 需求分析阶段完成 |
| 需求类 | 需求规格说明书 | 已完成（V4.0） |
| 设计类 | 概要设计说明书 | 设计阶段完成 |
| 设计类 | 详细设计说明书 | 设计阶段完成 |
| 测试类 | 软件测试计划 | 设计阶段完成 |
| 测试类 | 软件测试报告 | 测试阶段完成 |
| 交付类 | 用户手册 | 安装移交阶段完成 |
| 交付类 | 维护手册 | 安装移交阶段完成 |

各文档在编制过程中须遵循项目配置管理要求，文档模板和评审检查单由质量保证组统一提供，确保文档格式和内容的规范性。文档评审未通过的，须按评审意见完成整改后方可进入下一阶段。""",

    "3.3": """# 3.3  工程策略与约束

为确保项目高效推进并满足国产化部署要求，本项目的工程实施遵循以下策略：

- 采用混合开发模式，以迭代方式推进功能开发。每个迭代周期约 4 周，结合阶段评审节点进行里程碑验收，确保开发过程的可控性和透明度。每个迭代结束时须完成迭代总结会议，评审当次迭代成果并调整后续计划。
- 优先保障核心功能的稳定实现。核心功能涵盖数据上传、脚本管理和数据分析可视化三个关键领域，在开发排期和资源分配上予以优先保障。核心功能的单元测试覆盖率须达到 80% 以上。
- 国产化适配须在开发环境建立阶段完成基础验证。针对麒麟/统信操作系统和海光/飞腾 CPU 的适配工作应尽早开展，在开发环境建立阶段完成操作系统兼容性验证和基础功能适配测试，降低后期集成的技术风险。
- 持续集成与持续部署（CI/CD）管道须在开发环境建立阶段同步搭建，确保每次代码提交均能自动触发构建、静态检查和单元测试，及时发现集成问题。""",

    "3.4": """# 3.4  进度安排与资源约束

项目的进度安排与资源配置如下：

- 项目整体开发周期约为 35 个月，计划自 2025 年 2 月启动，至 2027 年 12 月完成验收交付。根据阶段划分设置 6 个里程碑节点：开发环境就绪、需求评审通过、设计评审通过、核心功能交付、系统测试完成、验收交付。
- 核心研发团队共计 27 人，涵盖前端开发、后端开发、算法工程、软件测试、质量保证和配置管理各专业岗位，确保各阶段工作有序衔接。各岗位人员根据阶段需求动态调配，资源紧缺时优先保障核心功能开发。
- 关键资源需求包括：支持国产 CPU 的开发测试服务器（麒麟/统信双操作系统环境）、StarRocks 分布式数据库集群（不少于 3 节点）、Redis 缓存服务器，以及配套的 CI/CD 持续集成环境。以上资源须在项目启动后 2 个月内到位，以保障开发环境建立阶段的顺利推进。""",

    "3.5": """# 3.5  配置管理与质量保证约束

项目的配置管理与质量保证工作遵循以下要求：

- 所有代码和文档须纳入配置管理，使用 Git 进行版本控制。代码库按模块实行分支管理（main / develop / feature / hotfix），文档按版本归档并建立版本变更记录。配置项的标识、变更、状态记录和审计按项目配置管理计划执行。
- 每个里程碑节点须通过内部评审后方可进入下一阶段。评审内容涵盖工作产品的完整性、一致性和规范性，评审记录须归档备查。对评审发现的问题，须建立问题跟踪表，明确责任人、整改措施和完成期限，整改完成后由质量保证组进行闭环确认。
- 关键功能须经过集成测试和合格性测试两级验证。集成测试验证模块间接口和功能交互的正确性，合格性测试在国产化目标环境下完整执行，确保系统满足全部功能和性能需求。测试用例须覆盖正常流程、异常流程和边界条件，测试结果须形成测试报告并归档。
- 质量保证组独立于开发团队，直接向项目经理汇报。质量保证组定期对项目过程和产品进行审计，审计结果纳入项目绩效评估。""",

    "3.6": """# 3.6  安装与移交约束

系统的安装部署与移交工作须满足以下要求：

- 系统须在用户生产现场完成安装部署和数据迁移。部署方案应包含数据迁移策略、回退预案和验收标准，确保生产环境的平稳过渡。部署前须与用户确认现场环境条件，包括网络配置、操作系统版本、硬件资源等是否满足部署要求。
- 移交前须完成用户培训，提供用户手册和维护手册。培训应覆盖工程师和管理员两类角色，确保用户能够独立完成日常操作和基本维护。培训结束后须进行培训效果评估，对未达标的内容安排补充培训。
- 移交后提供不少于 3 年的技术保障服务。服务内容包括系统故障处理、版本升级、安全补丁和技术咨询，保障系统长期稳定运行。技术保障服务须明确服务等级协议（SLA），约定故障响应时间和解决时限。""",

    "3.7": """# 3.7  保密性约束

系统须满足以下保密性要求：

- 用户密码须采用加密方式存储（推荐使用 bcrypt 或 Argon2 算法），不得以明文形式保存在数据库或日志文件中。
- 敏感数据在传输过程中须使用加密协议（如 HTTPS/TLS 1.2 及以上版本），防止数据在传输过程中被窃听或篡改。
- 系统应具备访问日志审计功能，记录所有敏感数据的访问和操作行为，包括操作人员、操作时间、操作类型和操作结果，确保可追溯。审计日志保存期限不少于 12 个月，且不得被普通用户修改或删除。
- 数据库访问须遵循最小权限原则，不同角色的数据库账号仅赋予其业务所必需的数据操作权限，避免越权访问。
- 系统的安全设计须通过安全评审后方可上线部署，安全评审由质量保证组组织，必要时引入第三方安全评测机构。""",
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
