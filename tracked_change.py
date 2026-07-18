"""
使用 python-docx + win32com 实现修订模式写入
在文档开头插入文字"我是陈小川"，并记录为修订标记
"""
import os
import sys
from docx import Document
import win32com.client
import pythoncom


def main():
    # 输出路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output")
    output_file = os.path.join(output_dir, "修订测试.docx")

    # ========== Step 1: python-docx 创建基础文档 ==========
    print("[1/3] python-docx 创建基础文档...")
    os.makedirs(output_dir, exist_ok=True)

    doc = Document()
    # 添加一个空白段落，作为文档的初始内容骨架
    doc.add_paragraph("")
    doc.save(output_file)
    print(f"       -> 已保存: {output_file}")

    # ========== Step 2: win32com 启用修订并写入 ==========
    print("[2/3] win32com 启用修订模式并插入文字...")

    pythoncom.CoInitialize()  # COM 初始化

    word = None
    wd_doc = None

    try:
        # 启动 Word（后台运行，不显示窗口）
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        # 打开 python-docx 创建的文档
        abs_path = os.path.abspath(output_file)
        wd_doc = word.Documents.Open(abs_path)

        # 启用修订模式 —— 必须在做任何修改之前设置
        wd_doc.TrackRevisions = True

        # 定位到文档开头（wdStory = 6）
        word.Selection.HomeKey(Unit=6)

        # 插入文字（此时 TrackRevisions=True，插入操作会被记录为修订）
        word.Selection.TypeText("我是陈小川")

        # 保存（修订标记自动保留在 docx 中）
        wd_doc.Save()
        print("       -> 修订文字已写入")

    except Exception as e:
        print(f"[错误] {e}", file=sys.stderr)
        sys.exit(1)

    finally:
        # ========== Step 3: 清理 Word 进程 ==========
        print("[3/3] 清理 Word 进程...")
        if wd_doc is not None:
            try:
                wd_doc.Close()
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    print(f"\n完成！")
    print(f"   请用 Word 打开: {output_file}")
    print('   在 Word 中开启「审阅 -> 所有标记」查看修订痕迹')


if __name__ == "__main__":
    main()
