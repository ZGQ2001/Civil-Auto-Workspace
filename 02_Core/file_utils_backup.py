"""
===============================================================================
脚本名称：通用文档/表格备份工具 (file_utils_backup.py)
作者: ZGQ
功能概述：
    为当前激活的 Word/WPS 或 Excel/ET 文档提供安全备份功能。
    自动识别应用程序类型（文字或表格），执行底层克隆备份。
===============================================================================
"""
import os
import time

def backup_current_document(app):
    try:
        # 1. 自动探测传入的 COM 对象类型
        is_word = True
        try:
            doc = app.ActiveDocument  # 尝试按 Word 处理
        except Exception:
            try:
                doc = app.ActiveWorkbook  # 尝试按 Excel 处理
                is_word = False
            except Exception:
                print("【备份阻断】无法识别当前的应用程序对象 (非Word/Excel)。")
                return False

        # 2. 安全读取文件路径属性
        try:
            doc_path = doc.Path
            doc_fullname = doc.FullName
            doc_name = doc.Name
        except Exception:
            doc_path = ""
            doc_fullname = ""
            doc_name = ""

        if not doc_path or doc_fullname == doc_name:
            print("【备份阻断】源文件尚未执行本地存储。")
            return False
        
        # 3. 保存源文件最新状态
        doc.Save()
        
        # 4. 构建备份路径
        base, ext = os.path.splitext(doc_fullname)
        timestamp = time.strftime("%Y%m%d_%H时%M分")
        backup_path = f"{base}_backup_{timestamp}{ext}"

        # 5. 根据软件类型执行不同的底层克隆逻辑
        if is_word:
            backup_doc = app.Documents.Add(doc_fullname)
            backup_doc.SaveAs2(backup_path)
            backup_doc.Close(0)  # 0 = wdDoNotSaveChanges
        else:
            # Excel 使用 SaveCopyAs，只备份不切换当前激活的工作簿
            doc.SaveCopyAs(backup_path)

        return True

    except Exception as e:
        print(f"【备份异常】底层执行出错: {e}")
        return False