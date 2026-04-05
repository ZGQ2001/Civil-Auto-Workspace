"""
===============================================================================
脚本名称：通用文档/表格备份工具 (file_utils_backup.py)
作者: ZGQ
功能概述：
    为传入的 Word/WPS 或 Excel/ET 内存对象提供安全备份功能。
    V3.0 修复版：废除 hasattr 探测，采用原生 Application.Name 显式鉴别。
===============================================================================
"""
import os
import time

def backup_current_document(target_obj):
    try:
        # 1. 安全读取文件路径属性
        try:
            doc_path = target_obj.Path
            doc_fullname = target_obj.FullName
            doc_name = target_obj.Name
        except Exception:
            doc_path = ""
            doc_fullname = ""
            doc_name = ""

        if not doc_path or doc_fullname == doc_name:
            print("【备份阻断】源文件尚未执行本地存储。")
            return False
        
        # 2. 保存源文件最新状态
        target_obj.Save()
        
        # 3. 构建备份路径
        base, ext = os.path.splitext(doc_fullname)
        timestamp = time.strftime("%Y%m%d_%H时%M分")
        backup_path = f"{base}_backup_{timestamp}{ext}"

        # 4. 【核心修复】：直接获取宿主程序的名称进行安全判定
        app_name = ""
        try:
            app_name = target_obj.Application.Name
        except Exception:
            pass

        # 匹配 Microsoft Excel, WPS 表格, ET 等环境
        if "Excel" in app_name or "表格" in app_name or "ET" in app_name:
            target_obj.SaveCopyAs(backup_path)
            
        # 默认按 Microsoft Word 或 WPS 文字环境处理
        else:
            app = target_obj.Application
            backup_doc = app.Documents.Add(doc_fullname)
            backup_doc.SaveAs2(backup_path)
            backup_doc.Close(0)  # 0 = wdDoNotSaveChanges

        return True

    except Exception as e:
        # 【新增兜底】：将真实的 COM 报错写入日志文件，方便后续排查
        error_log_path = os.path.join(os.path.dirname(__file__), "backup_error_log.txt")
        try:
            with open(error_log_path, "w", encoding="utf-8") as f:
                f.write(f"备份底层崩溃详情:\n{str(e)}")
        except:
            pass
            
        print(f"【备份异常】底层执行出错: {e}")
        return False