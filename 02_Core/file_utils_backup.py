"""
===============================================================================
脚本名称：通用文档备份工具 (file_utils_backup.py)
作者: ZGQ
功能概述：
    本脚本为当前激活的 Word/WPS 文档提供安全备份功能。
    在执行正文排版前，先校验文档是否已保存并具有有效路径，
    然后保存当前文档、生成带时间戳的备份副本，并静默保存该副本，
    以避免原文件直接被修改或损坏。

核心工作流：
    1. 获取当前活动文档并读取路径、完整名称和文件名。
    2. 校验文档是否已保存到本地磁盘，避免对未落盘文档执行备份。
    3. 保存当前文档当前状态。
    4. 构建带时间戳的备份文件路径。
    5. 创建新文档副本并保存，然后关闭副本。
    6. 返回备份是否成功。

前置依赖：
    - 需要可访问的 Word 或 WPS 应用对象。
    - 依赖 win32com.client 提供 COM 自动化接口。
===============================================================================
"""
import os
import time

def backup_current_document(app):
    try:
        doc = app.ActiveDocument
        
        # 安全读取属性，拦截 COM 底层异常
        try:
            doc_path = doc.Path
            doc_fullname = doc.FullName
            doc_name = doc.Name
        except Exception:
            doc_path = ""
            doc_fullname = ""
            doc_name = ""

        # 校验物理文件落盘状态
        if not doc_path or doc_fullname == doc_name:
            print("【备份阻断】源文件尚未执行本地存储。")
            return False
        
        # 1. 强制写入当前进度
        doc.Save()
        
        # 2. 构建副本物理地址 
        base, ext = os.path.splitext(doc_fullname)
        timestamp = time.strftime("%Y%m%d_%H时%M分")
        new_path = os.path.abspath(f"{base}_{timestamp}{ext}")
        
        # 3. 后台静默克隆并落盘 (不干涉 ActiveDocument 的焦点)
        backup_doc = app.Documents.Add(Template=doc_fullname)
        backup_doc.SaveAs2(new_path)
        backup_doc.Close(0)
        
        print(f"【系统】文件安全克隆完毕: {new_path}")
        return True
        
    except Exception as e:
        print(f"【系统】克隆进程抛出异常: {e}")
        return False