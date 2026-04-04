"""
===============================================================================
脚本名称：通用文档备份工具 (file_utils_backup.py)
作者: ZGQ
功能概述：
    本脚本为当前激活的 Word/WPS 文档提供安全备份功能。
    在执行正文排版前，先校验文档是否已保存并具有有效路径，
    然后保存当前文档、生成带时间戳的备份副本，并静默保存该副本，
    以避免原文件直接被修改或损坏。

    这个脚本是用来备份Word文档的工具。在修改文档之前，先创建一个带时间戳的备份副本，
    这样如果修改过程中出现问题，可以恢复到原始状态。适合新手使用，自动处理备份逻辑。

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
import os  # 用于文件路径操作
import time  # 用于生成时间戳

def backup_current_document(app):
    """
    备份当前活动文档。

    为当前打开的Word文档创建一个带时间戳的备份副本。

    参数:
        app: Word或WPS应用程序对象。

    返回:
        bool: 备份是否成功。
    """
    try:
        # 获取当前活动文档
        doc = app.ActiveDocument
        
        # 安全读取文档属性，防止COM异常
        try:
            doc_path = doc.Path  # 文档所在文件夹路径
            doc_fullname = doc.FullName  # 完整文件路径
            doc_name = doc.Name  # 文件名
        except Exception:
            # 如果读取失败，设置为空值
            doc_path = ""
            doc_fullname = ""
            doc_name = ""

        # 检查文档是否已保存到本地磁盘
        if not doc_path or doc_fullname == doc_name:
            print("【备份阻断】源文件尚未执行本地存储。")
            return False  # 未保存的文档不备份
        
        # 1. 保存当前文档的最新状态
        doc.Save()
        
        # 2. 构建备份文件路径
        base, ext = os.path.splitext(doc_fullname)  # 分离文件名和扩展名
        timestamp = time.strftime("%Y%m%d_%H时%M分")  # 生成时间戳
        new_path = os.path.abspath(f"{base}_{timestamp}{ext}")  # 新备份文件路径
        
        # 3. 创建文档副本并保存（不影响当前活动文档）
        backup_doc = app.Documents.Add(Template=doc_fullname)  # 基于原文档创建新副本
        backup_doc.SaveAs2(new_path)  # 保存副本到新路径
        backup_doc.Close(0)  # 关闭副本文档（0表示不保存）
        
        print(f"【系统】文件安全克隆完毕: {new_path}")
        return True  # 备份成功
        
    except Exception as e:
        # 如果出现异常，打印错误信息
        print(f"【系统】克隆进程抛出异常: {e}")
        return False  # 备份失败