"""
===============================================================================
脚本名称：修复交叉引用格式 (fix_cross_ref.py)
作者: ZGQ
功能概述：
    本脚本用于自动化处理 Word/WPS 检测报告中的交叉引用格式，
    解决手动调整耗时且易漏项的问题，保证每次的格式一致性。

核心工作流：
    1. 环境检测：抓取当前处于激活状态的 Word/WPS 文档。
    2. 遍历所有域代码, 定位交叉引用 (REF 域) 。
    3. 检查每个交叉引用是否已包含保留格式的开关 (\* MERGEFORMAT) 。
    4. 对于缺失该开关的交叉引用，自动追加 \* MERGEFORMAT 开关以确保格式稳定。
    5. 结果汇总：展示处理总数，完成修复闭环

前置依赖：
    - 运行前必须打开目标文档。
    - 同级目录下需存在 `file_utils_backup.py` 模块。
===============================================================================
"""
import os
import sys
import win32com.client
import pythoncom

# 【挂载外部模块备份文件】
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document

def update_cross_references():
    """
    接管当前活动的 Word/WPS 文档，遍历并修复交叉引用的格式开关
    """
    # 1. 尝试捕获当前正在运行的文字处理程序
    try:
        # 绝大多数情况下，WPS 在底层注册的 COM 名称也是 "Word.Application" 以保持兼容
        app = win32com.client.GetActiveObject("Word.Application")
    except pythoncom.com_error:
        try:
            # 备用方案：显式寻找 WPS 进程
            app = win32com.client.GetActiveObject("KWPS.Application")
        except pythoncom.com_error:
            print("【阻断】未能检测到正在运行的 Word/WPS 进程。请先打开一份报告。")
            return

    try:
        # 关闭屏幕更新以提升遍历速度
        app.ScreenUpdating = False
        
        doc = app.ActiveDocument
        fields = doc.Fields
        count = 0

        # 2. 遍历所有域代码 (注意：COM 对象的索引从 1 开始，而不是 0)
        for i in range(1, fields.Count + 1):
            f = fields.Item(i)
            
            # 3 代表 wdFieldRef，即交叉引用/引用域
            if f.Type == 3:
                code_text = f.Code.Text
                
                # 检查是否已经存在 \* MERGEFORMAT
                if "\\* MERGEFORMAT" not in code_text.upper():
                    # 追加保留格式的开关
                    f.Code.Text = code_text + " \\* MERGEFORMAT"
                    count += 1
                    
        print(f"【成功】处理完毕！共为 {count} 个交叉引用追加了保留格式开关。")

    except Exception as e:
        print(f"【异常】执行过程中出错: {e}")
    finally:
        # 强制恢复屏幕更新，防止文档卡死假死
        app.ScreenUpdating = True

if __name__ == "__main__":
    update_cross_references()