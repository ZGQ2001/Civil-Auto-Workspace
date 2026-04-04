"""
===============================================================================
脚本名称：修复交叉引用格式 (fix_cross_ref.py)
作者: ZGQ
功能概述：
    本脚本用于自动化处理 Word/WPS 检测报告中的交叉引用格式，
    解决手动调整耗时且易漏项的问题，保证每次的格式一致性。

    这个脚本是用来修复Word文档中交叉引用格式问题的工具。交叉引用是指文档中引用其他部分（如"见表1"）的链接。
    有时候这些引用在更新时会丢失格式，这个脚本会自动添加格式保护开关，确保引用保持正确的格式。

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
import os  # 用于路径操作
import sys  # 用于添加模块路径
import win32com.client  # 用于控制Word/WPS
import pythoncom  # COM组件初始化

# 【挂载外部模块备份文件】
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document  # 导入备份函数


def update_cross_references():
    """
    修复当前活动文档中的交叉引用格式。

    遍历文档中的所有域，找到交叉引用域，为缺失格式保护开关的引用添加 \* MERGEFORMAT 开关。
    """
    # 1. 尝试连接到正在运行的文字处理程序
    try:
        # 大多数情况下，WPS也使用"Word.Application"作为COM名称以保持兼容
        app = win32com.client.GetActiveObject("Word.Application")
    except pythoncom.com_error:
        try:
            # 备用方案：显式寻找WPS进程
            app = win32com.client.GetActiveObject("KWPS.Application")
        except pythoncom.com_error:
            print("【阻断】未能检测到正在运行的 Word/WPS 进程。请先打开一份报告。")
            return  # 如果都找不到，退出函数

    try:
        # 关闭屏幕更新以提高处理速度
        app.ScreenUpdating = False
        
        # 获取活动文档
        doc = app.ActiveDocument
        fields = doc.Fields  # 获取文档中的所有域
        count = 0  # 计数器，记录修复的引用数量

        # 2. 遍历所有域代码 (注意：COM对象的索引从1开始)
        for i in range(1, fields.Count + 1):
            f = fields.Item(i)
            
            # 3代表wdFieldRef，即交叉引用/引用域
            if f.Type == 3:
                code_text = f.Code.Text  # 获取域代码文本
                
                # 检查是否已经存在格式保护开关 \* MERGEFORMAT
                if "\\* MERGEFORMAT" not in code_text.upper():
                    # 如果没有，追加保留格式的开关
                    f.Code.Text = code_text + " \\* MERGEFORMAT"
                    count += 1  # 计数器加1
                    
        print(f"【成功】处理完毕！共为 {count} 个交叉引用追加了保留格式开关。")

    except Exception as e:
        # 如果出现异常，打印错误信息
        print(f"【异常】执行过程中出错: {e}")
    finally:
        # 无论成功还是失败，都要恢复屏幕更新，防止文档卡死
        app.ScreenUpdating = True


if __name__ == "__main__":
    """
    主程序入口。

    直接运行此脚本时，会自动修复当前打开文档中的交叉引用格式。
    """
    update_cross_references()  # 调用修复函数