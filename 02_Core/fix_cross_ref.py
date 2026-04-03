import win32com.client
import pythoncom

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