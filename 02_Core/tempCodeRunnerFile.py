import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
import os
import time
import re
import win32com.client
import pythoncom

# [配置类 GlobalConfig 和 审计类 AuditLog 保持不变...]
class GlobalConfig:
    def __init__(self):
        self.chinese_font = "宋体"
        self.english_font = "Times New Roman"
        self.font_size = 12.0
        self.table_width_percent = 100
        self.skip_pages = []
        self.empty_cell_color = 255
        self.max_table_threshold = 100
        self.backup_suffix = "_备份_"
        self.title_format = {"align": 1, "space_before": 0.5, "space_after": 0, "line_spacing_rule": 0}

class AuditLog:
    def __init__(self):
        self.total = 0
        self.success = 0
        self.skipped = 0
        self.errors = 0
        self.empty_cells = []

config = GlobalConfig()
audit_log = AuditLog()

# ---------------- 1. 增强版交互模块 (UI) ----------------
def show_ui_and_get_params(file_name):
    """唤起弹窗，并将文件名显示在正文提示区"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True) 

    # 1. 标题保持简洁，文件名作为正文第一行显示
    # 使用 \n 换行符将文件名和输入提示分开
    prompt_base = f"当前文件：{file_name}\n\n"

    # 1/5：模板
    tpl_input = simpledialog.askstring(
        "排版参数配置 - 1/5", 
        f"{prompt_base}请选择排版模板：\n1 - 公文(仿宋) | 2 - 标准(宋体)", 
        initialvalue="1", 
        parent=root
    )
    if not tpl_input: return False
    
    # 后续步骤同理，把文件名放在 prompt 里
    if tpl_input == "1":
        config.chinese_font = "仿宋_GB2312"
        config.font_size = 10.5
    else:
        config.chinese_font = "宋体"
        config.font_size = 10.5

    # 2/5：字号
    size_input = simpledialog.askstring(
        "排版参数配置 - 2/5", 
        f"{prompt_base}请输入全局字号：", 
        initialvalue=str(config.font_size), 
        parent=root
    )
    if not size_input: return False
    config.font_size = float(size_input)

    # 3/5：宽度
    width_input = simpledialog.askstring(
        "排版参数配置 - 3/5", 
        f"{prompt_base}请输入表格宽度百分比(10-100)：", 
        initialvalue="95", 
        parent=root
    )
    if not width_input: return False
    config.table_width_percent = int(width_input)

    # 4/5：跳过页
    skip_input = simpledialog.askstring(
        "排版参数配置 - 4/5", 
        f"{prompt_base}请输入需跳过的页码（英文逗号分隔）：", 
        initialvalue="4", 
        parent=root
    )
    if skip_input and skip_input.strip():
        config.skip_pages = [int(p.strip()) for p in skip_input.split(",") if p.strip().isdigit()]

    # 5/5：间距
    space_input = simpledialog.askstring(
        "排版参数配置 - 5/5", 
        f"{prompt_base}请输入表名段前间距行数：", 
        initialvalue="0.5", 
        parent=root
    )
    if not space_input: return False
    config.title_format["space_before"] = float(space_input)

    root.destroy()
    return True

# [get_word_app, backup_document, process_all_tables 函数保持不变...]
def get_word_app():
    try: return win32com.client.GetActiveObject("Word.Application")
    except:
        try: return win32com.client.GetActiveObject("KWPS.Application")
        except: return None

def backup_document(app):
    try:
        doc = app.ActiveDocument
        if not doc.FullName or doc.FullName == doc.Name: return False
        doc.Save()
        base, ext = os.path.splitext(doc.FullName)
        new_path = os.path.abspath(f"{base}{config.backup_suffix}{int(time.time())}{ext}")
        backup_doc = app.Documents.Add(Template=doc.FullName)
        backup_doc.SaveAs2(new_path)
        backup_doc.Close(0)
        return True
    except: return False

def process_all_tables(app):
    try:
        doc = app.ActiveDocument
        tables = doc.Tables
        table_count = tables.Count
        audit_log.total = table_count
        
        if table_count == 0: return True

        # --- 新增：创建视觉进度条窗口 ---
        pg_root = tk.Tk()
        pg_root.title("三检所自动化引擎")
        pg_root.attributes('-topmost', True)
        pg_root.geometry("350x120")
        
        # 居中显示进度窗
        tk.Label(pg_root, text=f"正在处理：{doc.Name}", fg="blue").pack(pady=5)
        progress_label = tk.Label(pg_root, text="准备开始...")
        progress_label.pack()
        
        bar = ttk.Progressbar(pg_root, length=280, mode='determinate', maximum=table_count)
        bar.pack(pady=10)
        pg_root.update() # 强制渲染初始界面

        app.ScreenUpdating = False
        
        for i in range(1, table_count + 1):
            # 更新视觉进度条
            percent = int((i / table_count) * 100)
            bar['value'] = i
            progress_label.config(text=f"正在格式化第 {i}/{table_count} 个表格 ({percent}%)")
            pg_root.update() # 极其重要：强制刷新 UI 界面，否则会卡死

            try:
                tbl = tables.Item(i)
                page_num = tbl.Range.Information(3)
                
                # [中间的表名排版和表格排版逻辑完全不动，直接保留之前的...]
                # --- 此处省略原有的 tbl.PreferredWidthType 等排版代码 ---
                # (请确保你粘贴时，把之前那段 tbl 处理逻辑放在这里)
                
                audit_log.success += 1
            except Exception as e:
                audit_log.errors += 1
                audit_log.error_details.append(f"T{i} 崩溃: {e}")

        # 处理结束，销毁进度窗口
        pg_root.destroy()
        
        doc.Save()
        app.ScreenUpdating = True
        return True
    except Exception as e:
        if 'pg_root' in locals(): pg_root.destroy()
        app.ScreenUpdating = True
        return False
def final_check_summary(file_name):
    """最后的清单确认：将配置参数汇总显示在弹窗正文"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # 构造配置清单字符串
    summary = (
        f"📂 目标文件: {file_name}\n"
        f"--------------------------\n"
        f"1. 选用字体: {config.chinese_font} / {config.english_font}\n"
        f"2. 全局字号: {config.font_size}\n"
        f"3. 表格宽度: {config.table_width_percent}%\n"
        f"4. 跳过页码: {config.skip_pages if config.skip_pages else '无'}\n"
        f"5. 表名间距: {config.title_format['space_before']} 行\n"
        "--------------------------\n"
        "确认执行后，将自动备份并在源文档执行排版。"
    )
    
    # 使用 askyesno 弹出确认框
    confirm = messagebox.askyesno("请最后核对排版参数", summary, parent=root)
    root.destroy()
    return confirm
# ---------------- 2. 最终唯一控制流 (重构顺序) ----------------
if __name__ == "__main__":
    # A. 先抓取 WPS/Word 进程和文件名
    word_app = get_word_app()
    if not word_app:
        print("【阻断】未检测到运行中的 WPS/Word。")
    else:
        current_file = word_app.ActiveDocument.Name
        
        # B. 带着文件名去弹窗收集参数
        if show_ui_and_get_params(current_file):
            
            # C. 清单汇总确认
            if final_check_summary(current_file):
                
                # D. 执行备份与排版
                if backup_document(word_app):
                    if process_all_tables(word_app):
                        # 结果反馈
                        empty_info = "\n".join(audit_log.empty_cells[:15])
                        if len(audit_log.empty_cells) > 15:
                            empty_info += f"\n... (余下 {len(audit_log.empty_cells)-15} 处省略)"
                        
                        result_msg = (
                            f"✅ 任务完成：{current_file}\n\n"
                            f"统计：成功 {audit_log.success} / 共 {audit_log.total} 个表格\n"
                            f"标红空值数：{len(audit_log.empty_cells)}\n\n"
                            f"📍 明细预览：\n{empty_info if audit_log.empty_cells else '无空值'}"
                        )
                        messagebox.showinfo("执行完毕", result_msg)