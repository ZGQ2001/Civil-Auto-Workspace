"""
===============================================================================
脚本名称：报告表格全量排版引擎 (table_format.py)
作者: ZGQ
功能概述：
    本脚本用于自动化处理 Word/WPS 检测报告中的表格及表名排版，解决手动调表耗时且易错的问题。

核心工作流：
    1. 环境检测：抓取当前处于激活状态的 Word/WPS 文档。
    2. 参数配置：通过 Tkinter 弹窗获取排版参数（字体、字号、表格宽度、跳过页码等）。
    3. 安全备份：调用外部模块 `file_utils.py` 进行静默落盘备份。
    4. 预先扫描：锁定所有表格的初始物理页码，防止排版过程中的动态位移导致误判。
    5. 全量排版：
        - 表名识别：精准捕捉表格上方以“表”开头的段落，并应用字体与间距。
        - 表格处理：统一中英文字体、字号、宽度 100% / 95% 居中对齐。
        - 空值标记：将未填数据的单元格底纹标红，记录三维坐标（页码-表号-单元格）。
    6. 结果汇总：展示处理总数、标红空值明细，完成排版闭环。

前置依赖：
    - 运行前必须打开目标文档。
    - 同级目录下需存在 `file_utils_backup.py` 模块。
===============================================================================
"""
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
import os
import sys
import time
import re
import win32com.client
import pythoncom

# 【挂载外部模块备份文件】
# 确保程序能准确找到同在 02_Core 目录下的 file_utils_backup.py
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document

# ---------------- 1. 配置与审计对象 ----------------
class GlobalConfig:
    def __init__(self):
        self.chinese_font = "宋体"
        self.english_font = "Times New Roman"
        self.font_size = 12.0
        self.table_width_percent = 100
        self.skip_pages = []
        self.empty_cell_color = 255 
        self.max_table_threshold = 100
        self.title_format = {"align": 1, "space_before": 0.5, "space_after": 0}

class AuditLog:
    def __init__(self):
        self.total = 0
        self.success = 0
        self.skipped = 0
        self.errors = 0
        self.empty_cells = []

config = GlobalConfig()
audit_log = AuditLog()

# ---------------- 2. 交互模块 (UI) ----------------
def show_ui_and_get_params(file_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True) 
    prompt_base = f"当前文件：{file_name}\n\n"

    # 1/5：模板
    tpl_input = simpledialog.askstring("1/5", f"{prompt_base}1-公文(仿宋) | 2-标准(宋体)", initialvalue="1", parent=root)
    if not tpl_input: return False
    if tpl_input == "1":
        config.chinese_font = "仿宋_GB2312"
        config.font_size = 10.5
    else:
        config.chinese_font = "宋体"
        config.font_size = 10.5

    # 2/5：字号
    size_input = simpledialog.askstring("2/5", f"{prompt_base}表格字号：", initialvalue=str(config.font_size), parent=root)
    if not size_input: return False
    config.font_size = float(size_input)

    # 3/5：宽度
    width_input = simpledialog.askstring("3/5", f"{prompt_base}宽度(10-100)：", initialvalue="95", parent=root)
    if not width_input: return False
    config.table_width_percent = int(width_input)

    # 4/5：跳过页码
    skip_input = simpledialog.askstring("4/5", f"{prompt_base}跳过页码(中英逗号均可)：", initialvalue="4", parent=root)
    if skip_input and skip_input.strip():
        normalized = skip_input.replace("，", ",")
        config.skip_pages = [int(p.strip()) for p in normalized.split(",") if p.strip().isdigit()]

    # 5/5：间距
    space_input = simpledialog.askstring("5/5", f"{prompt_base}表名段前间距(行)：", initialvalue="0.5", parent=root)
    if not space_input: return False
    config.title_format["space_before"] = float(space_input)

    root.destroy()
    return True

def final_check_summary(file_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    summary = (
        f"📂 目标文件: {file_name}\n"
        f"--------------------------\n"
        f"中文字体: {config.chinese_font}\n"
        f"英文字体: {config.english_font}\n"
        f"表格字号: {config.font_size}\n"
        f"表格宽度: {config.table_width_percent}%\n"
        f"表名间距: {config.title_format['space_before']} 行\n"
        f"跳过页码: {config.skip_pages if config.skip_pages else '无'}\n"
        "--------------------------\n"
        "确认执行后，将调用静默备份并开始排版。"
    )
    confirm = messagebox.askyesno("请最终确认配置清单", summary, parent=root)
    root.destroy()
    return confirm

# ---------------- 3. 核心引擎 (COM) ----------------
def get_word_app():
    try: return win32com.client.GetActiveObject("Word.Application")
    except:
        try: return win32com.client.GetActiveObject("KWPS.Application")
        except: return None


def process_all_tables(app):
    try:
        doc = app.ActiveDocument
        tables = doc.Tables
        table_count = tables.Count
        audit_log.total = table_count
        if table_count == 0: return True

        # --- 视觉进度条初始化 ---
        pg_root = tk.Tk()
        pg_root.title("表格自动排版程序")
        pg_root.attributes('-topmost', True)
        pg_root.geometry("350x120")
        tk.Label(pg_root, text=f"正在处理：{doc.Name}", fg="blue").pack(pady=5)
        progress_label = tk.Label(pg_root, text="锁定初始位置中...")
        progress_label.pack()
        bar = ttk.Progressbar(pg_root, length=280, mode='determinate', maximum=table_count)
        bar.pack(pady=10)
        pg_root.update()

        # 预扫描
        table_queue = []
        for i in range(1, table_count + 1):
            tbl = tables.Item(i)
            table_queue.append({
                "obj": tbl,
                "orig_page": tbl.Range.Information(3),
                "index": i
            })

        app.ScreenUpdating = False
        
        # 执行循环
        for item in table_queue:
            tbl = item["obj"]
            page_num = item["orig_page"]
            idx = item["index"]
            
            percent = int((idx / table_count) * 100)
            bar['value'] = idx
            progress_label.config(text=f"正在排版: {idx}/{table_count} (初始页码:{page_num})")
            pg_root.update()

            try:
                # A. 表名判定
                try:
                    title_range = tbl.Range.Previous(4, 1)
                    if title_range and re.sub(r'[\s\x07]', '', title_range.Text).startswith("表"):
                        tf = title_range.Font
                        tf.Name = config.english_font
                        tf.NameFarEast = config.chinese_font
                        tf.Size = config.font_size
                        pf = title_range.ParagraphFormat
                        pf.Alignment = config.title_format["align"]
                        pf.LineUnitBefore = config.title_format["space_before"]
                        pf.LineUnitAfter = 0
                except: pass

                # B. 跳过页码
                if page_num in config.skip_pages:
                    audit_log.skipped += 1
                    continue

                # C. 表格整体格式
                tbl.PreferredWidthType = 2
                tbl.PreferredWidth = config.table_width_percent
                tbl.Rows.Alignment = 1

                # D. 单元格一维遍历
                cells = tbl.Range.Cells
                for j in range(1, cells.Count + 1):
                    cell = cells.Item(j)
                    clean_text = re.sub(r'[\r\n\x07\s]', '', cell.Range.Text)
                    if not clean_text:
                        cell.Shading.BackgroundPatternColor = config.empty_cell_color
                        audit_log.empty_cells.append(f"P{page_num}-T{idx}-C{j}")
                    else:
                        f = cell.Range.Font
                        f.Name = config.english_font
                        f.NameFarEast = config.chinese_font
                        f.Size = config.font_size
                        cell.VerticalAlignment = 1
                        cell.Range.ParagraphFormat.Alignment = 1
                
                audit_log.success += 1
            except Exception as e:
                audit_log.errors += 1
                audit_log.error_details.append(f"T{idx} 崩溃: {e}")

        pg_root.destroy()
        doc.Save()
        app.ScreenUpdating = True
        return True
    except Exception as e:
        if 'pg_root' in locals(): pg_root.destroy()
        app.ScreenUpdating = True
        return False

# ---------------- 4. 最终主控制流 ----------------
if __name__ == "__main__":
    word_app = get_word_app()
    if not word_app:
        print("【阻断】未检测到运行中的 WPS/Word。")
    else:
        current_file = word_app.ActiveDocument.Name
        if show_ui_and_get_params(current_file) and final_check_summary(current_file):
            
            # 【标注 3：调用外部备份模块并增加熔断机制】
            print("正在调用外部模块进行静默备份...")
            if backup_current_document(word_app):
                
                # 备份成功，放行排版逻辑
                if process_all_tables(word_app):
                    empty_info = "\n".join(audit_log.empty_cells[:15])
                    if len(audit_log.empty_cells) > 15:
                        empty_info += f"\n... (余下 {len(audit_log.empty_cells)-15} 处省略)"
                    
                    result_msg = (
                        f"✅ 任务完成：{current_file}\n\n"
                        f"成功排版: {audit_log.success} / 共 {audit_log.total} 表\n"
                        f"标红空值: {len(audit_log.empty_cells)} 处\n\n"
                        f"📍 坐标参考（锁定初始位置）：\n{empty_info if audit_log.empty_cells else '无'}"
                    )
                    messagebox.showinfo("执行流转完成", result_msg)
            else:
                # 备份失败，UI 拦截（因为底层模块现在是静默的）
                err_root = tk.Tk()
                err_root.withdraw()
                err_root.attributes('-topmost', True)
                messagebox.showerror(
                    "安全熔断", 
                    "⚠️ 备份模块(file_utils)返回失败信号！\n\n为防止原文件损坏，排版程序已自动终止。\n请检查当前文档是否已保存，或查看后台报错日志。", 
                    parent=err_root
                )
                err_root.destroy()