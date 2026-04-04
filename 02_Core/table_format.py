"""
===============================================================================
脚本名称：报告表格全量排版引擎 (table_format.py)
作者: ZGQ
功能概述：
    本脚本用于自动化处理 Word/WPS 检测报告中的表格及表名排版。
    V2.0 重构版：全面接入 report_style_config.json，解除字体字号硬编码。
===============================================================================
"""
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
import os
import sys
import json
import re
import win32com.client
import pythoncom

# 挂载外部模块备份文件
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document

# ---------------- 1. 配置与规则读取 ----------------

def load_style_config(report_type="检测报告"):
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '04_Config', 'report_style_config.json'))
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"【阻断】未找到配置文件：{config_path}")
    with open(config_path, 'r', encoding='utf-8') as f:
        full_config = json.load(f)
    if report_type not in full_config:
        raise ValueError(f"【阻断】配置文件中不存在该报告类型：{report_type}")
    return full_config[report_type]

class GlobalConfig:
    def __init__(self):
        self.report_type = "检测报告"
        self.table_width_percent = 100
        self.skip_pages = []
        self.empty_cell_color = 255 

class AuditLog:
    def __init__(self):
        self.total = 0
        self.success = 0
        self.skipped = 0
        self.errors = 0
        self.empty_cells = []
        self.error_details = []

config = GlobalConfig()
audit_log = AuditLog()

# ---------------- 2. 交互模块 (UI) ----------------
def show_ui_and_get_params(file_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True) 
    prompt_base = f"当前文件：{file_name}\n\n"

    # 1/3：报告类型
    tpl_input = simpledialog.askstring("1/3", f"{prompt_base}请选择处理的报告类型：\n1 - 检测报告\n2 - 鉴定报告", initialvalue="", parent=root)
    if not tpl_input: return False
    config.report_type = "鉴定报告" if tpl_input == "2" else "检测报告"

    # 2/3：宽度
    width_input = simpledialog.askstring("2/3", f"{prompt_base}表格全局宽度(10-100)：", initialvalue="95", parent=root)
    if not width_input: return False
    config.table_width_percent = int(width_input)

    # 3/3：跳过页码
    skip_input = simpledialog.askstring("3/3", f"{prompt_base}跳过页码(如封面、资质等)：\n页码间用逗号分隔，例如：1,2,3", initialvalue="", parent=root)
    if skip_input and skip_input.strip():
        normalized = skip_input.replace("，", ",")
        config.skip_pages = [int(p.strip()) for p in normalized.split(",") if p.strip().isdigit()]

    root.destroy()
    return True

def final_check_summary(file_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    summary = (
        f"📂 目标文件: {file_name}\n"
        f"--------------------------\n"
        f"报告类型: {config.report_type}\n"
        f"表格宽度: {config.table_width_percent}%\n"
        f"跳过页码: {config.skip_pages if config.skip_pages else '无'}\n"
        "--------------------------\n"
        "字体、字号及表名间距将自动从 JSON 配置库读取。\n\n"
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
        # 读取 JSON 数据源
        style_db = load_style_config(config.report_type)
        # 获取表名卡片（如果漏配，则兜底）
        title_cfg = style_db.get("图表名称", {"chinese_font": "宋体", "english_font": "Times New Roman", "font_size": 10.5, "alignment": 1, "space_before": 0.5, "space_after": 0})
        # 获取表内容卡片
        cell_cfg = style_db.get("表格正文", {"chinese_font": "宋体", "english_font": "Times New Roman", "font_size": 10.5})

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
            try:
                orig_page = tbl.Range.Information(3)
            except:
                orig_page = 0
            table_queue.append({
                "obj": tbl,
                "orig_page": orig_page,
                "index": i
            })

        app.ScreenUpdating = False
        
        # 执行循环
        for item in table_queue:
            tbl = item["obj"]
            page_num = item["orig_page"]
            idx = item["index"]
            
            bar['value'] = idx
            progress_label.config(text=f"正在排版: {idx}/{table_count} (初始页码:{page_num})")
            pg_root.update()

            try:
                # B. 跳过页码
                if page_num in config.skip_pages:
                    audit_log.skipped += 1
                    continue

                # A. 表名判定与 JSON 规则下发
                try:
                    title_range = tbl.Range.Previous(4, 1)
                    if title_range and re.sub(r'[\s\x07]', '', title_range.Text).startswith("表"):
                        tf = title_range.Font
                        tf.Name = title_cfg["english_font"]
                        tf.NameFarEast = title_cfg["chinese_font"]
                        tf.Size = title_cfg["font_size"]
                        tf.Bold = title_cfg.get("bold", False)
                        
                        pf = title_range.ParagraphFormat
                        pf.Alignment = title_cfg.get("alignment", 1)
                        pf.LineUnitBefore = title_cfg.get("space_before", 0.5)
                        pf.LineUnitAfter = title_cfg.get("space_after", 0.0)
                        pf.CharacterUnitFirstLineIndent = 0
                        pf.FirstLineIndent = 0
                        pf.CharacterUnitLeftIndent = 0
                        pf.LeftIndent = 0
                except: pass

                # C. 表格整体格式
                tbl.PreferredWidthType = 2
                tbl.PreferredWidth = config.table_width_percent
                tbl.Rows.Alignment = 1

                # D. 单元格一维遍历与 JSON 规则下发
                cells = tbl.Range.Cells
                for j in range(1, cells.Count + 1):
                    cell = cells.Item(j)
                    clean_text = re.sub(r'[\r\n\x07\s]', '', cell.Range.Text)
                    if not clean_text:
                        cell.Shading.BackgroundPatternColor = config.empty_cell_color
                        audit_log.empty_cells.append(f"P{page_num}-T{idx}-C{j}")
                    else:
                        f = cell.Range.Font
                        f.Name = cell_cfg["english_font"]
                        f.NameFarEast = cell_cfg["chinese_font"]
                        f.Size = cell_cfg["font_size"]
                        f.Bold = cell_cfg.get("bold", False)
                        
                        cell.VerticalAlignment = 1
                        cell.Range.ParagraphFormat.Alignment = cell_cfg.get("alignment", 1)
                
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
            print("正在调用外部模块进行静默备份...")
            if backup_current_document(word_app):
                if process_all_tables(word_app):
                    empty_info = "\n".join(audit_log.empty_cells[:15])
                    if len(audit_log.empty_cells) > 15:
                        empty_info += f"\n... (余下 {len(audit_log.empty_cells)-15} 处省略)"
                    
                    result_msg = (
                        f"✅ 任务完成：{current_file}\n\n"
                        f"成功排版: {audit_log.success} / 共 {audit_log.total} 表\n"
                        f"因页码跳过: {audit_log.skipped} 表\n"
                        f"标红空值: {len(audit_log.empty_cells)} 处\n\n"
                        f"📍 坐标参考（锁定初始位置）：\n{empty_info if audit_log.empty_cells else '无'}"
                    )
                    messagebox.showinfo("执行流转完成", result_msg)
            else:
                err_root = tk.Tk()
                err_root.withdraw()
                err_root.attributes('-topmost', True)
                messagebox.showerror(
                    "安全熔断", 
                    "⚠️ 备份模块(file_utils)返回失败信号！\n\n为防止原文件损坏，排版程序已自动终止。\n请检查当前文档是否已保存，或查看后台报错日志。", 
                    parent=err_root
                )
                err_root.destroy()