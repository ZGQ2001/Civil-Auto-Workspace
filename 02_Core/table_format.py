"""
===============================================================================
脚本名称：报告表格全量排版引擎 (table_format.py)
作者: ZGQ
功能概述：
    本脚本用于自动化处理 Word/WPS 检测报告中的表格及表名排版。
    V2.0 重构版：全面接入 report_style_config.json，解除字体字号硬编码。

    这个脚本可以自动排版Word文档中的所有表格，包括设置字体、字号、对齐方式、表格宽度等。
    它会根据配置文件自动应用统一的格式，确保表格看起来专业一致。新手可以通过简单的选择来完成复杂的表格排版。
===============================================================================
"""
import tkinter as tk  # GUI库
from tkinter import simpledialog, messagebox, ttk  # tkinter子模块
import os  # 文件路径操作
import sys  # 系统操作
import json  # JSON文件处理
import re  # 正则表达式（虽然在这个文件中可能没用，但保留）
import win32com.client  # 控制Word/WPS
import pythoncom  # COM组件初始化
from word_env_utils import word_optimized_environment  # Word环境优化上下文管理器

# 挂载外部模块备份文件
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document  # 导入备份函数

# ---------------- 1. 配置与规则读取 ----------------

def load_style_config(report_type="检测报告"):
    """
    加载样式配置文件。

    参数:
        report_type (str): 报告类型。

    返回:
        dict: 样式配置字典。
    """
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '04_Config', 'report_style_config.json'))
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"【阻断】未找到配置文件：{config_path}")
    with open(config_path, 'r', encoding='utf-8') as f:
        full_config = json.load(f)
    if report_type not in full_config:
        raise ValueError(f"【阻断】配置文件中不存在该报告类型：{report_type}")
    return full_config[report_type]

class GlobalConfig:
    """
    全局配置类。

    存储用户选择的配置参数。
    """
    def __init__(self):
        self.report_type = "检测报告"  # 报告类型
        self.table_width_percent = 100  # 表格宽度百分比
        self.skip_pages = []  # 跳过页码列表
        self.empty_cell_color = 255  # 空单元格颜色

class AuditLog:
    """
    审计日志类。

    记录处理过程中的统计信息和错误详情。
    """
    def __init__(self):
        self.total = 0  # 总表格数
        self.success = 0  # 成功处理的表格数
        self.skipped = 0  # 跳过的表格数
        self.errors = 0  # 错误数
        self.empty_cells = []  # 空单元格列表
        self.error_details = []  # 错误详情列表

# 全局实例
config = GlobalConfig()
audit_log = AuditLog()

config = GlobalConfig()
audit_log = AuditLog()

# ---------------- 2. 交互模块 (UI) ----------------
def show_ui_and_get_params(file_name):
    """
    显示用户界面并获取参数。

    通过弹窗让用户选择报告类型、表格宽度和跳过页码。

    参数:
        file_name (str): 当前文件名。

    返回:
        bool: 是否成功获取参数。
    """
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
        normalized = skip_input.replace("，", ",")  # 支持中文逗号
        config.skip_pages = [int(p.strip()) for p in normalized.split(",") if p.strip().isdigit()]

    root.destroy()
    return True

def final_check_summary(file_name):
    """
    显示最终确认摘要。

    参数:
        file_name (str): 文件名。

    返回:
        bool: 用户是否确认。
    """
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
    """
    获取Word应用程序对象。

    返回:
        Word应用程序对象或None。
    """
    try: 
        return win32com.client.GetActiveObject("Word.Application")
    except:
        try: 
            return win32com.client.GetActiveObject("KWPS.Application")
        except: 
            return None

def process_all_tables(app):
    """
    处理文档中的所有表格。

    参数:
        app: Word应用程序对象。

    返回:
        bool: 处理是否成功。
    """
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
        pg_root.geometry("350x160")  # 加高窗口
        tk.Label(pg_root, text=f"正在处理：{doc.Name}", fg="blue").pack(pady=5)
        progress_label = tk.Label(pg_root, text="准备排版...")
        progress_label.pack()
        bar = ttk.Progressbar(pg_root, length=280, mode='determinate', maximum=table_count)
        bar.pack(pady=10)

        # 【新增】紧急停止机制
        cancel_flag = {"is_cancelled": False}
        def stop_process():
            cancel_flag["is_cancelled"] = True
            progress_label.config(text="正在安全中止，请稍候...", fg="red")
            pg_root.update()

        tk.Button(pg_root, text="紧急停止", command=stop_process, fg="red", width=10).pack()
        pg_root.protocol("WM_DELETE_WINDOW", stop_process)
        pg_root.update()

        # 调用上下文管理器接管环境
        # 调用上下文管理器接管环境
        with word_optimized_environment(app):
            
            # ==========================================
            # 【提速核心：计算最大跳过页，建立“越界断路器”】
            # ==========================================
            max_skip = max(config.skip_pages) if config.skip_pages else 0
            passed_skip_zone = False  # 是否已越过跳过区的标志

            # 【修复】使用 COM 安全的索引遍历，废除极慢的预扫描队列
            for idx in range(1, table_count + 1):
                tbl = tables.Item(idx)
                
                # 监听停止信号
                if cancel_flag["is_cancelled"]:
                    print("【中断】用户手动终止了表格排版。")
                    break

                # 更新UI
                if pg_root.winfo_exists():
                    bar['value'] = idx
                    progress_label.config(text=f"正在排版: {idx}/{table_count}")
                    pg_root.update()

                try:
                    # ==========================================
                    # 【提速 3 过界免检机制】
                    # ==========================================
                    page_num = 999  # 默认赋予安全区页码
                    
                    # 只有当用户填了跳过页，且【还未越过最大跳过页】时，才去查真实页码
                    if config.skip_pages and not passed_skip_zone:
                        try:
                            page_num = tbl.Range.Information(3)
                            # 核心断路器：只要查出来的页码大于最大跳过页，永久关闭查询开关！
                            if page_num > max_skip:
                                passed_skip_zone = True  
                        except:
                            pass
                            
                    # B. 跳过页码判定
                    if page_num != 999 and page_num in config.skip_pages:
                        audit_log.skipped += 1
                        continue

                    # A. 表名判定与 JSON 规则下发
                    try:
                        title_range = tbl.Range.Previous(4, 1)
                        if title_range and re.sub(r'[\s\x07]', '', title_range.Text).startswith("表"):
                            tf = title_range.Font
                            eng_font = title_cfg["english_font"]
                            tf.Name = eng_font
                            tf.NameAscii = eng_font
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
                    # 【修复】使用 COM 安全的索引遍历 Cells
                    cells = tbl.Range.Cells
                    for j in range(1, cells.Count + 1):
                        cell = cells.Item(j)
                        clean_text = re.sub(r'[\r\n\x07\s]', '', cell.Range.Text)
                        
                        if not clean_text:
                            cell.Shading.BackgroundPatternColor = config.empty_cell_color
                            audit_log.empty_cells.append(f"P{page_num}-T{idx}-C{j}")
                        else:
                            f = cell.Range.Font
                            eng_font = cell_cfg["english_font"]
                            f.Name = eng_font
                            f.NameAscii = eng_font
                            f.NameFarEast = cell_cfg["chinese_font"]
                            
                            f.Size = cell_cfg["font_size"]
                            f.Bold = cell_cfg.get("bold", False)
                            
                            cell.VerticalAlignment = 1
                            cell.Range.ParagraphFormat.Alignment = cell_cfg.get("alignment", 1)
                    
                    audit_log.success += 1
                except Exception as e:
                    audit_log.errors += 1
                    audit_log.error_details.append(f"T{idx} 崩溃: {e}")

            doc.Save()
            return True
        
    except Exception as e:
        print(f"执行异常: {e}")
        return False
        
    finally:
        # UI 的销毁保留，Word 状态恢复已交由上下文管理器自动完成
        if 'pg_root' in locals() and pg_root.winfo_exists(): 
            pg_root.destroy()

# ---------------- 4. 最终主控制流 ----------------
if __name__ == "__main__":
    word_app = get_word_app()
    if not word_app:
        err_root = tk.Tk()
        err_root.withdraw()
        err_root.attributes('-topmost', True)
        messagebox.showerror("运行阻断", "未检测到运行中的 WPS 或 Word 程序。\n\n请先打开需要排版的报告文档！", parent=err_root)
        err_root.destroy()
    else:
        # 【新增：隐患拦截】
        if word_app.ActiveDocument.Path == "":
            err_root = tk.Tk()
            err_root.withdraw()
            err_root.attributes('-topmost', True)
            messagebox.showwarning("操作阻断", "该文档尚未保存到本地硬盘。\n请先手动保存一次（Ctrl+S）后再执行排版引擎！", parent=err_root)
            err_root.destroy()
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