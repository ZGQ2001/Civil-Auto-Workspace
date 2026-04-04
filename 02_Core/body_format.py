"""
===============================================================================
脚本名称：报告正文排版引擎 (body_format.py)
作者：ZGQ
功能概述：
    基于外部 JSON 配置和正则表达式，对 Word 文档的非结构化正文进行特征识别与精准排版。
    已集成：交互选择、防呆确认、可视化进度白盒、物理页码隔离、图表格式强保护、防御性兜底。
===============================================================================
"""

import os
import sys
import json
import re
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk

# 挂载外部备份模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document


# ==================== 板块 1：配置与规则大脑 ====================

def load_style_config(report_type="检测报告"):
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '04_Config', 'report_style_config.json'))
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"【阻断】未找到配置文件：{config_path}")
    with open(config_path, 'r', encoding='utf-8') as f:
        full_config = json.load(f)
    if report_type not in full_config:
        raise ValueError(f"【阻断】配置文件中不存在该报告类型：{report_type}")
    return full_config[report_type]

class ParagraphClassifier:
    def __init__(self):
        self.re_fig_tbl = re.compile(r'^\s*(图|表)\s*(\d+(\.\d+)*)')
        self.re_note = re.compile(r'^(注|说明)[：:]')
        self.re_list_item = re.compile(r'^(\d+[.、）\)]|[①②③④⑤⑥⑦⑧⑨⑩])')
        self.re_blank = re.compile(r'.*[（(]?(本页)?以下空白[）)].*')
        self.re_no_indent = re.compile(r'^\s*[《\(\[（]')
        self.re_h1 = re.compile(r'^(\d+[\.\s\u3000\t]+|[一二三四五六七八九十]+[、\.\s\u3000\t]+)')
        self.re_h2 = re.compile(r'^\d+[\.．]\d+[\s\u3000\t]*')
        self.re_h3 = re.compile(r'^\d+[\.．]\d+[\.．]\d+[\s\u3000\t]*')
        self.re_appr_c_h1 = re.compile(r'^[检\s·]*测[\s·]*结[\s·]*论[\s·]*与[\s·]*建[\s·]*议$')
        self.re_appr_c_h2 = re.compile(r'^\d+[\.．\s\u3000\t]+[\u4e00-\u9fa5]+')
        self.re_basis_title = re.compile(r'^[\d\.．\s\u3000\t]*(检测|鉴定)依据.*') 
        self.re_suggest_title = re.compile(r'^[处\s]*理[\s]*建[\s]*议$')

    def classify(self, text, list_string="", is_in_note_mode=False, is_in_basis_mode=False, is_in_conclusion_mode=False, report_type="检测报告"):
        raw_text = f"{list_string}{text}"
        clean_text = re.sub(r'\x13.*?\x14', '', raw_text)
        clean_text = re.sub(r'[\x13\x14\x15\x07\x01\x02]', '', clean_text).replace('\xa0', ' ').strip()
        
        if not clean_text: return "空行"
        if self.re_blank.search(clean_text): return "空白提示"
        if self.re_fig_tbl.match(clean_text): return "图表名称"

        if report_type == "鉴定报告":
            condensed = clean_text.replace(" ", "").replace("·", "").replace("\u3000", "")
            if condensed == "检测结论与建议" or self.re_suggest_title.match(condensed):
                return "结论一级标题"
            if is_in_conclusion_mode and self.re_appr_c_h2.match(clean_text):
                return "结论二级标题"
            
            if self.re_h3.match(clean_text): return "三级标题"
            if self.re_h2.match(clean_text): return "二级标题"
            if self.re_h1.match(clean_text): return "一级标题"
            
            if is_in_basis_mode: return "无缩进正文"
            if self.re_no_indent.match(clean_text): return "无缩进正文"
            return "标准正文"
        else:
            if self.re_note.match(clean_text): return "表注说明_起点"
            if is_in_note_mode and self.re_list_item.match(clean_text): return "表注说明_延续"
            if self.re_h3.match(clean_text): return "三级标题"
            if self.re_h2.match(clean_text): return "二级标题"
            if self.re_h1.match(clean_text): return "一级标题"
            
            # 【核心修正1】：将无缩进特征放开给检测报告，识别《XX标准》
            if is_in_basis_mode: return "无缩进正文"
            if self.re_no_indent.match(clean_text): return "无缩进正文"
            
            return "标准正文"


# ==================== 板块 2：交互与参数获取 (UI) ====================

def get_user_params(file_name):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    prompt_base = f"当前排版文件：{file_name}\n\n"

    type_input = simpledialog.askstring(
        "1/2", 
        f"{prompt_base}请选择处理的报告类型：\n1 - 检测报告\n2 - 鉴定报告", 
        initialvalue="", 
        parent=root
    )
    if not type_input: return None
    report_type = "鉴定报告" if type_input == "2" else "检测报告"

    skip_input = simpledialog.askstring(
        "2/2", 
        f"{prompt_base}请输入需要跳过的页码（如封面、资质、目录等）：\n页码间用逗号分隔，例如：1,2,3\n若无跳过页，直接点确定。", 
        initialvalue="1,2,3,4", 
        parent=root
    )
    if skip_input is None: return None
    
    skip_pages = []
    if skip_input.strip():
        normalized = skip_input.replace("，", ",")
        skip_pages = [int(p.strip()) for p in normalized.split(",") if p.strip().isdigit()]
        
    root.destroy()
    return {"report_type": report_type, "skip_pages": skip_pages}

def final_check_summary(file_name, params):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    summary = (
        f"📂 目标文件: {file_name}\n"
        f"--------------------------\n"
        f"报告类型: {params['report_type']}\n"
        f"跳过页码: {params['skip_pages'] if params['skip_pages'] else '无'}\n"
        "--------------------------\n"
        "确认执行后，将调用静默备份并启动全量正文排版。"
    )
    confirm = messagebox.askyesno("请最终确认排版清单", summary, parent=root)
    root.destroy()
    return confirm


# ==================== 板块 3：格式引擎 ====================

def apply_paragraph_format(para, style_config, para_type):
    try:
        if para_type == "图片":
            pf = para.Format
            pf.Alignment = style_config.get("alignment", 1) 
            pf.SpaceBefore = style_config.get("space_before", 0.5) * 12
            pf.SpaceAfter = style_config.get("space_after", 0.5) * 12
            pf.LineSpacingRule = 0  
            pf.CharacterUnitFirstLineIndent = 0
            pf.FirstLineIndent = 0
            pf.LeftIndent = 0
            return

        # 具备高容错度的配置读取，防止 JSON 漏配属性
        f = para.Range.Font
        f.NameFarEast = style_config.get("chinese_font", "宋体") 
        f.Name = style_config.get("english_font", "Times New Roman") 
        f.Size = style_config.get("font_size", 12.0) 
        f.Bold = style_config.get("bold", False) 

        pf = para.Format
        pf.Alignment = style_config.get("alignment", 3) 
        
        # 【核心修正2】：智能推测大纲级别，无视 JSON 是否漏填 outline_level，保障导航目录存活
        outline_lvl = style_config.get("outline_level", 10)
        if "一级标题" in para_type or "结论一级标题" in para_type: outline_lvl = 1
        elif "二级标题" in para_type or "结论二级标题" in para_type: outline_lvl = 2
        elif "三级标题" in para_type: outline_lvl = 3
        pf.OutlineLevel = outline_lvl 

        pf.SpaceBefore = style_config.get("space_before", 0) * 12  
        pf.SpaceAfter = style_config.get("space_after", 0) * 12

        if para_type == "图表名称":
            pf.SpaceAfter = 0
            pf.LineSpacingRule = 0 
            pf.CharacterUnitFirstLineIndent = 0
            pf.FirstLineIndent = 0
            pf.CharacterUnitLeftIndent = 0
            pf.LeftIndent = 0
            pf.DisableLineHeightGrid = False
        else:
            ls_rule = style_config.get("line_spacing_rule", 5)
            if ls_rule == 1: pf.LineSpacingRule = 1
            elif ls_rule == 0: pf.LineSpacingRule = 0
            else:
                pf.LineSpacingRule = 5
                pf.LineSpacing = style_config.get("line_spacing", 1.5) * 12
            
            pf.DisableLineHeightGrid = False
            
            if "first_line_indent" in style_config:
                pf.CharacterUnitFirstLineIndent = style_config["first_line_indent"]
                if style_config["first_line_indent"] == 0:
                    pf.FirstLineIndent = 0
            else:
                pf.CharacterUnitFirstLineIndent = 0
                pf.FirstLineIndent = 0

            if "right_indent" in style_config:
                pf.CharacterUnitRightIndent = style_config["right_indent"]
                
            if para_type in ["一级标题", "二级标题", "三级标题", "结论一级标题", "结论二级标题"]:
                pf.CharacterUnitLeftIndent = 0
                pf.LeftIndent = 0 
                pf.CharacterUnitFirstLineIndent = 0
                pf.FirstLineIndent = 0 
            elif para_type == "无缩进正文":
                pf.CharacterUnitLeftIndent = 0
                pf.CharacterUnitFirstLineIndent = 0
                pf.LeftIndent = 30.05
                pf.FirstLineIndent = -20.98
                
    except Exception as e:
        pass


def process_document_body(app, params):
    report_type = params["report_type"]
    skip_pages = params["skip_pages"]
    
    full_config = load_style_config(report_type)
    classifier = ParagraphClassifier()

    doc = app.ActiveDocument
    paragraphs = doc.Paragraphs
    total_paras = paragraphs.Count 
    
    pg_root = tk.Tk()
    pg_root.title("正文自动排版程序")
    pg_root.attributes('-topmost', True)
    pg_root.geometry("350x120")
    tk.Label(pg_root, text=f"正在处理：{doc.Name}", fg="blue").pack(pady=5)
    progress_label = tk.Label(pg_root, text="准备排版...")
    progress_label.pack()
    bar = ttk.Progressbar(pg_root, length=280, mode='determinate', maximum=total_paras)
    bar.pack(pady=10)
    pg_root.update()

    app.ScreenUpdating = False 
    success_count = 0 
    skipped_count = 0
    note_mode = basis_mode = conclusion_mode = False 
    
    try:
        for i in range(1, total_paras + 1):
            if i % 10 == 0 or i == total_paras:
                bar['value'] = i
                progress_label.config(text=f"正在排版: {i}/{total_paras} 段")
                pg_root.update()
                
            para = paragraphs.Item(i)
            
            try:
                page_num = para.Range.Information(3) 
                if page_num in skip_pages:
                    skipped_count += 1
                    continue
            except:
                pass 

            if para.Range.Information(12) or "目录" in para.Style.NameLocal or "TOC" in para.Style.NameLocal:
                skipped_count += 1
                continue
                
            text = para.Range.Text
            list_str = para.Range.ListFormat.ListString 

            clean_t = text.strip()
            is_basis_header = classifier.re_basis_title.match(clean_t) 
            is_suggest_header = classifier.re_suggest_title.match(clean_t.replace(" ", ""))

            has_image = False
            try:
                if para.Range.InlineShapes.Count > 0:
                    has_image = True
            except:
                pass

            if has_image:
                para_type = "图片"
            else:
                para_type = classifier.classify(text, list_str, note_mode, basis_mode, conclusion_mode, report_type)

            if para_type == "结论一级标题":
                conclusion_mode = True 
            elif para_type == "一级标题":
                conclusion_mode = False 
                
            if is_basis_header:
                basis_mode = True
            elif para_type in ["一级标题", "二级标题", "三级标题", "结论一级标题", "结论二级标题"]:
                basis_mode = False

            if para_type == "表注说明_起点":
                note_mode = True
            elif para_type not in ["表注说明_延续", "空行"]:
                note_mode = False
                
            if para_type == "标准正文" and re.match(r'^(\d+[.、）\)]|[①②③④⑤⑥⑦⑧⑨⑩])', text.strip()):
                continue
                
            if para_type == "标准正文" and re.match(r'^[\s]*[•\-*]\s+', text.strip()):
                para_type = "无缩进正文"
                
            # 【核心修正3】：当 JSON 未配置特定板块时，启动兜底容错排版
            if para_type in full_config:
                apply_paragraph_format(para, full_config[para_type], para_type)
                success_count += 1
            elif para_type in ["图片", "图表名称", "无缩进正文", "空白提示"]:
                fallback_cfg = {
                    "chinese_font": "宋体", "english_font": "Times New Roman", 
                    "font_size": 10.5 if para_type == "图表名称" else 12.0, 
                    "alignment": 1 if para_type in ["图片", "图表名称"] else 0,
                    "first_line_indent": 0
                }
                apply_paragraph_format(para, fallback_cfg, para_type)
                success_count += 1

    finally:
        if 'pg_root' in locals():
            pg_root.destroy()
        app.ScreenUpdating = True 
        
    return success_count, skipped_count


# ==================== 最终主控流 ====================

if __name__ == "__main__":
    try:
        app = win32com.client.GetActiveObject("Word.Application") 
    except:
        try:
            app = win32com.client.GetActiveObject("KWPS.Application") 
        except:
            app = None 

    if not app:
        print("【阻断】未检测到运行中的 WPS/Word。")
    else:
        current_file = app.ActiveDocument.Name
        run_params = get_user_params(current_file)
        
        if run_params is None:
            print("【取消】用户取消了操作，程序终止。")
        else:
            if final_check_summary(current_file, run_params):
                print("正在调用外部模块进行静默备份...")
                if backup_current_document(app):
                    succ_cnt, skip_cnt = process_document_body(app, run_params)
                    
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes('-topmost', True)
                    messagebox.showinfo(
                        "执行完毕", 
                        f"✅ 正文排版任务完成！\n\n"
                        f"处理报告类型：{run_params['report_type']}\n"
                        f"成功刷入格式：{succ_cnt} 段\n"
                        f"因规则或页码跳过：{skip_cnt} 段"
                    )
                    root.destroy()
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