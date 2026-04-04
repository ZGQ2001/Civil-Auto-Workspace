"""
===============================================================================
脚本名称：报告正文排版引擎 (body_format.py)
作者：ZGQ
功能概述：
    基于外部 JSON 配置和正则表达式，对 Word 文档的非结构化正文进行特征识别与精准排版。
===============================================================================
"""

import os
import sys
import json
import re
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import messagebox

# 挂载外部备份模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document


# ==================== 板块 1：配置与规则大脑 ====================

def load_style_config(report_type="检测报告"):
    """
    读取 04_Config 下的样式字典
    """
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '04_Config', 'report_style_config.json'))
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"【阻断】未找到配置文件：{config_path}")
        
    with open(config_path, 'r', encoding='utf-8') as f:
        full_config = json.load(f)
        
    if report_type not in full_config:
        raise ValueError(f"【阻断】配置文件中不存在该报告类型：{report_type}")
        
    return full_config[report_type]

class ParagraphClassifier:
    """
    段落特征识别器（正则雷达）
    依据报告的通用编号和文本规律，对段落类型进行精准定性。
    """
    def __init__(self):
        # 1. 图名/表名 (优化：兼容交叉引用)
        self.re_fig_tbl = re.compile(r'^(图|表)\s*\d+[-_.\s]*\d+')
        # 2. 表注/说明
        self.re_note = re.compile(r'^(注|说明)[：:]')
        # 3. 自动编号与多行表注的连续性前缀
        self.re_list_item = re.compile(r'^(\d+[.、）\)]|[①②③④⑤⑥⑦⑧⑨⑩])')
        # 4. 一级标题 (匹配 "1 概况", "1. 概况", "一、概况")
        self.re_h1 = re.compile(r'^(\d+[\.\s\u3000\t]+|[一二三四五六七八九十]+[、\.\s\u3000\t]+)')
        # 5. 二级标题 (兼容 1.1, 1．1，允许后面紧跟汉字无空格)
        self.re_h2 = re.compile(r'^\d+[\.．]\d+[\s\u3000\t]*')
        # 6. 三级标题 (兼容 1.1.1，允许后面紧跟汉字无空格)
        self.re_h3 = re.compile(r'^\d+[\.．]\d+[\.．]\d+[\s\u3000\t]*')

    def classify(self, text, list_string="", is_in_note_mode=False):
        """核心分类逻辑"""
        raw_text = f"{list_string}{text}"
        
        # 清洗交叉引用与不可见控制符
        clean_text = re.sub(r'\x13.*?\x14', '', raw_text)
        clean_text = re.sub(r'[\x15\x07]', '', clean_text)
        clean_text = clean_text.replace('\xa0', ' ').strip()
        
        if not clean_text:
            return "空行"

        if self.re_fig_tbl.match(clean_text): return "图表名称"
        if self.re_note.match(clean_text): return "表注说明_起点"
        
        if is_in_note_mode:
            if self.re_list_item.match(clean_text): return "表注说明_延续"
            else: pass

        if self.re_h3.match(clean_text): return "三级标题"
        if self.re_h2.match(clean_text): return "二级标题"
        if self.re_h1.match(clean_text): return "一级标题"

        return "标准正文"


def apply_paragraph_format(para, style_config, para_type):
    """
    根据配置字典，将样式物理硬刷到 Word 段落上
    """
    try:
        # 1. 字体同步
        f = para.Range.Font
        f.NameFarEast = style_config["chinese_font"]
        f.Name = style_config["english_font"]
        f.Size = style_config["font_size"]
        f.Bold = style_config.get("bold", False)

        # 2. 对齐与大纲级别 (大纲级别能让左侧导航栏自动生成目录)
        pf = para.Format
        pf.Alignment = style_config["alignment"]
        if para_type == "一级标题":
            pf.OutlineLevel = 1
        elif para_type == "二级标题":
            pf.OutlineLevel = 2
        elif para_type == "三级标题":
            pf.OutlineLevel = 3
        else:
            pf.OutlineLevel = 10 # 10 代表 wdOutlineLevelBodyText (正文文本)

        # 3. 间距处理
        pf.SpaceBefore = style_config.get("space_before", 0) * 12  
        pf.SpaceAfter = style_config.get("space_after", 0) * 12
        
        # 读取 JSON 中精确的行距规则
        ls_rule = style_config.get("line_spacing_rule", 5)
        if ls_rule == 1:
            pf.LineSpacingRule = 1 # 标准 1.5 倍行距
        elif ls_rule == 0:
            pf.LineSpacingRule = 0 # 单倍行距
        else:
            pf.LineSpacingRule = 5
            pf.LineSpacing = style_config.get("line_spacing", 1.5) * 12
        
        # 4. 强制取消“与文档网格对齐”，防止行距被异常拉长
        pf.DisableLineHeightGrid = False 
        
        # 5. 缩进处理 (非破坏性修改与幽灵缩进压制)
        if "first_line_indent" in style_config:
            pf.CharacterUnitFirstLineIndent = style_config["first_line_indent"]
        if "right_indent" in style_config:
            pf.CharacterUnitRightIndent = style_config["right_indent"]
            
        # 【核心修改点：彻底抹杀自动编号带来的“文本之前 1cm”幽灵左缩进】
        if para_type in ["一级标题", "二级标题", "三级标题"]:
            pf.CharacterUnitLeftIndent = 0
            pf.LeftIndent = 0  # 强制将磅值归零
            
    except Exception as e:
        print(f"应用格式异常: {e}")

def process_document_body(app, report_type="检测报告"):
    """
    接管 Word 文档，执行遍历排版
    """
    full_config = load_style_config(report_type)
    classifier = ParagraphClassifier()
    
    doc = app.ActiveDocument
    paragraphs = doc.Paragraphs
    total_paras = paragraphs.Count
    
 # === 寻找启动锚点 ===
    start_index = 0
    for i in range(1, min(300, total_paras + 1)):
        p = paragraphs.Item(i)
        
        # 【强力免疫1】通过Word自带的样式名，直接无视所有的目录页！
        if "目录" in p.Style.NameLocal or "TOC" in p.Style.NameLocal:
            continue
            
        list_str = p.Range.ListFormat.ListString
        text = p.Range.Text
        
        if classifier.classify(text, list_str) == "一级标题":
            # 【强力免疫2】双保险防止抓到文本假目录
            if '\t' in text or re.search(r'[.·…]{3,}', text):
                continue 
            start_index = i
            break

    if start_index == 0:
        print("【阻断】未能找到启动锚点(一级标题)，为防止误伤封面，程序终止。")
        return False
        
    print(f"【定位】成功锁定启动锚点在第 {start_index} 段，开始向下排版...")

    # === 开始遍历排版 ===
    app.ScreenUpdating = False
    success_count = 0
    note_mode = False
    
    try:
        for i in range(start_index, total_paras + 1):
            para = paragraphs.Item(i)
            
            # 【全局免疫】表格内的不管，目录页里的也不管！
            if para.Range.Information(12) or "目录" in para.Style.NameLocal or "TOC" in para.Style.NameLocal:
                continue
                
            list_str = para.Range.ListFormat.ListString
            text = para.Range.Text
            
            para_type = classifier.classify(text, list_str, note_mode)
            
            if para_type == "表注说明_起点": note_mode = True
            elif para_type not in ["表注说明_延续", "空行"]: note_mode = False
                
            if para_type in full_config:
                apply_paragraph_format(para, full_config[para_type], para_type)
                success_count += 1

    except Exception as e:
        print(f"【异常】排版中断: {e}")
    finally:
        app.ScreenUpdating = True
        
    print(f"【完成】排版结束。共刷入格式 {success_count} 个段落。")
    return True

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
        # 先调用静态备份模块
        if backup_current_document(app):
            # 备份成功后，执行排版
            process_document_body(app, "检测报告")
        else:
            print("【安全阻断】备份失败，排版取消。")