"""
===============================================================================
脚本名称：报告正文排版引擎 (body_format.py)
作者：ZGQ
功能概述：
    基于外部 JSON 配置和正则表达式，对 Word 文档的非结构化正文进行特征识别与精准排版。
    V2.0 重构版：引擎纯粹化，解除代码硬编码，全量依赖 JSON 数据驱动。

    这个脚本是用来自动排版Word文档中的正文部分的。它可以识别标题、段落、图表等，
    并根据配置文件自动设置字体、间距、对齐方式等格式。适合新手使用，只需运行脚本并选择选项即可。
===============================================================================
"""

import os  # 用于处理文件路径
import sys  # 用于系统操作，如添加模块路径
import json  # 用于读取JSON配置文件
import re  # 正则表达式模块，用于文本匹配
import win32com.client  # 用于控制Word或WPS应用程序
import pythoncom  # COM组件初始化
import tkinter as tk  # GUI库，用于创建用户界面
from tkinter import simpledialog, messagebox, ttk  # tkinter的子模块，用于对话框和进度条

# 挂载外部备份模块
# 这行代码将当前文件的上级目录添加到Python路径中，这样就可以导入同级目录下的模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document  # 导入备份函数，用于在修改前备份文档


# ==================== 板块 1：配置与规则大脑 ====================

def load_style_config(report_type="检测报告"):
    """
    加载样式配置文件。

    这个函数从配置文件中读取排版样式设置。根据报告类型返回相应的配置。

    参数:
        report_type (str): 报告类型，如"检测报告"或"鉴定报告"。默认为"检测报告"。

    返回:
        dict: 包含样式配置的字典。

    异常:
        FileNotFoundError: 如果配置文件不存在。
        ValueError: 如果配置文件中没有指定的报告类型。
    """
    # 构建配置文件的绝对路径
    config_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '04_Config', 'report_style_config.json'))
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"【阻断】未找到配置文件：{config_path}")
    # 打开并读取JSON文件
    with open(config_path, 'r', encoding='utf-8') as f:
        full_config = json.load(f)
    if report_type not in full_config:
        raise ValueError(f"【阻断】配置文件中不存在该报告类型：{report_type}")
    return full_config[report_type]


class ParagraphClassifier:
    """
    段落分类器类。

    这个类用于分析Word文档中的段落文本，并将其分类为不同的类型，如标题、正文、图表等。
    它使用正则表达式来匹配文本模式。
    """

    def __init__(self):
        """
        初始化分类器。

        在这里定义各种正则表达式，用于匹配不同类型的段落。
        """
        # 匹配图或表的标题，如"图 1"或"表 2.1"
        self.re_fig_tbl = re.compile(r'^\s*(图|表)\s*(\d+(\.\d+)*)')
        # 匹配注或说明的开始
        self.re_note = re.compile(r'^(注|说明)[：:]')
        # 匹配列表项，如数字或中文序号
        self.re_list_item = re.compile(r'^(\d+[.、）\)]|[①②③④⑤⑥⑦⑧⑨⑩])')
        # 匹配空白提示
        self.re_blank = re.compile(r'.*[（(]?(本页)?以下空白[）)].*')
        # 匹配不需要缩进的段落，如以括号或书名号开始
        self.re_no_indent = re.compile(r'^\s*[《\(\[（]')
        # 匹配一级标题，如"1."或"一、"
        self.re_h1 = re.compile(r'^(\d+[\.\s\u3000\t]+|[一二三四五六七八九十]+[、\.\s\u3000\t]+)')
        # 匹配二级标题，如"1.1"
        self.re_h2 = re.compile(r'^\d+[\.．]\d+[\s\u3000\t]*')
        # 匹配三级标题，如"1.1.1"
        self.re_h3 = re.compile(r'^\d+[\.．]\d+[\.．]\d+[\s\u3000\t]*')
        # 匹配结论一级标题
        self.re_appr_c_h1 = re.compile(r'^[检\s·]*测[\s·]*结[\s·]*论[\s·]*与[\s·]*建[\s·]*议$')
        # 匹配结论二级标题
        self.re_appr_c_h2 = re.compile(r'^\d+[\.．\s\u3000\t]+[\u4e00-\u9fa5]+')
        # 匹配检测依据标题
        self.re_basis_title = re.compile(r'^[\d\.．\s\u3000\t]*(检测|鉴定)依据.*')
        # 匹配处理建议标题
        self.re_suggest_title = re.compile(r'^[处\s]*理[\s]*建[\s]*议$')

    def classify(self, text, list_string="", is_in_note_mode=False, is_in_basis_mode=False, is_in_conclusion_mode=False, report_type="检测报告"):
        """
        分类段落类型。

        根据段落的文本内容和上下文状态，将其分类为不同的类型。

        参数:
            text (str): 段落的文本内容。
            list_string (str): 列表字符串，如果段落是列表项的话。默认为空字符串。
            is_in_note_mode (bool): 是否处于表注说明模式。默认为False。
            is_in_basis_mode (bool): 是否处于检测依据模式。默认为False。
            is_in_conclusion_mode (bool): 是否处于结论模式。默认为False。
            report_type (str): 报告类型。默认为"检测报告"。

        返回:
            str: 段落的类型，如"一级标题"、"标准正文"等。
        """
        # 合并列表字符串和文本，用于完整分析
        raw_text = f"{list_string}{text}"
        # 清理文本，移除Word的特殊字符
        clean_text = re.sub(r'\x13.*?\x14', '', raw_text)
        clean_text = re.sub(r'[\x13\x14\x15\x07\x01\x02]', '', clean_text).replace('\xa0', ' ').strip()
        
        # 如果清理后文本为空，返回空行类型
        if not clean_text: return "空行"
        # 检查是否是空白提示
        if self.re_blank.search(clean_text): return "空白提示"
        # 检查是否是图表名称
        if self.re_fig_tbl.match(clean_text): return "图表名称"

        # 根据报告类型进行不同的分类逻辑
        if report_type == "鉴定报告":
            # 移除空格和特殊字符，用于匹配
            condensed = clean_text.replace(" ", "").replace("·", "").replace("\u3000", "")
            # 检查是否是结论一级标题
            if condensed == "检测结论与建议" or self.re_suggest_title.match(condensed):
                return "结论一级标题"
            # 如果在结论模式下，检查是否是结论二级标题
            if is_in_conclusion_mode and self.re_appr_c_h2.match(clean_text):
                return "结论二级标题"
            
            # 检查标题层级
            if self.re_h3.match(clean_text): return "三级标题"
            if self.re_h2.match(clean_text): return "二级标题"
            if self.re_h1.match(clean_text): return "一级标题"
            
            # 如果在检测依据模式下，返回无缩进正文
            if is_in_basis_mode: return "无缩进正文"
            # 检查是否不需要缩进
            if self.re_no_indent.match(clean_text): return "无缩进正文"
            # 默认返回标准正文
            return "标准正文"
        else:
            # 对于检测报告的分类逻辑
            # 检查是否是表注说明的开始
            if self.re_note.match(clean_text): return "表注说明_起点"
            # 如果在表注模式下，且是列表项，则为延续
            if is_in_note_mode and self.re_list_item.match(clean_text): return "表注说明_延续"
            # 检查标题层级
            if self.re_h3.match(clean_text): return "三级标题"
            if self.re_h2.match(clean_text): return "二级标题"
            if self.re_h1.match(clean_text): return "一级标题"
            
            # 如果在检测依据模式下，返回无缩进正文
            if is_in_basis_mode: return "无缩进正文"
            # 检查是否不需要缩进
            if self.re_no_indent.match(clean_text): return "无缩进正文"
            
            # 默认返回标准正文
            return "标准正文"


# ==================== 板块 2：交互与参数获取 (UI) ====================

def get_user_params(file_name):
    """
    获取用户输入的参数。

    通过弹窗界面让用户选择报告类型和需要跳过的页码。

    参数:
        file_name (str): 当前正在处理的文件名。

    返回:
        dict or None: 包含'report_type'和'skip_pages'的字典，如果用户取消则返回None。
    """
    # 创建隐藏的主窗口
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # 窗口置顶
    prompt_base = f"当前排版文件：{file_name}\n\n"

    # 第一个对话框：选择报告类型
    type_input = simpledialog.askstring(
        "1/2", 
        f"{prompt_base}请选择处理的报告类型：\n1 - 检测报告\n2 - 鉴定报告", 
        initialvalue="", 
        parent=root
    )
    if not type_input: return None  # 用户取消
    # 根据输入确定报告类型
    report_type = "鉴定报告" if type_input == "2" else "检测报告"

    # 第二个对话框：输入跳过页码
    skip_input = simpledialog.askstring(
        "2/2", 
        f"{prompt_base}请输入需要跳过的页码（如封面、资质、目录等）：\n页码间用逗号分隔，例如：1,2,3\n若无跳过页，直接点确定。", 
        initialvalue="1,2,3,4", 
        parent=root
    )
    if skip_input is None: return None  # 用户取消
    
    # 解析跳过页码
    skip_pages = []
    if skip_input.strip():
        normalized = skip_input.replace("，", ",")  # 支持中文逗号
        skip_pages = [int(p.strip()) for p in normalized.split(",") if p.strip().isdigit()]
        
    root.destroy()  # 销毁窗口
    return {"report_type": report_type, "skip_pages": skip_pages}


def final_check_summary(file_name, params):
    """
    显示最终确认摘要。

    显示用户选择的参数，并让用户确认是否执行排版。

    参数:
        file_name (str): 文件名。
        params (dict): 包含参数的字典。

    返回:
        bool: 用户是否确认执行。
    """
    # 创建隐藏的主窗口
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    # 构建摘要信息
    summary = (
        f"📂 目标文件: {file_name}\n"
        f"--------------------------\n"
        f"报告类型: {params['report_type']}\n"
        f"跳过页码: {params['skip_pages'] if params['skip_pages'] else '无'}\n"
        "--------------------------\n"
        "确认执行后，将调用静默备份并启动全量正文排版。"
    )
    # 显示确认对话框
    confirm = messagebox.askyesno("请最终确认排版清单", summary, parent=root)
    root.destroy()
    return confirm


# ==================== 板块 3：格式引擎 ====================

def apply_paragraph_format(para, style_config, para_type):
    """
    应用段落格式。

    根据样式配置设置段落的字体、对齐、间距等格式。

    参数:
        para: Word段落对象。
        style_config (dict): 样式配置字典。
        para_type (str): 段落类型，用于调试。

    返回:
        None
    """
    try:
        # 设置字体格式
        f = para.Range.Font
        
        # 第一步：先设置英文字体（作为底层主属性，必须先执行）
        eng_font = style_config.get("english_font", "Times New Roman")
        f.Name = eng_font
        f.NameAscii = eng_font
        
        # 第二步：后设置中文字体（远东字体，覆盖于主属性之上）
        f.NameFarEast = style_config.get("chinese_font", "宋体")  
        
        # 第三步：设置其他通用属性
        f.Size = style_config.get("font_size", 12.0)  
        f.Bold = style_config.get("bold", False)

        # 设置段落格式
        pf = para.Format
        pf.Alignment = style_config.get("alignment", 3)  # 对齐方式 (1左 2中 3右)
        pf.OutlineLevel = style_config.get("outline_level", 10)  # 大纲级别
        pf.SpaceBefore = style_config.get("space_before", 0) * 12   # 段前间距 (磅)
        pf.SpaceAfter = style_config.get("space_after", 0) * 12  # 段后间距 (磅)

        # 设置行距
        ls_rule = style_config.get("line_spacing_rule", 5)
        if ls_rule == 1:  # 单倍行距
            pf.LineSpacingRule = 1
        elif ls_rule == 0:  # 最小行距
            pf.LineSpacingRule = 0
        else:  # 多倍行距
            pf.LineSpacingRule = 5
            pf.LineSpacing = style_config.get("line_spacing", 1.5) * 12  # 行距值
        pf.DisableLineHeightGrid = False  # 启用行高网格
        
        # 设置缩进
        pf.CharacterUnitRightIndent = style_config.get("right_indent", 0)  # 右缩进
        char_first = style_config.get("first_line_indent", 0)  # 首行缩进
        pf.CharacterUnitFirstLineIndent = char_first
        
        # 清零绝对位移，避免Word的列表引起的位移问题
        if char_first == 0:
            pf.FirstLineIndent = 0  # 首行绝对缩进
        pf.CharacterUnitLeftIndent = 0  # 左缩进
        pf.LeftIndent = 0  # 左绝对缩进
        
        # 如果配置中有绝对磅值缩进，则覆盖
        if style_config.get("left_indent_pt", 0) != 0:
            pf.LeftIndent = style_config["left_indent_pt"]  # 左缩进磅值
        if style_config.get("first_line_indent_pt", 0) != 0:
            pf.FirstLineIndent = style_config["first_line_indent_pt"]  # 首行缩进磅值
                
    except Exception as e:
        # 如果出错，静默跳过（避免中断整个过程）
        pass


def process_document_body(app, params):
    """
    处理文档正文。

    遍历文档的所有段落，分类并应用相应的格式。

    参数:
        app: Word应用程序对象。
        params (dict): 包含'report_type'和'skip_pages'的参数字典。

    返回:
        tuple: (成功处理的段落数, 跳过的段落数)
    """
    report_type = params["report_type"]
    skip_pages = params["skip_pages"]
    
    # 加载样式配置
    full_config = load_style_config(report_type)
    # 创建段落分类器
    classifier = ParagraphClassifier()

    # 获取活动文档
    doc = app.ActiveDocument
    paragraphs = doc.Paragraphs
    total_paras = paragraphs.Count  # 总段落数
    
    # 创建进度条窗口
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

    # 关闭屏幕更新以提高性能
    app.ScreenUpdating = False 
    success_count = 0  # 成功处理的段落数
    skipped_count = 0  # 跳过的段落数
    # 模式标志，用于跟踪上下文
    note_mode = basis_mode = conclusion_mode = False 
    
    try:
        # 遍历所有段落
        for i in range(1, total_paras + 1):
            # 每10个段落或最后一个更新进度条
            if i % 10 == 0 or i == total_paras:
                bar['value'] = i
                progress_label.config(text=f"正在排版: {i}/{total_paras} 段")
                pg_root.update()
                
            para = paragraphs.Item(i)
            
            try:
                # 获取段落所在页码
                page_num = para.Range.Information(3) 
                if page_num in skip_pages:
                    skipped_count += 1
                    continue  # 跳过指定页码
            except:
                pass  # 如果获取页码失败，继续处理

            # 跳过目录相关的段落
            if para.Range.Information(12) or "目录" in para.Style.NameLocal or "TOC" in para.Style.NameLocal:
                skipped_count += 1
                continue
                
            # 获取段落文本和列表字符串
            text = para.Range.Text
            list_str = para.Range.ListFormat.ListString 

            clean_t = text.strip()
            # 检查是否是检测依据或处理建议的标题
            is_basis_header = classifier.re_basis_title.match(clean_t) 
            is_suggest_header = classifier.re_suggest_title.match(clean_t.replace(" ", ""))

            # 检查段落是否包含图片
            has_image = False
            try:
                if para.Range.InlineShapes.Count > 0:
                    has_image = True
            except:
                pass

            # 分类段落
            if has_image:
                para_type = "图片"
            else:
                para_type = classifier.classify(text, list_str, note_mode, basis_mode, conclusion_mode, report_type)

            # 更新模式标志
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
                
            # 跳过已经是列表项的标准正文
            if para_type == "标准正文" and re.match(r'^(\d+[.、）\)]|[①②③④⑤⑥⑦⑧⑨⑩])', text.strip()):
                continue
                
            # 将符号列表转换为无缩进正文
            if para_type == "标准正文" and re.match(r'^[\s]*[•\-*]\s+', text.strip()):
                para_type = "无缩进正文"
                
            # 如果配置中有此类型，则应用格式
            if para_type in full_config:
                apply_paragraph_format(para, full_config[para_type], para_type)
                success_count += 1

    finally:
        # 清理：销毁进度条窗口，恢复屏幕更新
        if 'pg_root' in locals():
            pg_root.destroy()
        app.ScreenUpdating = True 
        
    return success_count, skipped_count


# ==================== 最终主控流 ====================

if __name__ == "__main__":
    """
    主程序入口。

    检测Word应用程序，获取用户参数，执行排版任务。
    """
    try:
        # 尝试连接到Word应用程序
        app = win32com.client.GetActiveObject("Word.Application") 
    except:
        try:
            # 如果Word失败，尝试WPS
            app = win32com.client.GetActiveObject("KWPS.Application") 
        except:
            app = None  # 都没有找到

    if not app:
        err_root = tk.Tk()
        err_root.withdraw()
        err_root.attributes('-topmost', True)
        messagebox.showerror("运行阻断", "未检测到运行中的 WPS 或 Word 程序。\n\n请先打开需要排版的报告文档！", parent=err_root)
        err_root.destroy()
    else:
        # 隐患拦截：拦截未保存的新建文档，防止静默备份引发异常
        if app.ActiveDocument.Path == "":
            err_root = tk.Tk()
            err_root.withdraw()
            err_root.attributes('-topmost', True)
            messagebox.showwarning("操作阻断", "该文档尚未保存到本地硬盘。\n请先手动保存一次（Ctrl+S）后再执行排版引擎！", parent=err_root)
            err_root.destroy()
        else:
            current_file = app.ActiveDocument.Name
        # 获取用户输入的参数
        run_params = get_user_params(current_file)
        
        if run_params is None:
            # 用户取消了操作
            print("【取消】用户取消了操作，程序终止。")
        else:
            # 显示确认摘要
            if final_check_summary(current_file, run_params):
                print("正在调用外部模块进行静默备份...")
                # 调用备份函数
                if backup_current_document(app):
                    # 执行正文排版
                    succ_cnt, skip_cnt = process_document_body(app, run_params)
                    
                    # 显示完成信息
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
                    # 备份失败，显示错误
                    err_root = tk.Tk()
                    err_root.withdraw()
                    err_root.attributes('-topmost', True)
                    messagebox.showerror(
                        "安全熔断", 
                        "⚠️ 备份模块(file_utils)返回失败信号！\n\n为防止原文件损坏，排版程序已自动终止。\n请检查当前文档是否已保存，或查看后台报错日志。", 
                        parent=err_root
                    )
                    err_root.destroy()