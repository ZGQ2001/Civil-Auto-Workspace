"""
===============================================================================
脚本名称：文档全局括号半全角专项纠偏引擎 (bracket_format.py)
作者：ZGQ
功能概述：
    基于 Word/WPS 通配符引擎，全局规范括号的全/半角格式。
    已集成：国标/行标代号专属全角保护、书名号精准兜底、纯数字序号防呆。
===============================================================================
"""
import os
import sys
import win32com.client
import tkinter as tk
from tkinter import messagebox

# 挂载外部备份模块（与正文、表格脚本共用）
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_utils_backup import backup_current_document

def process_brackets(app, doc_name):
    # 步骤1：防呆确认弹窗
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    prompt = (
        f"📂 当前文件：{doc_name}\n\n"
        "是否执行【括号半全角专项纠偏】处理？\n\n"
        "处理逻辑：\n"
        "1. 全局小括号建立基准（统一转全角）\n"
        "2. 纯技术参数转半角\n"
        "3. 书名号及国标规范代号锁定全角（如 GB/T, JGJ 等）\n"
        "4. 纯数字层级序号锁定全角\n"
        "5. 锚定排除：包含“第”字的编号保留半角\n\n"
        "确认后将调用静默备份并开始执行。"
    )
    if not messagebox.askyesno("括号专项引擎启动", prompt, parent=root):
        root.destroy()
        return False
    root.destroy()

    # 步骤2：执行统一静默备份
    print("正在调用外部模块进行静默备份...")
    if not backup_current_document(app):
        err_root = tk.Tk()
        err_root.withdraw()
        err_root.attributes('-topmost', True)
        messagebox.showerror("安全熔断", "⚠️ 备份模块返回失败信号！操作已终止。", parent=err_root)
        err_root.destroy()
        return False

    # 步骤3：核心替换逻辑
    doc = app.ActiveDocument
    rng = doc.Content
    fnd = rng.Find
    
    # 构建正则通配符替换规则集
    rules = [
        # 1. 清洗环境基准：全部抹平为全角
        {"f": "(", "r": "（", "wc": False},
        {"f": ")", "r": "）", "wc": False},
        
        # 2. 波浪号纠偏：修复输入法打出的全角波浪号
        {"f": "～", "r": "~", "wc": False},
        
        # 3. 提取技术参数转为半角 (涵盖尺寸、单位、公式等，- 必须放在方括号最末尾)
        {"f": r"（([a-zA-Z0-9 .,/\\_~—–:%°+=±×÷·　-]@)）", "r": r"(\1)", "wc": True},
        
        # 4. 特例纠偏 - 标准代号强制全角（针对 GB/T, JGJ, DB 等国标/行标规范）
        # 特征：以至少2个大写字母开头，包含大写字母、数字、斜杠、空格、点号、冒号、中划线
        {"f": r"\(([A-Z]{2,}[A-Z0-9/ \.:　-]@)\)", "r": r"（\1）", "wc": True},
        
        # 5. 特例纠偏 - 书名号标准代号锁定全角（兜底无前缀的标准或企标 Q/ 等）
        # [!)] 表示匹配除了半角右括号之外的任意字符，避开 Word 对反斜杠的底层转义限制
        {"f": r"》\(([!)]@)\)", "r": r"》（\1）", "wc": True},
        {"f": r"》 \(([!)]@)\)", "r": r"》 （\1）", "wc": True},
        {"f": r"》　\(([!)]@)\)", "r": r"》　（\1）", "wc": True},
        
        # 6. 特例纠偏 - 中文纯数字层级序号锁定全角 (1), (2) -> （1）, （2）
        {"f": r"\(([0-9]@)\)", "r": r"（\1）", "wc": True},
        
        # 7. 锚定排除 - 编号“第(xxx)”保留半角，抗击第6步误伤
        {"f": r"第（([0-9]@)）", "r": r"第(\1)", "wc": True},
        {"f": r"第 （([0-9]@)）", "r": r"第 (\1)", "wc": True},
        {"f": r"第　（([0-9]@)）", "r": r"第　(\1)", "wc": True}
    ]

    app.ScreenUpdating = False
    try:
        total = len(rules)
        for i, rule in enumerate(rules):
            app.StatusBar = f"正在处理括号规范，进度: ({i+1}/{total}) ..."
            fnd.ClearFormatting()
            fnd.Replacement.ClearFormatting()
            
            fnd.Execute(
                FindText=rule["f"],
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=rule["wc"],
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=1,  # wdFindContinue
                Format=False,
                ReplaceWith=rule["r"],
                Replace=2  # wdReplaceAll
            )
            
        # 成功反馈
        succ_root = tk.Tk()
        succ_root.withdraw()
        succ_root.attributes('-topmost', True)
        messagebox.showinfo(
            "执行完毕", 
            f"✅ 文档括号专项纠偏完成！\n\n"
            f"- 技术参数：已转半角\n"
            f"- 规范代号：已锁定全角\n"
            f"- 层级序号：已锁定全角\n"
            f"- 锚定排除：第() 已保留半角",
            parent=succ_root
        )
        succ_root.destroy()
        return True

    except Exception as e:
        err_root = tk.Tk()
        err_root.withdraw()
        err_root.attributes('-topmost', True)
        messagebox.showerror("运行期错误拦截", f"架构运行遭遇异常抛出：\n{str(e)}", parent=err_root)
        err_root.destroy()
        return False
        
    finally:
        app.ScreenUpdating = True
        app.StatusBar = "就绪"

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
        # 隐患拦截：拦截未保存的新建文档，防止静默备份引发异常
        if app.ActiveDocument.Path == "":
            err_root = tk.Tk()
            err_root.withdraw()
            err_root.attributes('-topmost', True)
            messagebox.showwarning("操作阻断", "该文档尚未保存到本地硬盘。\n请先手动保存一次（Ctrl+S）后再执行排版引擎！", parent=err_root)
            err_root.destroy()
        else:
            process_brackets(app, app.ActiveDocument.Name)