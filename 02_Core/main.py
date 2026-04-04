"""
===============================================================================
脚本名称：主程序控制台入口 (main.py)
作者: ZGQ
功能概述：
    统一启动面板。
    负责调度和异步唤醒各个独立的自动化子模块。
===============================================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os

class MainDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("报告自动化 V2.0")
        self.root.geometry("450x520")
        self.root.resizable(False, False)  # 锁定窗口大小，保持界面规整
        
        # 设置全局样式
        style = ttk.Style()
        style.configure("TButton", font=("Microsoft YaHei", 11), padding=10)
        style.configure("TLabel", font=("Microsoft YaHei", 12, "bold"))
        style.configure("Header.TLabel", font=("Microsoft YaHei", 16, "bold"), foreground="#2c3e50")

        self.setup_ui()

    def setup_ui(self):
        # 顶部标题栏
        header_frame = tk.Frame(self.root, bg="#ecf0f1", pady=20)
        header_frame.pack(fill=tk.X)
        ttk.Label(header_frame, text="报告自动化排版矩阵", style="Header.TLabel", background="#ecf0f1").pack()
        ttk.Label(header_frame, text="Unified Automation Dashboard", font=("Arial", 9), background="#ecf0f1", foreground="#7f8c8d").pack()

        # 核心功能区
        main_frame = tk.Frame(self.root, pady=20, padx=40)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 按钮列表：包含文本描述和对应的脚本文件名
        modules = [
            ("报告正文排版引擎", "body_format.py"),
            ("报告表格排版引擎", "table_format.py"),
            ("全局括号半全角纠偏", "bracket_format.py"),
            ("交叉引用格式修复", "fix_cross_ref.py"),
            ("Word转PDF与合并工具", "word2pdf.py"),
            ("PDF坐标拾取导出器", "pdf_coordinate_picker.py")
        ]

        # 循环生成启动按钮
        for text, script in modules:
            btn = ttk.Button(
                main_frame, 
                text=text, 
                command=lambda s=script: self.launch_module(s)
            )
            btn.pack(fill=tk.X, pady=8)

        # 底部状态栏
        footer_frame = tk.Frame(self.root, pady=10)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(footer_frame, text="状态: 引擎就绪 | 依赖 JSON 规则库驱动", font=("Microsoft YaHei", 9)).pack()

    def launch_module(self, script_name):
        """
        异步调用子脚本
        """
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
        
        if not os.path.exists(script_path):
            messagebox.showerror("模块缺失", f"系统未能找到执行文件：\n{script_path}")
            return
            
        try:
            # 使用同级 Python 解释器异步启动子脚本，避免阻塞主面板
            subprocess.Popen([sys.executable, script_path])
        except Exception as e:
            messagebox.showerror("启动失败", f"无法唤醒模块 {script_name}，底层错误：\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MainDashboard(root)
    
    # 将窗口居中显示
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()