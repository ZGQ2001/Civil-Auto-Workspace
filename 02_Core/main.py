"""
===============================================================================
脚本名称：主程序控制台入口 (main.py)
作者: ZGQ
功能概述：
    统一启动面板。
    负责调度和异步唤醒各个独立的自动化子模块。
===============================================================================
"""

import os
import sys
import subprocess
import customtkinter as ctk

class MainDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("工程自动化程序 V0.1.0")
        self.root.geometry("450x520")
        self.root.resizable(False, False)
        
        # 顶部标题区
        self.header = ctk.CTkFrame(self.root, fg_color="transparent")
        self.header.pack(fill="x", pady=(30, 20))
        ctk.CTkLabel(self.header, text="工程自动化程序", font=("微软雅黑", 22, "bold"), text_color="#0078d4").pack()
        ctk.CTkLabel(self.header, text="Automation Dashboard", font=("Consolas", 12), text_color="gray50").pack()

        # 核心按钮区
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=15)
        self.main_frame.pack(fill="both", expand=True, padx=40, pady=10)

        modules = [
            ("报告正文排版引擎", "body_format.py"),
            ("报告表格排版引擎", "table_format.py"),
            ("全局括号半全角纠偏", "bracket_format.py"),
            ("交叉引用格式修复", "fix_cross_ref.py"),
            ("文档转换小工具", "word2pdf.py"),
        ]

        for text, script in modules:
            btn = ctk.CTkButton(self.main_frame, text=text, font=("微软雅黑", 14, "bold"), height=40,
                                command=lambda s=script: self.launch_module(s))
            btn.pack(fill="x", padx=20, pady=12)

        # 底部状态
        ctk.CTkLabel(self.root, text=" 作者: ZGQ ", font=("微软雅黑", 11), text_color="gray60").pack(side="bottom", pady=15)

    def launch_module(self, script_name):
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
        if not os.path.exists(script_path):
            print(f"缺失模块: {script_path}")
            return
        subprocess.Popen([sys.executable, script_path])

if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    app = MainDashboard(root)
    root.mainloop()