"""
===============================================================================
脚本名称：全局现代 UI 组件库 (ui_components.py)
功能概述：
    为各类处理脚本提供统一的、现代化的图形交互界面组件。
    包含：确认弹窗、进度条控制台、信息反馈弹窗。
===============================================================================
"""
import tkinter as tk
from tkinter import ttk

# 全局视觉规范
UI_COLORS = {
    "primary": "#0078d4",      # 主题蓝
    "background": "#f5f5f5",   # 浅灰底色
    "surface": "#ffffff",      # 纯白面板
    "text_main": "#333333",    # 主文本
    "text_sub": "#666666",     # 次文本
    "danger": "#d32f2f",       # 警告红
    "danger_bg": "#ffebee"     # 警告红底
}

class BaseDialog:
    """弹窗基类，处理居中和基础属性"""
    def __init__(self, title, width, height):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry(f"{width}x{height}")
        self.root.configure(bg=UI_COLORS["background"])
        self.root.attributes('-topmost', True)
        
        # 屏幕居中计算
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"+{x}+{y}")

class ModernConfirmDialog(BaseDialog):
    """现代确认弹窗"""
    def __init__(self, title, message, sub_message=""):
        super().__init__(title, 450, 250)
        self.result = False
        
        main_frame = tk.Frame(self.root, bg=UI_COLORS["surface"])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(main_frame, text=message, font=("微软雅黑", 11, "bold"), 
                 bg=UI_COLORS["surface"], fg=UI_COLORS["text_main"], justify=tk.LEFT).pack(anchor=tk.W, pady=(10, 5))
        
        if sub_message:
            tk.Label(main_frame, text=sub_message, font=("微软雅黑", 9), 
                     bg=UI_COLORS["surface"], fg=UI_COLORS["text_sub"], justify=tk.LEFT).pack(anchor=tk.W)
            
        btn_frame = tk.Frame(self.root, bg=UI_COLORS["background"])
        btn_frame.pack(fill=tk.X, pady=15, padx=20)
        
        tk.Button(btn_frame, text="确认执行", bg=UI_COLORS["primary"], fg="white", 
                  font=("微软雅黑", 9, "bold"), width=12, command=self._confirm).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="取消", bg="#e1e1e1", fg=UI_COLORS["text_main"], 
                  font=("微软雅黑", 9), width=10, command=self._cancel).pack(side=tk.RIGHT)

    def _confirm(self):
        self.result = True
        self.root.destroy()
        
    def _cancel(self):
        self.result = False
        self.root.destroy()
        
    def show(self):
        self.root.mainloop()
        return self.result

class ModernProgressConsole(BaseDialog):
    """现代进度控制台，带熔断机制"""
    def __init__(self, title, max_val):
        super().__init__(title, 400, 180)
        self.is_cancelled = False
        
        tk.Label(self.root, text="引擎运行中...", font=("微软雅黑", 12, "bold"), 
                 bg=UI_COLORS["background"], fg=UI_COLORS["text_main"]).pack(pady=(20, 10))
        
        self.lbl_status = tk.Label(self.root, text="初始化...", bg=UI_COLORS["background"], fg=UI_COLORS["text_sub"])
        self.lbl_status.pack()
        
        # 自定义进度条样式
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Custom.Horizontal.TProgressbar", thickness=8, background=UI_COLORS["primary"], 
                        troughcolor="#e1e1e1", borderwidth=0)
        
        self.bar = ttk.Progressbar(self.root, length=320, mode='determinate', 
                                   maximum=max_val, style="Custom.Horizontal.TProgressbar")
        self.bar.pack(pady=15)
        
        tk.Button(self.root, text="紧急停止", command=self._stop, bg=UI_COLORS["danger_bg"], 
                  fg=UI_COLORS["danger"], relief=tk.FLAT, width=12).pack()
        
        self.root.protocol("WM_DELETE_WINDOW", self._stop)
        self.root.update()

    def update_progress(self, current_val, status_text):
        if self.root.winfo_exists():
            self.bar['value'] = current_val
            self.lbl_status.config(text=status_text)
            self.root.update()

    def _stop(self):
        self.is_cancelled = True
        self.lbl_status.config(text="正在安全中断连接，请稍候...", fg=UI_COLORS["danger"])
        self.root.update()

    def close(self):
        if self.root.winfo_exists():
            self.root.destroy()

class ModernInfoDialog(BaseDialog):
    """现代信息反馈弹窗"""
    def __init__(self, title, message):
        super().__init__(title, 400, 200)
        
        main_frame = tk.Frame(self.root, bg=UI_COLORS["surface"])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(main_frame, text=message, font=("微软雅黑", 10), 
                 bg=UI_COLORS["surface"], fg=UI_COLORS["text_main"], justify=tk.LEFT).pack(pady=20)
        
        tk.Button(self.root, text="关闭", bg=UI_COLORS["primary"], fg="white", 
                  font=("微软雅黑", 9), width=15, command=self.root.destroy).pack(pady=15)

    def show(self):
        self.root.mainloop()