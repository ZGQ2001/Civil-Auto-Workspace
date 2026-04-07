"""
===============================================================================
脚本名称：全局现代 UI 组件库 - 工业级防吞窗版 (ui_components.py)
功能概述：
    采用 Singleton (单例) 隐藏根窗口 + Toplevel 架构。
    彻底解决由于连续创建/销毁 CTk 实例导致的“弹窗被系统静默吞噬”或闪退 Bug。
===============================================================================
"""
import customtkinter as ctk
import os
import json
from tkinter import filedialog

# 全局基础设置
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# 【核心护城河】：全局唯一隐藏主窗口
_global_root = None

def _get_root():
    global _global_root
    if _global_root is None or not _global_root.winfo_exists():
        _global_root = ctk.CTk()
        _global_root.withdraw() # 永远隐藏，仅做锚点
    return _global_root

class BaseDialog:
    """弹窗基类，处理居中和基础属性"""
    def __init__(self, title, width, height):
        # 【架构升级】：所有弹窗作为子窗口依附于隐藏的根窗口
        self.root = ctk.CTkToplevel(_get_root())
        self.root.title(title)
        self.root.geometry(f"{width}x{height}")
        
        # 强制置顶与焦点获取，彻底防止弹窗被 Word 挡住或被吞掉
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.focus_force()
        self.root.resizable(False, False)
        
        # 屏幕居中计算
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"+{x}+{y}")

class ModernConfirmDialog(BaseDialog):
    """现代确认弹窗"""
    def __init__(self, title, message, sub_message=""):
        super().__init__(title, 500, 320)
        self.result = False
        
        self.frame = ctk.CTkFrame(self.root, corner_radius=10)
        self.frame.pack(fill="both", expand=True, padx=25, pady=(25, 10))
        
        self.lbl_msg = ctk.CTkLabel(self.frame, text=message, font=("微软雅黑", 14, "bold"), justify="center")
        self.lbl_msg.pack(pady=(25, 5), padx=20)
        
        if sub_message:
            self.lbl_sub = ctk.CTkLabel(self.frame, text=sub_message, font=("微软雅黑", 12), 
                                        text_color="gray60", justify="center")
            self.lbl_sub.pack(pady=10, padx=20)
            
        self.btn_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.btn_frame.pack(pady=25)
        
        self.btn_confirm = ctk.CTkButton(self.btn_frame, text="确定执行", font=("微软雅黑", 13, "bold"), 
                                         width=160, height=45, command=self._confirm)
        self.btn_confirm.pack(side="left", padx=10)
        
        self.btn_cancel = ctk.CTkButton(self.btn_frame, text="取消", font=("微软雅黑", 13), 
                                        width=120, height=45, fg_color="transparent", 
                                        border_width=1, text_color=("gray10", "gray90"), command=self._cancel)
        self.btn_cancel.pack(side="left", padx=10)

    def _confirm(self):
        self.result = True
        self.root.destroy()
        
    def _cancel(self):
        self.result = False
        self.root.destroy()
        
    def show(self):
        # 模态阻塞：接管底层事件队列，防止主程序提前偷跑
        self.root.grab_set()
        self.root.master.wait_window(self.root)
        return self.result

class ModernProgressConsole(BaseDialog):
    """现代进度控制台"""
    def __init__(self, title, max_val):
        super().__init__(title, 420, 220)
        self.is_cancelled = False
        self.max_val = max_val
        
        self.lbl_title = ctk.CTkLabel(self.root, text="引擎运行中...", font=("微软雅黑", 16, "bold"))
        self.lbl_title.pack(pady=(25, 5))
        
        self.lbl_status = ctk.CTkLabel(self.root, text="初始化...", font=("Consolas", 11), text_color="gray60")
        self.lbl_status.pack()
        
        self.bar = ctk.CTkProgressBar(self.root, width=340, height=12, corner_radius=6)
        self.bar.pack(pady=20)
        self.bar.set(0)
        
        self.btn_stop = ctk.CTkButton(self.root, text="紧急停止", font=("微软雅黑", 12, "bold"), 
                                      fg_color="#d32f2f", hover_color="#b71c1c", width=140, height=40, command=self._stop)
        self.btn_stop.pack(pady=5)
        
        self.root.protocol("WM_DELETE_WINDOW", self._stop)
        self.root.update()

    def update_progress(self, current_val, status_text):
        if self.root.winfo_exists():
            progress_ratio = current_val / self.max_val if self.max_val > 0 else 0
            self.bar.set(progress_ratio)
            self.lbl_status.configure(text=status_text)
            self.root.update()

    def _stop(self):
        self.is_cancelled = True
        self.lbl_status.configure(text="正在中止安全环境...", text_color="#d32f2f")
        self.btn_stop.configure(state="disabled")
        self.root.update()

    def close(self):
        if self.root.winfo_exists():
            self.root.destroy()

class ModernInfoDialog(BaseDialog):
    """现代信息反馈弹窗"""
    def __init__(self, title, message):
        super().__init__(title, 550, 420)
        
        self.frame = ctk.CTkFrame(self.root, corner_radius=10)
        self.frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        inner_content = ctk.CTkFrame(self.frame, fg_color="transparent")
        inner_content.pack(expand=True)
        
        self.lbl_msg = ctk.CTkLabel(inner_content, text=message, font=("微软雅黑", 13), 
                                    justify="left")
        self.lbl_msg.pack(pady=20, padx=20)
        
        self.btn_close = ctk.CTkButton(self.root, text="确定", font=("微软雅黑", 14, "bold"), 
                                       width=180, height=48, command=self.root.destroy)
        self.btn_close.pack(pady=(0, 25))

    def show(self):
        self.root.grab_set()
        self.root.master.wait_window(self.root)

class ModernParamDialog(BaseDialog):
    """现代参数输入面板"""
    def __init__(self, title, file_name, show_width=False):
        super().__init__(title, 500, 380 if show_width else 340)
        self.params = None
        
        ctk.CTkLabel(self.root, text=f"📄 目标文件: {file_name}", font=("微软雅黑", 13, "bold")).pack(pady=(25, 15))
        
        # 必须绑定 master=self.root，防止变量脱离作用域
        self.type_var = ctk.StringVar(master=self.root, value="检测报告")
        
        # ---------------- 核心排版：Grid 网格布局 ----------------
        form_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        form_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # 配置列权重：左右两列为弹簧列拉伸，中间两列被强制居中对齐
        form_frame.grid_columnconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=0, minsize=100) 
        form_frame.grid_columnconfigure(2, weight=0, minsize=220) 
        form_frame.grid_columnconfigure(3, weight=1)

        row_idx = 0
        
        # 第一行：报告类型
        ctk.CTkLabel(form_frame, text="报告类型:", font=("微软雅黑", 12)).grid(row=row_idx, column=1, sticky="e", pady=12, padx=(0, 15))
        radio_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        radio_frame.grid(row=row_idx, column=2, sticky="w", pady=12)
        ctk.CTkRadioButton(radio_frame, text="检测报告", variable=self.type_var, value="检测报告").pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(radio_frame, text="鉴定报告", variable=self.type_var, value="鉴定报告").pack(side="left")
        row_idx += 1

        # 第二行：表格宽度（按需渲染）
        self.width_entry = None
        if show_width:
            # 【核心修改 1】：文案剥离，强行缩减为4个字，与上下保持绝对物理等长
            ctk.CTkLabel(form_frame, text="表格宽度:", font=("微软雅黑", 12)).grid(row=row_idx, column=1, sticky="e", pady=12, padx=(0, 15))
            
            # 【核心修改 2】：创建一个内部小容器，用来横向包裹“输入框”和“%”符号
            width_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
            width_frame.grid(row=row_idx, column=2, sticky="w", pady=12)
            
            # 输入框长度稍微缩短至 195，给后面的 % 腾出视觉空间，保证整体 220 的总宽度
            self.width_entry = ctk.CTkEntry(width_frame, width=195)
            self.width_entry.insert(0, "95")
            self.width_entry.pack(side="left")
            
            # 【核心修改 3】：把 % 作为后缀单位，贴在输入框的右侧
            ctk.CTkLabel(width_frame, text="%", font=("微软雅黑", 12)).pack(side="left", padx=(5, 0))
            row_idx += 1

        # 第三行：跳过页码
        ctk.CTkLabel(form_frame, text="跳过页码:", font=("微软雅黑", 12)).grid(row=row_idx, column=1, sticky="e", pady=12, padx=(0, 15))
        self.skip_entry = ctk.CTkEntry(form_frame, placeholder_text="如: 1,2,3 (留空全排)", width=220)
        self.skip_entry.insert(0, "1,2,3,4")
        self.skip_entry.grid(row=row_idx, column=2, sticky="w", pady=12)
        # --------------------------------------------------------

        self.btn_confirm = ctk.CTkButton(self.root, text="确定", command=self._confirm, 
                                         font=("微软雅黑", 14, "bold"), width=180, height=48)
        self.btn_confirm.pack(pady=(20, 30))

    def _confirm(self):
        skips = []
        if self.skip_entry.get().strip():
            skips = [int(p.strip()) for p in self.skip_entry.get().replace("，", ",").split(",") if p.strip().isdigit()]
        
        self.params = {"report_type": self.type_var.get(), "skip_pages": skips}
        if self.width_entry:
            self.params["width"] = int(self.width_entry.get() or "100")
        self.root.destroy()

    def show(self):
        self.root.grab_set()
        self.root.master.wait_window(self.root)
        return self.params
    
class ModernHandwriteDialog(BaseDialog):
    """现代仿生手写生成器主控面板"""
    def __init__(self, title="仿生手写配置台"):
        # 弹窗尺寸需要比普通参数面板大，因为配置项很多
        super().__init__(title, 750, 800)
        self.config_data = None # 用于存储最终点击“下一步”时返回的数据
        
        # 配置文件保存路径（存放在代码同级目录）
        self.config_file = "handwrite_config.json"

        # 【核心变量绑定】：必须绑定 master=self.root，防止变量脱离作用域导致报错或数据不更新
        self.var_excel_path = ctk.StringVar(master=self.root)
        self.var_json_path = ctk.StringVar(master=self.root)
        self.var_img_path = ctk.StringVar(master=self.root)
        self.var_font_dir = ctk.StringVar(master=self.root)
        self.var_output_dir = ctk.StringVar(master=self.root)
        
        self.var_sheet_name = ctk.StringVar(master=self.root, value="Sheet2")
        self.var_font_scale = ctk.DoubleVar(master=self.root, value=1.68)
        self.var_y_offset = ctk.DoubleVar(master=self.root, value=-1.5)
        self.var_spacing = ctk.IntVar(master=self.root, value=-5)

        self._build_ui()
        self.load_config() # 启动时自动读取上次配置

    def _build_ui(self):
        """构建界面的总指挥"""
        # ================= 板块 1：文件与路径配置 =================
        frame_files = ctk.CTkFrame(self.root, corner_radius=10)
        frame_files.pack(pady=15, padx=20, fill="x")
        
        ctk.CTkLabel(frame_files, text="📂 核心文件配置", font=("微软雅黑", 15, "bold")).pack(pady=(15, 5))

        self._add_file_selector(frame_files, "Excel 数据源:", self.var_excel_path, file_types=[("Excel", "*.xlsx *.xlsm")])
        self._add_file_selector(frame_files, "JSON 坐标库:", self.var_json_path, file_types=[("JSON", "*.json")])
        self._add_file_selector(frame_files, "空白底图文件:", self.var_img_path, file_types=[("图片", "*.png *.jpg")])
        self._add_dir_selector(frame_files, "手写字体目录:", self.var_font_dir)
        self._add_dir_selector(frame_files, "PDF 输出目录:", self.var_output_dir)

        # ================= 板块 2：全局视觉参数 =================
        frame_params = ctk.CTkFrame(self.root, corner_radius=10)
        frame_params.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(frame_params, text="🎨 全局视觉微调", font=("微软雅黑", 15, "bold")).pack(pady=(15, 5))

        self._add_entry_row(frame_params, "目标 Sheet 名称:", self.var_sheet_name)
        self._add_slider_row(frame_params, "全局字号缩放 (倍):", self.var_font_scale, 1.0, 2.5)
        self._add_slider_row(frame_params, "纵向偏移补偿 (px):", self.var_y_offset, -15.0, 15.0)
        self._add_slider_row(frame_params, "字距收缩程度:", self.var_spacing, -15, 5, is_int=True)

        # ================= 板块 3：状态管理与执行 =================
        frame_actions = ctk.CTkFrame(self.root, fg_color="transparent")
        frame_actions.pack(pady=20, fill="x")

        btn_save = ctk.CTkButton(frame_actions, text="💾 保存参数配置", width=140, command=self.save_config)
        btn_save.pack(side="left", padx=(40, 10))

        btn_load = ctk.CTkButton(frame_actions, text="🔄 重新加载配置", width=140, fg_color="#F39C12", hover_color="#D68910", command=self.load_config)
        btn_load.pack(side="left", padx=10)

        # 第一阶段的终点，点击进入第二阶段（返回收集到的数据）
        btn_next = ctk.CTkButton(frame_actions, text="下一步：配置数据映射 ➡️", width=180, 
                                 fg_color="#27AE60", hover_color="#1E8449", command=self._confirm)
        btn_next.pack(side="right", padx=(10, 40))

    # ---------------- 辅助构建方法 ----------------
    def _add_file_selector(self, parent, label_text, string_var, file_types):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=6, padx=15)
        ctk.CTkLabel(row, text=label_text, width=110, anchor="e", font=("微软雅黑", 12)).pack(side="left", padx=(0, 10))
        entry = ctk.CTkEntry(row, textvariable=string_var, state="readonly") 
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn = ctk.CTkButton(row, text="浏览", width=60, command=lambda: self._browse_file(string_var, file_types))
        btn.pack(side="right")

    def _add_dir_selector(self, parent, label_text, string_var):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=6, padx=15)
        ctk.CTkLabel(row, text=label_text, width=110, anchor="e", font=("微软雅黑", 12)).pack(side="left", padx=(0, 10))
        entry = ctk.CTkEntry(row, textvariable=string_var, state="readonly")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn = ctk.CTkButton(row, text="选择", width=60, command=lambda: self._browse_dir(string_var))
        btn.pack(side="right")

    def _add_entry_row(self, parent, label_text, string_var):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=8, padx=15)
        ctk.CTkLabel(row, text=label_text, width=120, anchor="e", font=("微软雅黑", 12)).pack(side="left", padx=(0, 10))
        ctk.CTkEntry(row, textvariable=string_var, width=150).pack(side="left")

    def _add_slider_row(self, parent, label_text, var, min_val, max_val, is_int=False):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=8, padx=15)
        ctk.CTkLabel(row, text=label_text, width=120, anchor="e", font=("微软雅黑", 12)).pack(side="left", padx=(0, 10))
        val_label = ctk.CTkLabel(row, text=str(var.get()), width=40)
        val_label.pack(side="right", padx=(10, 0))

        def slider_callback(value):
            final_val = int(value) if is_int else round(value, 2)
            var.set(final_val)
            val_label.configure(text=str(final_val))

        slider = ctk.CTkSlider(row, from_=min_val, to=max_val, variable=var, command=slider_callback)
        slider.pack(side="left", fill="x", expand=True)

    # ---------------- 业务逻辑方法 ----------------
    def _browse_file(self, string_var, file_types):
        # 注意：这里需要确保弹出的系统文件框依然在顶层
        self.root.attributes('-topmost', False) 
        path = filedialog.askopenfilename(filetypes=file_types)
        self.root.attributes('-topmost', True)
        if path:
            string_var.set(path)

    def _browse_dir(self, string_var):
        self.root.attributes('-topmost', False)
        path = filedialog.askdirectory()
        self.root.attributes('-topmost', True)
        if path:
            string_var.set(path)

    def save_config(self):
        config = {
            "excel_path": self.var_excel_path.get(),
            "json_path": self.var_json_path.get(),
            "img_path": self.var_img_path.get(),
            "font_dir": self.var_font_dir.get(),
            "output_dir": self.var_output_dir.get(),
            "sheet_name": self.var_sheet_name.get(),
            "font_scale": self.var_font_scale.get(),
            "y_offset": self.var_y_offset.get(),
            "spacing": self.var_spacing.get()
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            print("配置已保存！") # 后续这里可以接你现成的 ModernInfoDialog
        except Exception as e:
            print(f"保存失败: {e}")

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                self.var_excel_path.set(config.get("excel_path", ""))
                self.var_json_path.set(config.get("json_path", ""))
                self.var_img_path.set(config.get("img_path", ""))
                self.var_font_dir.set(config.get("font_dir", ""))
                self.var_output_dir.set(config.get("output_dir", ""))
                self.var_sheet_name.set(config.get("sheet_name", "Sheet2"))
                self.var_font_scale.set(config.get("font_scale", 1.68))
                self.var_y_offset.set(config.get("y_offset", -1.5))
                self.var_spacing.set(config.get("spacing", -5))
            except Exception as e:
                pass

    def _confirm(self):
        """点击下一步时，收集所有参数并销毁当前弹窗"""
        # 前置校验：必须选择 JSON，因为下一步严重依赖它
        if not self.var_json_path.get() or not os.path.exists(self.var_json_path.get()):
            # 可以在这里调出你的 ModernInfoDialog 提示用户
            return

        # 把所有数据打包成字典
        self.config_data = {
            "excel_path": self.var_excel_path.get(),
            "json_path": self.var_json_path.get(),
            "img_path": self.var_img_path.get(),
            "font_dir": self.var_font_dir.get(),
            "output_dir": self.var_output_dir.get(),
            "sheet_name": self.var_sheet_name.get(),
            "font_scale": self.var_font_scale.get(),
            "y_offset": self.var_y_offset.get(),
            "spacing": self.var_spacing.get()
        }
        self.root.destroy()

    def show(self):
        """显示弹窗并阻塞，直到点击下一步被销毁"""
        self.root.grab_set()
        self.root.master.wait_window(self.root)
        return self.config_data # 返回收集到的数据字典