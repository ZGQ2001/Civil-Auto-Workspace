"""
===============================================================================
脚本名称：PDF坐标拾取工具 (pdf_coordinate_picker.py)
作者: ZGQ
功能概述：
    本脚本用于处理 PDF 文件中的坐标拾取和标记，提供一个可视化界面让用户点击 PDF 上的任意位置，输入字段名称，并将坐标以 JSON 格式导出，方便后续自动化排版等应用。

    这个工具可以让你在PDF页面上点击位置，记录坐标，并保存为配置文件。
    适合用来标记PDF表单中的字段位置，为自动化填写做准备。新手可以通过简单的点击和命名来创建坐标配置。

核心工作流：
    1. 环境检测：抓取当前处于激活状态的 PDF 文档。
    2. 可视化界面：使用 Tkinter 创建一个窗口显示 PDF 的第一页，支持缩放和平移。
    3. 坐标拾取：用户左键点击 PDF 上的任意位置，弹出输入框让用户为该坐标命名，并记录坐标的 PDF 内部位置。
    4. 数据管理：用户可以随时保存已拾取的坐标到 JSON 文件，程序会提示未保存的修改，防止数据丢失。
    5. 结果汇总：展示已拾取坐标的列表，完成坐标拾取闭环。

前置依赖：
    - 运行前必须选择一个 PDF 文件。
    - 需要安装 fitz（PyMuPDF）和 Pillow 库。
===============================================================================
"""
import fitz  # PDF处理库
import tkinter as tk  # GUI库
from tkinter import filedialog, simpledialog, messagebox  # tkinter子模块
from PIL import Image, ImageTk  # 图像处理库
import json  # JSON文件处理
import os  # 文件路径操作

class PDFCoordinatePicker:
    """
    PDF坐标拾取器类。

    提供一个图形界面来显示PDF并让用户点击拾取坐标。
    """

    def __init__(self, pdf_path, page_num=0):
        """
        初始化坐标拾取器。

        参数:
            pdf_path (str): PDF文件的路径。
            page_num (int): 要显示的页码，默认为0（第一页）。
        """
        self.pdf_path = pdf_path
        # 打开PDF并获取第一页
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)  # 渲染页面为图像
        self.img_full = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_width, self.pdf_height = page.rect.width, page.rect.height  # PDF页面尺寸
        doc.close()

        # 缩放和平移相关变量
        self.zoom_scale = 1.0
        self.offset_x, self.offset_y = 0, 0
        self.last_mouse_x, self.last_mouse_y = 0, 0
        
        # 坐标数据存储
        self.saved_coords = {}  # 保存的坐标字典
        self.has_unsaved_changes = False  # 是否有未保存的修改

        # 创建主窗口
        self.root = tk.Tk()
        self.root.title(f"坐标拾取导出工具 - {os.path.basename(pdf_path)}")
        self.root.geometry("1000x900")

        # 拦截窗口关闭事件，防止意外丢失数据
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.canvas = tk.Canvas(self.root, bg="gray", cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Button-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.do_pan)
        self.root.bind("<Control-s>", lambda e: self.save_to_json())
        self.root.bind("<Configure>", self.on_window_resize)
        
        # 绑定 Esc 到统一的退出检查函数
        self.root.bind("<Escape>", lambda e: self.on_closing())

        self.status = tk.Label(self.root, text="左键点击拾取 | Ctrl+滚轮缩放 | 右键拖动 | Ctrl+S 保存 | Esc/关闭 退出检查", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        self.tk_img = None
        self.first_show = True

    def on_click(self, event):
        """
        处理鼠标左键点击事件。

        计算点击位置对应的PDF坐标，并让用户命名该坐标。
        """
        # 获取窗口和图像尺寸
        win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        img_w, img_h = self.img_full.size
        dw, dh = img_w * self.zoom_scale, img_h * self.zoom_scale
        img_left = win_w // 2 + self.offset_x - (dw / 2)
        img_top = win_h // 2 + self.offset_y - (dh / 2)

        # 检查点击是否在图像范围内
        if img_left <= event.x <= img_left + dw and img_top <= event.y <= img_top + dh:
            # 计算PDF坐标
            pdf_x = ((event.x - img_left) / dw) * self.pdf_width
            pdf_y = ((event.y - img_top) / dh) * self.pdf_height
            # 弹出输入框让用户命名坐标
            field_name = simpledialog.askstring("坐标命名", f"拾取坐标({pdf_x:.1f}, {pdf_y:.1f})\n请输入字段名称：")
            
            if field_name:
                # 保存坐标到字典
                self.saved_coords[field_name] = {"x": round(pdf_x, 1), "y": round(pdf_y, 1), "size": 12}
                self.has_unsaved_changes = True  # 标记有未保存的修改
                self.render_image()  # 重新渲染图像，显示标记

    def save_to_json(self):
        """保存逻辑，增加返回值以供退出检查调用"""
        if not self.saved_coords:
            messagebox.showwarning("提示", "当前没有记录任何坐标！")
            return False
            
        save_path = filedialog.asksaveasfilename(
            title="保存坐标配置文件",
            initialdir=os.path.dirname(self.pdf_path),
            initialfile="coords_config.json",
            filetypes=[("JSON files", "*.json")]
        )
        
        if save_path:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(self.saved_coords, f, ensure_ascii=False, indent=4)
            self.has_unsaved_changes = False  # 保存后重置修改标记
            messagebox.showinfo("成功", f"坐标已导出至：\n{save_path}")
            return True
        return False

    def on_closing(self):
        """防呆退出检查"""
        if self.has_unsaved_changes:
            response = messagebox.askyesnocancel("退出确认", "检测到有未保存的坐标，是否保存后再退出？\n\n【是】保存并退出\n【否】放弃保存直接退出\n【取消】返回继续工作")
            if response is True:  # 用户选“是”
                if self.save_to_json(): # 执行保存
                    self.root.destroy()
            elif response is False: # 用户选“否”
                self.root.destroy()
            else: # 用户选“取消”或直接关掉提示框
                pass
        else:
            # 如果没有修改过数据，或者已经保存过了，直接退出
            self.root.destroy()

    # (render_image, draw_markers, on_window_resize, on_mousewheel, start_pan, do_pan 等函数保持原样)
    def render_image(self):
        """
        渲染PDF图像到画布上。

        根据当前的缩放比例和偏移量显示图像。
        """
        win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if win_w <= 1: return
        img_w, img_h = self.img_full.size
        dw, dh = int(img_w * self.zoom_scale), int(img_h * self.zoom_scale)
        # 根据缩放比例选择重采样方法
        resample = Image.Resampling.LANCZOS if self.zoom_scale > 0.5 else Image.Resampling.NEAREST
        self.tk_img = ImageTk.PhotoImage(self.img_full.resize((dw, dh), resample))
        self.canvas.delete("all")  # 清空画布
        # 在画布中心显示图像
        self.canvas.create_image(win_w//2 + self.offset_x, win_h//2 + self.offset_y, image=self.tk_img)
        self.draw_markers()  # 绘制坐标标记

    def draw_markers(self):
        """
        在图像上绘制坐标标记。

        为每个保存的坐标绘制红色圆点和蓝色标签。
        """
        win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        img_w, img_h = self.img_full.size
        dw, dh = img_w * self.zoom_scale, img_h * self.zoom_scale
        img_left = win_w // 2 + self.offset_x - (dw / 2)
        img_top = win_h // 2 + self.offset_y - (dh / 2)
        for name, pos in self.saved_coords.items():
            # 计算屏幕坐标
            screen_x = (pos['x'] / self.pdf_width) * dw + img_left
            screen_y = (pos['y'] / self.pdf_height) * dh + img_top
            # 绘制红色圆点
            self.canvas.create_oval(screen_x-3, screen_y-3, screen_x+3, screen_y+3, fill="red")
            # 绘制蓝色标签
            self.canvas.create_text(screen_x+5, screen_y, text=name, anchor=tk.W, fill="blue", font=("Arial", 10, "bold"))
            self.canvas.create_text(screen_x+5, screen_y, text=name, anchor=tk.W, fill="blue", font=("Arial", 10, "bold"))

    def on_window_resize(self, event):
        if self.first_show:
            win_w, win_h = event.width, event.height
            img_w, img_h = self.img_full.size
            self.zoom_scale = min(win_w / img_w, win_h / img_h) * 0.95
            self.first_show = False
        self.render_image()

    def on_mousewheel(self, event):
        if event.state & 0x0004:
            zoom_step = 1.1 if event.delta > 0 else 0.9
            win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
            mouse_rel_x = event.x - (win_w // 2 + self.offset_x)
            mouse_rel_y = event.y - (win_h // 2 + self.offset_y)
            old_scale = self.zoom_scale
            self.zoom_scale = max(0.1, min(self.zoom_scale * zoom_step, 5.0))
            real_step = self.zoom_scale / old_scale
            self.offset_x -= (mouse_rel_x * real_step - mouse_rel_x)
            self.offset_y -= (mouse_rel_y * real_step - mouse_rel_y)
            self.render_image()

    def start_pan(self, event):
        self.last_mouse_x, self.last_mouse_y = event.x, event.y

    def do_pan(self, event):
        self.offset_x += event.x - self.last_mouse_x
        self.offset_y += event.y - self.last_mouse_y
        self.last_mouse_x, self.last_mouse_y = event.x, event.y
        self.render_image()

    def run(self):
        """
        启动应用程序的主循环。
        """
        self.root.mainloop()

if __name__ == "__main__":
    """
    主程序入口。

    让用户选择PDF文件，然后启动坐标拾取器。
    """
    # 创建临时窗口用于文件选择对话框
    temp_root = tk.Tk()
    temp_root.withdraw()
    # 设置初始目录
    initial_dir = r"D:\01_VC CODE\Civil-Auto-Workspace\01_Input"
    # 弹出文件选择对话框
    pdf_file = filedialog.askopenfilename(
        title="选择检测记录表 PDF 模板",
        initialdir=initial_dir if os.path.exists(initial_dir) else os.getcwd(),
        filetypes=[("PDF files", "*.pdf")]
    )
    temp_root.destroy()  # 销毁临时窗口
    if pdf_file:
        # 如果选择了文件，创建坐标拾取器并运行
        app = PDFCoordinatePicker(pdf_file)
        app.run()