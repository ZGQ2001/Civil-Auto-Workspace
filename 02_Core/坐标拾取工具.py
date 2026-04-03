import fitz
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PIL import Image, ImageTk
import json
import os

class PDFCoordinatePicker:
    def __init__(self, pdf_path, page_num=0):
        self.pdf_path = pdf_path
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        self.img_full = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_width, self.pdf_height = page.rect.width, page.rect.height
        doc.close()

        self.zoom_scale = 1.0
        self.offset_x, self.offset_y = 0, 0
        self.last_mouse_x, self.last_mouse_y = 0, 0
        
        self.saved_coords = {}
        self.has_unsaved_changes = False  # 防呆开关：标记是否有未保存的修改

        self.root = tk.Tk()
        self.root.title(f"坐标拾取导出工具 - {os.path.basename(pdf_path)}")
        self.root.geometry("1000x900")

        # 拦截窗口关闭事件（右上角 X）
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
        # ... (此处省略与上一版本相同的坐标换算逻辑)
        # 如果用户成功输入了字段名并确认：
        # self.has_unsaved_changes = True 
        # self.render_image()
        
        # 为了简洁，以下仅展示变动部分：
        win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        img_w, img_h = self.img_full.size
        dw, dh = img_w * self.zoom_scale, img_h * self.zoom_scale
        img_left = win_w // 2 + self.offset_x - (dw / 2)
        img_top = win_h // 2 + self.offset_y - (dh / 2)

        if img_left <= event.x <= img_left + dw and img_top <= event.y <= img_top + dh:
            pdf_x = ((event.x - img_left) / dw) * self.pdf_width
            pdf_y = ((event.y - img_top) / dh) * self.pdf_height
            field_name = simpledialog.askstring("坐标命名", f"拾取坐标({pdf_x:.1f}, {pdf_y:.1f})\n请输入字段名称：")
            
            if field_name:
                self.saved_coords[field_name] = {"x": round(pdf_x, 1), "y": round(pdf_y, 1), "size": 12}
                self.has_unsaved_changes = True  # 触发修改标记
                self.render_image()

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
        win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if win_w <= 1: return
        img_w, img_h = self.img_full.size
        dw, dh = int(img_w * self.zoom_scale), int(img_h * self.zoom_scale)
        resample = Image.Resampling.LANCZOS if self.zoom_scale > 0.5 else Image.Resampling.NEAREST
        self.tk_img = ImageTk.PhotoImage(self.img_full.resize((dw, dh), resample))
        self.canvas.delete("all")
        self.canvas.create_image(win_w//2 + self.offset_x, win_h//2 + self.offset_y, image=self.tk_img)
        self.draw_markers()

    def draw_markers(self):
        win_w, win_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        img_w, img_h = self.img_full.size
        dw, dh = img_w * self.zoom_scale, img_h * self.zoom_scale
        img_left = win_w // 2 + self.offset_x - (dw / 2)
        img_top = win_h // 2 + self.offset_y - (dh / 2)
        for name, pos in self.saved_coords.items():
            screen_x = (pos['x'] / self.pdf_width) * dw + img_left
            screen_y = (pos['y'] / self.pdf_height) * dh + img_top
            self.canvas.create_oval(screen_x-3, screen_y-3, screen_x+3, screen_y+3, fill="red")
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
        self.root.mainloop()

if __name__ == "__main__":
    temp_root = tk.Tk()
    temp_root.withdraw()
    initial_dir = r"D:\01_VC CODE\Civil-Auto-Workspace\01_Input"
    pdf_file = filedialog.askopenfilename(
        title="选择检测记录表 PDF 模板",
        initialdir=initial_dir if os.path.exists(initial_dir) else os.getcwd(),
        filetypes=[("PDF files", "*.pdf")]
    )
    temp_root.destroy()
    if pdf_file:
        app = PDFCoordinatePicker(pdf_file)
        app.run()