import fitz
import tkinter as tk
from PIL import Image, ImageTk

class PDFCoordinatePicker:
    def __init__(self, pdf_path, page_num=0):
        # 1. 加载 PDF 基础数据
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        # 母版使用 150 DPI，保证放大后依然清晰
        pix = page.get_pixmap(dpi=150)
        self.img_full = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_width, self.pdf_height = page.rect.width, page.rect.height
        doc.close()

        # 2. 初始化状态变量
        self.zoom_scale = 1.0       # 缩放倍率
        self.offset_x = 0           # 图像在画布上的偏移
        self.offset_y = 0
        self.last_mouse_x = 0       # 用于右键拖动平移
        self.last_mouse_y = 0

        # 3. 初始化窗口
        self.root = tk.Tk()
        self.root.title("坐标拾取 [Ctrl+滚轮缩放 | 右键拖动平移 | 左键拾取 | Esc退出]")
        self.root.geometry("900x1000")

        # 4. 初始化画布
        self.canvas = tk.Canvas(self.root, bg="gray", cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 5. 绑定事件
        self.canvas.bind("<Button-1>", self.on_click)           # 左键点击拾取
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)     # 滚轮缩放
        self.canvas.bind("<Button-3>", self.start_pan)          # 右键按下（平移开始）
        self.canvas.bind("<B3-Motion>", self.do_pan)            # 右键拖动
        self.root.bind("<Configure>", self.on_window_resize)    # 窗口大小改变
        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.tk_img = None
        self.first_show = True # 标记是否是第一次加载，用于自动适配大小

    def render_image(self):
        """重新计算并绘制当前视角的图像"""
        win_w = self.canvas.winfo_width()
        win_h = self.canvas.winfo_height()
        if win_w <= 1: return # 窗口未准备好

        # 计算当前缩放下的尺寸
        img_w, img_h = self.img_full.size
        display_w = int(img_w * self.zoom_scale)
        display_h = int(img_h * self.zoom_scale)

        # 重新缩放图像（放大使用自适应采样）
        resample = Image.Resampling.LANCZOS if self.zoom_scale > 0.5 else Image.Resampling.NEAREST
        resized_img = self.img_full.resize((display_w, display_h), resample)
        self.tk_img = ImageTk.PhotoImage(resized_img)

        # 清除画布并绘制
        self.canvas.delete("all")
        # 将图像根据偏移量绘制在画布上
        # 初始默认居中显示
        self.canvas.create_image(win_w//2 + self.offset_x, win_h//2 + self.offset_y, image=self.tk_img)

    def on_window_resize(self, event):
        """窗口大小改变时，首次加载执行自适应"""
        if self.first_show:
            win_w, win_h = event.width, event.height
            img_w, img_h = self.img_full.size
            # 自动计算一个初始缩放，使 PDF 适应窗口高度
            self.zoom_scale = min(win_w / img_w, win_h / img_h) * 0.95
            self.first_show = False
        self.render_image()

    def on_mousewheel(self, event):
        """Ctrl + 滚轮实现以鼠标为中心缩放"""
        if event.state & 0x0004:  # 检查 Ctrl 是否按下
            # 缩放灵敏度
            zoom_step = 1.1 if event.delta > 0 else 0.9
            
            # 计算鼠标相对于图像中心的相对位置（用于锁定缩放点）
            win_w = self.canvas.winfo_width()
            win_h = self.canvas.winfo_height()
            
            # 鼠标在图像上的相对坐标
            mouse_rel_x = event.x - (win_w // 2 + self.offset_x)
            mouse_rel_y = event.y - (win_h // 2 + self.offset_y)

            # 更新缩放比
            old_scale = self.zoom_scale
            self.zoom_scale *= zoom_step
            # 限制缩放范围
            self.zoom_scale = max(0.1, min(self.zoom_scale, 5.0))
            
            # 重新调整偏移量，实现以鼠标指针为中心缩放
            real_step = self.zoom_scale / old_scale
            self.offset_x -= (mouse_rel_x * real_step - mouse_rel_x)
            self.offset_y -= (mouse_rel_y * real_step - mouse_rel_y)

            self.render_image()

    def start_pan(self, event):
        """右键按下，记录坐标"""
        self.last_mouse_x = event.x
        self.last_mouse_y = event.y

    def do_pan(self, event):
        """右键拖动平移图像"""
        dx = event.x - self.last_mouse_x
        dy = event.y - self.last_mouse_y
        self.offset_x += dx
        self.offset_y += dy
        self.last_mouse_x = event.x
        self.last_mouse_y = event.y
        self.render_image()

    def on_click(self, event):
        """左键点击映射坐标"""
        win_w = self.canvas.winfo_width()
        win_h = self.canvas.winfo_height()

        # 计算图像左上角在画布上的实际位置
        img_w, img_h = self.img_full.size
        cur_display_w = img_w * self.zoom_scale
        cur_display_h = img_h * self.zoom_scale
        
        img_left = win_w // 2 + self.offset_x - (cur_display_w / 2)
        img_top = win_h // 2 + self.offset_y - (cur_display_h / 2)

        # 检查点击是否在图像区域
        if img_left <= event.x <= img_left + cur_display_w and \
           img_top <= event.y <= img_top + cur_display_h:
            
            rel_x = (event.x - img_left) / cur_display_w
            rel_y = (event.y - img_top) / cur_display_h
            
            pdf_x = rel_x * self.pdf_width
            pdf_y = rel_y * self.pdf_height
            
            print(f"坐标 -> x: {pdf_x:.1f}, y: {pdf_y:.1f}")
            # 画一个持久的小红点
            self.canvas.create_oval(event.x-3, event.y-3, event.x+3, event.y+3, fill="red", outline="white")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # 请确保路径正确
    pdf_file = r"D:\01_VC CODE\Civil-Auto-Workspace\01_Input\49_构件截面尺寸检测记录表-钢结构.pdf"
    app = PDFCoordinatePicker(pdf_file)
    app.run()