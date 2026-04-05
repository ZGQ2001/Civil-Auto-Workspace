"""
===============================================================================
脚本名称：工程报告处理工具集 (word2pdf_pro.py)
作者: ZGQ / Gemini
功能概述：
    1. Word 批量转 PDF (支持 Word/WPS)
    2. PDF 自定义排序合并
    3. 文档转高清底图 (300/600 DPI PNG)
===============================================================================
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import win32com.client
import threading
import pythoncom
import time
from pypdf import PdfWriter
import fitz  # PyMuPDF

class EngineeringDocTool:
    def __init__(self, root):
        self.root = root
        self.root.title("工程报告处理工具集 Pro (Word/PDF/PNG)")
        self.root.geometry("800x750")
        self.root.configure(bg="#f5f5f5") # 设置全局背景色
        
        # --- 数据存储初始化 ---
        self.word_files = []      # 功能一：Word文件路径
        self.word_out_dir = ""    # 功能一：PDF输出目录
        
        self.merge_files = []     # 功能二：PDF待合并路径
        self.merge_out_path = ""  # 功能二：合并后保存位置
        
        self.png_files = []       # 功能三：待转图文档路径
        self.png_out_dir = ""     # 功能三：底图输出目录
        
        # 定义 UI 风格常量 (专业配色方案)
        self.CLR_PRIMARY = "#0078d4"    # 微软蓝：主执行按钮
        self.CLR_SECONDARY = "#e1e1e1"  # 浅灰：辅助按钮
        self.CLR_TEXT_BG = "#ffffff"    # 纯白：输入框/列表框
        self.CLR_LOG_BG = "#2b2b2b"     # 深灰：日志区域
        self.CLR_LOG_FG = "#a9b7c6"     # 亮灰：日志文字

        self.setup_ui()

    def setup_ui(self):
        """全局界面构建：选项卡 + 日志区"""
        # 1. 选项卡容器
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=15, fill=tk.BOTH, expand=True)

        # 2. 实例化并添加三个功能页面
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab1, text="  Word 批量转 PDF  ")
        self.notebook.add(self.tab2, text="  PDF 排序合并  ")
        self.notebook.add(self.tab3, text="  文档转高清 PNG  ")

        # 依次构建各页面的详细组件
        self._build_word_tab()
        self._build_merge_tab()
        self._build_png_tab()

        # 3. 底部系统日志区域 (工业感配色)
        log_frame = tk.LabelFrame(self.root, text=" 任务执行监控 ", bg="#f5f5f5")
        log_frame.pack(pady=10, padx=15, fill=tk.X)
        self.log_area = scrolledtext.ScrolledText(
            log_frame, height=10, bg=self.CLR_LOG_BG, fg=self.CLR_LOG_FG, 
            font=("Consolas", 9), state='disabled'
        )
        self.log_area.pack(pady=5, padx=5, fill=tk.BOTH)

    def log(self, msg):
        """实时日志刷新逻辑"""
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update()

    # ==================== 界面通用模板 (保持统一感) ====================
    def _create_standard_layout(self, parent, list_title, list_height=8):
        """封装标准的三段式布局，确保每个页面长得一模一样"""
        # A. 列表区
        frame_list = tk.LabelFrame(parent, text=f" {list_title} ", bg="white")
        frame_list.pack(pady=10, padx=10, fill=tk.X)
        
        lbox = tk.Listbox(frame_list, height=list_height, bg=self.CLR_TEXT_BG, borderwidth=0, highlightthickness=1)
        lbox.pack(fill=tk.X, pady=5, padx=5)
        
        btn_box = tk.Frame(frame_list, bg="white")
        btn_box.pack(fill=tk.X, pady=5, padx=5)
        
        # B. 输出路径设置区 (Entry + Button)
        frame_out = tk.LabelFrame(parent, text=" 输出与配置 ", bg="white")
        frame_out.pack(pady=10, padx=10, fill=tk.X)
        
        entry_path = tk.Entry(frame_out, state='readonly', readonlybackground="#f8f9fa")
        entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=10)
        
        return lbox, btn_box, entry_path, frame_out

    # ==================== 功能一：Word 转 PDF ====================
    def _build_word_tab(self):
        lbox, b_box, ent, f_out = self._create_standard_layout(self.tab1, "待处理 Word 文件列表")
        self.lb_word, self.ent_word = lbox, ent
        
        tk.Button(b_box, text="添加文件", command=self.add_word_files).pack(side=tk.LEFT, padx=2)
        tk.Button(b_box, text="清空列表", command=self.clear_word).pack(side=tk.LEFT, padx=5)
        tk.Button(f_out, text="选择输出路径", command=self.set_word_out).pack(side=tk.RIGHT, padx=5)
        
        self.btn_run_word = tk.Button(
            self.tab1, text="🚀 开始批量转换", bg=self.CLR_PRIMARY, fg="white", 
            font=("微软雅黑", 10, "bold"), command=self.run_word_task
        )
        self.btn_run_word.pack(pady=15, ipadx=50, ipady=8)

    def add_word_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Word", "*.doc;*.docx")])
        for f in files:
            p = os.path.abspath(f)
            if p not in self.word_files:
                self.word_files.append(p)
                self.lb_word.insert(tk.END, f"📄 {os.path.basename(p)}")

    def clear_word(self):
        self.word_files.clear()
        self.lb_word.delete(0, tk.END)

    def set_word_out(self):
        d = filedialog.askdirectory()
        if d:
            self.word_out_dir = d
            self._update_entry(self.ent_word, d)

    # ==================== 功能二：PDF 排序合并 ====================
    def _build_merge_tab(self):
        # 增加排序控制按钮
        lbox, b_box, ent, f_out = self._create_standard_layout(self.tab2, "PDF 合并列表 (自上而下合并)")
        self.lb_merge, self.ent_merge = lbox, ent
        
        tk.Button(b_box, text="添加文件", command=self.add_merge_files).pack(side=tk.LEFT, padx=2)
        tk.Button(b_box, text="上移", command=lambda: self.move_item(-1)).pack(side=tk.LEFT, padx=2)
        tk.Button(b_box, text="下移", command=lambda: self.move_item(1)).pack(side=tk.LEFT, padx=2)
        tk.Button(b_box, text="清空", command=self.clear_merge).pack(side=tk.LEFT, padx=2)
        
        tk.Button(f_out, text="指定保存文件", command=self.set_merge_out).pack(side=tk.RIGHT, padx=5)

        self.btn_run_merge = tk.Button(
            self.tab2, text="🔗 开始合并 PDF", bg=self.CLR_PRIMARY, fg="white", 
            font=("微软雅黑", 10, "bold"), command=self.run_merge_task
        )
        self.btn_run_merge.pack(pady=15, ipadx=50, ipady=8)

    def add_merge_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for f in files:
            p = os.path.abspath(f)
            if p not in self.merge_files:
                self.merge_files.append(p)
                self.lb_merge.insert(tk.END, f"📑 {os.path.basename(p)}")

    def move_item(self, step):
        idx = self.lb_merge.curselection()
        if not idx: return
        i = idx[0]
        ni = i + step
        if 0 <= ni < len(self.merge_files):
            self.merge_files[i], self.merge_files[ni] = self.merge_files[ni], self.merge_files[i]
            txt = self.lb_merge.get(i)
            self.lb_merge.delete(i)
            self.lb_merge.insert(ni, txt)
            self.lb_merge.select_set(ni)

    def clear_merge(self):
        self.merge_files.clear()
        self.lb_merge.delete(0, tk.END)

    def set_merge_out(self):
        f = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="合并结果.pdf")
        if f:
            self.merge_out_path = f
            self._update_entry(self.ent_merge, f)

    # ==================== 功能三：文档转高清 PNG ====================
    def _build_png_tab(self):
        lbox, b_box, ent, f_out = self._create_standard_layout(self.tab3, "待转底图文件列表")
        self.lb_png, self.ent_png = lbox, ent
        
        tk.Button(b_box, text="添加文件", command=self.add_png_files).pack(side=tk.LEFT, padx=2)
        tk.Button(b_box, text="清空列表", command=self.clear_png).pack(side=tk.LEFT, padx=2)
        
        # 配置行：DPI选择
        cfg_row = tk.Frame(f_out, bg="white")
        cfg_row.pack(side=tk.LEFT, padx=10)
        tk.Label(cfg_row, text="DPI:", bg="white").pack(side=tk.LEFT)
        self.cb_dpi = ttk.Combobox(cfg_row, values=[150, 300, 600], width=5)
        self.cb_dpi.set(300)
        self.cb_dpi.pack(side=tk.LEFT, padx=5)

        tk.Button(f_out, text="选择保存目录", command=self.set_png_out).pack(side=tk.RIGHT, padx=5)

        self.btn_run_png = tk.Button(
            self.tab3, text="🖼️ 批量生成高清底图", bg=self.CLR_PRIMARY, fg="white", 
            font=("微软雅黑", 10, "bold"), command=self.run_png_task
        )
        self.btn_run_png.pack(pady=15, ipadx=50, ipady=8)

    def add_png_files(self):
        files = filedialog.askopenfilenames(filetypes=[("文档", "*.doc;*.docx;*.pdf")])
        for f in files:
            p = os.path.abspath(f)
            if p not in self.png_files:
                self.png_files.append(p)
                self.lb_png.insert(tk.END, f"🖼️ {os.path.basename(p)}")

    def clear_png(self):
        self.png_files.clear()
        self.lb_png.delete(0, tk.END)

    def set_png_out(self):
        d = filedialog.askdirectory()
        if d:
            self.png_out_dir = d
            self._update_entry(self.ent_png, d)

    def _update_entry(self, entry, text):
        """统一更新只读输入框内容"""
        entry.config(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, text)
        entry.config(state='readonly')

    # ==================== 后台引擎逻辑 (核心工作流) ====================
    def _mount_word_engine(self):
        """挂载 COM 引擎：支持 Microsoft Word 或 WPS"""
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            return word, "Microsoft Word"
        except:
            try:
                word = win32com.client.DispatchEx("KWPS.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                return word, "WPS Office"
            except: raise Exception("未检测到 Word/WPS，请检查软件安装环境")

    def run_word_task(self):
        if not self.word_files or not self.word_out_dir:
            messagebox.showwarning("警告", "请先添加文件并设置输出路径！")
            return
        self.btn_run_word.config(state='disabled')
        threading.Thread(target=self._proc_word, daemon=True).start()

    def _proc_word(self):
        """Word 转 PDF 核心子线程"""
        pythoncom.CoInitialize()
        word_app = None
        try:
            word_app, name = self._mount_word_engine()
            self.log(f"已启动引擎: {name}")
            for p in self.word_files:
                self.log(f"正在转换: {os.path.basename(p)} ...")
                out = os.path.join(self.word_out_dir, os.path.splitext(os.path.basename(p))[0] + ".pdf")
                doc = word_app.Documents.Open(p, ReadOnly=1)
                doc.SaveAs(os.path.abspath(out), FileFormat=17) # 17 = wdFormatPDF
                doc.Close(0)
            self.log("✅ 所有 Word 转换任务已完成")
        except Exception as e: self.log(f"❌ 错误: {str(e)}")
        finally:
            if word_app: word_app.Quit()
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.btn_run_word.config(state='normal'))

    def run_merge_task(self):
        if len(self.merge_files) < 2 or not self.merge_out_path:
            messagebox.showwarning("提示", "合并至少需要2个文件且需设置保存位置！")
            return
        self.btn_run_merge.config(state='disabled')
        threading.Thread(target=self._proc_merge, daemon=True).start()

    def _proc_merge(self):
        """PDF 合并核心子线程"""
        writer = PdfWriter()
        try:
            self.log("--- 启动 PDF 合并序列 ---")
            for p in self.merge_files:
                self.log(f"追加: {os.path.basename(p)}")
                writer.append(p)
            with open(self.merge_out_path, "wb") as f:
                writer.write(f)
            self.log(f"✅ 合并成功！已保存至: {self.merge_out_path}")
        except Exception as e: self.log(f"❌ 合并失败: {str(e)}")
        finally:
            writer.close()
            self.root.after(0, lambda: self.btn_run_merge.config(state='normal'))

    def run_png_task(self):
        if not self.png_files or not self.png_out_dir:
            messagebox.showwarning("提示", "请检查待转列表和输出文件夹！")
            return
        self.btn_run_png.config(state='disabled')
        threading.Thread(target=self._proc_png, daemon=True).start()

    def _proc_png(self):
        """文档转图片核心子线程"""
        pythoncom.CoInitialize()
        word_app = None
        dpi = int(self.cb_dpi.get())
        mat = fitz.Matrix(dpi/72, dpi/72) # PDF/Word 默认 72DPI，此处计算比例
        try:
            for p in self.png_files:
                ext = os.path.splitext(p)[1].lower()
                pdf_path = p
                is_tmp = False
                
                # 如果是 Word，先走后台静默转 PDF
                if ext in ['.doc', '.docx']:
                    if not word_app: word_app, _ = self._mount_word_engine()
                    pdf_path = os.path.join(self.png_out_dir, "temp_render.pdf")
                    doc = word_app.Documents.Open(p, ReadOnly=1)
                    doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
                    doc.Close(0)
                    is_tmp = True
                
                # 开始图片像素渲染
                self.log(f"正在渲染高清图像: {os.path.basename(p)}")
                doc_pdf = fitz.open(pdf_path)
                for i in range(len(doc_pdf)):
                    pix = doc_pdf[i].get_pixmap(matrix=mat, alpha=False)
                    out_name = f"{os.path.splitext(os.path.basename(p))[0]}_P{i+1}.png"
                    pix.save(os.path.join(self.png_out_dir, out_name))
                doc_pdf.close()
                if is_tmp and os.path.exists(pdf_path): os.remove(pdf_path)
                self.log(f"完成: {os.path.basename(p)}")
            self.log("✅ 所有底图生成任务已完成")
        except Exception as e: self.log(f"❌ 渲染异常: {str(e)}")
        finally:
            if word_app: word_app.Quit()
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.btn_run_png.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    # 强制启用系统原生 DPI 缩放（防止在 4K 屏幕上模糊）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    
    app = EngineeringDocTool(root)
    root.mainloop()