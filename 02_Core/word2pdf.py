"""
===============================================================================
脚本名称：Word 转 PDF 工具 / PDF 合并 (word2pdf.py)
作者: ZGQ
功能概述：
    本脚本用于自动化处理 Word 文档到 PDF 的转换，解决手动导出耗时且易错的问题。

    这个工具可以批量将Word文档转换为PDF格式，还可以合并多个PDF文件。
    它提供了友好的图形界面，让新手也能轻松完成文档转换任务。
    支持Microsoft Word和WPS Office，自动处理转换过程中的错误。

核心工作流：
    1. 弹出窗口：提供一个简洁的界面让用户选择多个 Word 文件和一个输出文件夹。
    2. 引擎挂载：通过 pywin32 调用系统 COM 接口，兼容 Microsoft Word 和 WPS Office 两大主流办公软件。
    3. 批量转换：遍历用户选择的 Word 文件列表，逐一打开并另存为 PDF 格式，输出到指定文件夹。
    4. 错误处理：针对单个文件的转换失败，记录日志并跳过继续处理剩余文件；针对引擎崩溃的情况，自动重启引擎并重试当前文件。
    5. 结果汇总：在界面日志区域展示转换结果统计，完成转换闭环。
    
前置依赖：
    - 运行前必须打开目标文档。
    - 同级目录下需存在 `file_utils_backup.py` 模块。
===============================================================================
"""
# ==================== 1. 依赖库导入阶段 ====================
import os  # 用于处理操作系统级别的文件路径解析、文件扩展名分离等操作
import tkinter as tk  # 导入 Tkinter 基础 GUI 组件库，并重命名为 tk
from tkinter import filedialog, messagebox, scrolledtext, ttk  # 从 tkinter 独立导入文件选择对话框、消息弹窗、滚动文本框和高级主题组件 (ttk 用于选项卡)
import win32com.client  # 导入 pywin32 的客户端模块，用于调用 Windows 系统的 COM 接口 (控制 Word/WPS)
import threading  # 导入多线程模块，用于将耗时操作与 GUI 界面分离，防止界面假死
import pythoncom  # 导入 Python 的 COM 基础组件，用于在子线程中初始化和释放 COM 环境
import time  # 导入时间模块，用于在引擎崩溃重启时提供缓冲等待时间
from pypdf import PdfWriter  # 从 pypdf 库导入 PdfWriter 类，专门用于在内存中执行 PDF 拼接和写入

# ==================== 2. 主程序类定义 ====================
class EngineeringDocTool:
    """
    工程文档处理工具类。

    提供Word转PDF和PDF合并功能的图形界面工具。
    """
    def __init__(self, root):
        """
        初始化工具类。

        参数:
            root: tkinter主窗口对象。
        """
        self.root = root  # 接收并绑定主窗口对象
        self.root.title("工程报告处理工具集 (Word/PDF)")  # 设置窗口左上角的标题文字
        self.root.geometry("750x650")  # 设置窗口的初始分辨率为 750 宽 x 650 高
        
        # 定义公共数据结构，用于在不同函数间传递状态
        self.word_files = []  # 初始化空列表，存储待转换的 Word 文件的绝对路径
        self.word_output_folder = ""  # 初始化空字符串，存储 Word 转换后 PDF 的输出目录路径
        self.pdf_files = []  # 初始化空列表，存储待合并的 PDF 文件的绝对路径
        
        self.setup_ui()  # 调用界面初始化方法，开始绘制所有 UI 元素

    # ==================== 3. 全局 UI 构建逻辑 ====================
    def setup_ui(self):
        """
        设置用户界面。

        创建选项卡界面，包含Word转PDF和PDF合并两个功能。
        """
        # 3.1 实例化选项卡容器
        self.notebook = ttk.Notebook(self.root)  # 创建 Notebook 组件，作为选项卡的父容器
        self.notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)  # 将容器放置在窗口中，设置外边距，并允许其填满剩余空间

        # 3.2 创建具体的选项卡页面
        self.tab_word2pdf = ttk.Frame(self.notebook)  # 创建第一个 Frame 框架，用于装载 Word 转 PDF 功能的 UI
        self.tab_pdfmerge = ttk.Frame(self.notebook)  # 创建第二个 Frame 框架，用于装载 PDF 合并功能的 UI

        # 将这两个 Frame 添加到 Notebook 容器中，并设置顶部显示的标签文本
        self.notebook.add(self.tab_word2pdf, text="功能一: Word 批量转 PDF")
        self.notebook.add(self.tab_pdfmerge, text="功能二: PDF 自定义合并")

        # 分别调用两个独立的方法，向这两个空白选项卡中填充具体的按钮和列表
        self.build_word2pdf_tab()
        self.build_pdfmerge_tab()

        # 3.3 创建底部的全局运行日志区域
        frame_log = tk.LabelFrame(self.root, text="系统运行日志")  # 创建带有边框和标题的 LabelFrame 容器
        frame_log.pack(pady=5, padx=10, fill=tk.X)  # 将日志容器放置在底部，并在水平方向拉伸填满 (fill=tk.X)
        # 创建滚动文本框组件，初始状态设为 'disabled' (不可编辑)，防止用户误输入
        self.text_log = scrolledtext.ScrolledText(frame_log, height=8, state='disabled')
        self.text_log.pack(pady=5, padx=5, fill=tk.BOTH)  # 将文本框放入日志容器中

    # 封装的日志打印方法，用于向 UI 文本框写入实时信息
    def log(self, message):
        self.text_log.config(state='normal')  # 临时解除文本框的只读锁定
        self.text_log.insert(tk.END, message + "\n")  # 在文本框末尾 (tk.END) 插入新的日志信息并换行
        self.text_log.see(tk.END)  # 强制滚动条自动滚动到最底部，保持最新日志可见
        self.text_log.config(state='disabled')  # 重新锁定文本框为只读状态
        self.root.update()  # 强制刷新 tkinter 窗口，确保主线程在忙碌时也能重绘界面

    # ==================== 4. 功能一：Word 转 PDF 界面与逻辑 ====================
    def build_word2pdf_tab(self):
        # 4.1 Word 文件列表区域构建
        frame_list = tk.Frame(self.tab_word2pdf)  # 创建存放列表的框架
        frame_list.pack(pady=10, padx=10, fill=tk.X)  # 放置框架
        
        tk.Label(frame_list, text="待转换 Word 文件列表:").pack(anchor=tk.W)  # 创建文本标签，左对齐 (anchor=tk.W)
        self.listbox_word = tk.Listbox(frame_list, height=6)  # 创建列表框，显示高度为 6 行
        self.listbox_word.pack(fill=tk.X, pady=5)  # 放置列表框
        
        btn_frame1 = tk.Frame(frame_list)  # 创建存放列表控制按钮的子框架
        btn_frame1.pack(fill=tk.X)
        # 创建“添加文件”按钮，点击时触发 self.add_word_files 方法
        tk.Button(btn_frame1, text="添加文件", command=self.add_word_files).pack(side=tk.LEFT, padx=5)
        # 创建“清空列表”按钮，点击时触发 self.clear_word_files 方法
        tk.Button(btn_frame1, text="清空列表", command=self.clear_word_files).pack(side=tk.LEFT, padx=5)

        # 4.2 PDF 输出路径设置区域构建
        frame_out = tk.Frame(self.tab_word2pdf)  # 创建输出设置框架
        frame_out.pack(pady=10, padx=10, fill=tk.X)
        
        tk.Label(frame_out, text="PDF 输出文件夹:").pack(anchor=tk.W)  # 创建文本标签
        self.entry_word_out = tk.Entry(frame_out, state='readonly')  # 创建单行输入框，设为只读，防止手动输入错误路径
        self.entry_word_out.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))  # 放置输入框，允许水平扩展占用多余空间
        # 创建“选择路径”按钮，点击时触发 self.set_word_output 方法
        tk.Button(frame_out, text="选择路径", command=self.set_word_output).pack(side=tk.RIGHT)

        # 4.3 执行启动按钮构建
        # 创建开始转换按钮，背景色设为浅蓝以作强调，点击触发 self.start_word_conversion
        self.btn_start_word = tk.Button(self.tab_word2pdf, text="开始转换", command=self.start_word_conversion, bg="lightblue")
        self.btn_start_word.pack(pady=20, ipadx=30, ipady=5)  # ipadx/ipady 用于增加按钮内边距，使其变大

    def add_word_files(self):
        # 呼出系统文件选择对话框，限制只能选择 .doc 和 .docx 后缀的文件
        files = filedialog.askopenfilenames(title="选择 Word 文档", filetypes=[("Word 文档", "*.doc;*.docx")])
        for f in files:  # 遍历用户选中的所有文件
            path = os.path.abspath(f)  # 将路径强制转换为 Windows 标准的绝对路径
            if path not in self.word_files:  # 防止重复添加同一个文件
                self.word_files.append(path)  # 将路径存入后台数据列表
                self.listbox_word.insert(tk.END, os.path.basename(path))  # 仅提取文件名 (不含长路径)，显示在 UI 列表框中

    def clear_word_files(self):
        self.word_files.clear()  # 清空后台存储的 Word 文件路径列表
        self.listbox_word.delete(0, tk.END)  # 从索引 0 到末尾，清空 UI 列表框的显示内容

    def set_word_output(self):
        # 呼出系统文件夹选择对话框
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:  # 如果用户选择了文件夹且未点击取消
            self.word_output_folder = folder  # 将路径存入全局变量
            self.entry_word_out.config(state='normal')  # 临时解除输入框的只读锁定
            self.entry_word_out.delete(0, tk.END)  # 清空输入框旧内容
            self.entry_word_out.insert(0, folder)  # 插入用户选择的新路径
            self.entry_word_out.config(state='readonly')  # 恢复输入框为只读状态

    def mount_engine(self):
        # 引擎挂载函数：负责调用系统的 COM 接口启动办公软件
        try:
            # 尝试通过 COM 接口派生 Microsoft Word 应用程序实例
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False  # 强制后台运行，不显示 Word 软件的图形界面
            word.DisplayAlerts = 0  # 屏蔽宏警告、字体缺失等各种弹窗阻断
            word.Options.UpdateLinksAtOpen = False  # 核心设置：全局禁止打开文档时更新外部网络链接，防止假死
            return word, "Microsoft Word"  # 返回成功挂载的实例对象和引擎名称
        except:
            try:
                # 若挂载微软 Word 失败，尝试通过 COM 接口派生 WPS 实例
                word = win32com.client.DispatchEx("KWPS.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                return word, "WPS Office"
            except Exception as e:
                # 若两套引擎均失败，抛出致命异常
                raise Exception(f"底层接口异常: {str(e)}")

    def kill_engine(self, word):
        # 引擎释放函数：负责在任务完成或崩溃时清理内存占用
        if word:
            try:
                word.Quit()  # 调用 COM 接口发送退出指令关闭 Word 进程
            except:
                pass  # 若进程已经死亡无法响应 Quit，则忽略报错
            finally:
                del word  # 删除 Python 层面的变量引用，触发垃圾回收

    def start_word_conversion(self):
        # 任务启动预检逻辑
        if not self.word_files or not self.word_output_folder:
            messagebox.showwarning("提示", "请选择 Word 文件及输出路径")  # 校验前置条件是否满足
            return
        
        self.btn_start_word.config(state='disabled')  # 任务开始，禁用转换按钮，防止用户重复点击导致多重线程堆叠
        self.log("--- 启动 Word 转 PDF 任务 ---")
        # 核心：实例化一个新线程执行 process_word_conversion，设置 daemon=True 确保主窗口关闭时后台线程一并销毁
        threading.Thread(target=self.process_word_conversion, daemon=True).start()

    def process_word_conversion(self):
        # 在子线程中初始化 COM 运行环境（多线程调用 pywin32 的强制要求）
        pythoncom.CoInitialize()
        word = None
        success_cnt = 0  # 初始化成功计数器
        
        try:
            word, engine_name = self.mount_engine()  # 获取引擎实例
            self.log(f"挂载引擎: {engine_name}")

            for input_path in self.word_files:  # 遍历待处理队列
                file_name = os.path.basename(input_path)  # 提取当前处理的文件名
                if file_name.startswith('~$'): continue  # 跳过 Office 临时生成的隐藏锁定文件，避免访问越权报错
                
                # 拼接输出的绝对路径：输出文件夹 + 原文件名(去后缀) + .pdf
                output_path = os.path.join(self.word_output_folder, os.path.splitext(file_name)[0] + ".pdf")
                self.log(f"处理中: {file_name}")
                
                retry, max_retries, success = 0, 1, False  # 设定单个文件的重试控制变量
                while retry <= max_retries and not success:
                    try:
                        # 指挥引擎打开文档，启用只读模式(ReadOnly=1)进一步提升稳定性和速度
                        doc = word.Documents.Open(input_path, ReadOnly=1, Visible=False)
                        # 指挥引擎执行另存为，FileFormat=17 代表官方定义的 wdFormatPDF 枚举值
                        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
                        doc.Close(0)  # 指挥引擎关闭当前文档，不保存更改 (0)
                        success = True
                        success_cnt += 1
                    except Exception as e:
                        err_msg = str(e)
                        # 检测是否发生了 RPC 丢失错误（意味着底层 Word.exe 进程崩溃了）
                        if "-214702317" in err_msg or "RPC" in err_msg:
                            self.log("引擎通信中断，执行强退重启...")
                            self.kill_engine(word)  # 彻底清理死锁的旧实例
                            time.sleep(1)  # 挂起当前线程 1 秒，等待系统层释放文件锁定
                            word, _ = self.mount_engine()  # 重新生成一个干净的 Word 进程继续战斗
                            retry += 1  # 增加重试次数
                        else:
                            self.log(f"跳过文件 [{file_name}]: {err_msg}")  # 常规错误（如文件加密）直接跳过
                            break  # 退出 while 循环，放弃处理该文件
            self.log(f"--- 转换结束 | 成功: {success_cnt}/{len(self.word_files)} ---")
        except Exception as e:
            self.log(f"系统级错误: {str(e)}")
        finally:
            self.kill_engine(word)  # 全局任务结束，释放引擎
            pythoncom.CoUninitialize()  # 释放当前线程的 COM 环境，避免内存泄漏
            # 利用 lambda 表达式，跨线程指挥主线程安全地将“开始转换”按钮恢复为可点击状态
            self.root.after(0, lambda: self.btn_start_word.config(state='normal'))

    # ==================== 5. 功能二：PDF 合并界面与逻辑 ====================
    def build_pdfmerge_tab(self):
        # 5.1 待合并列表区域
        frame_list = tk.Frame(self.tab_pdfmerge)
        frame_list.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        tk.Label(frame_list, text="待合并 PDF 列表 (自上而下合并):").pack(anchor=tk.W)
        
        list_inner = tk.Frame(frame_list)  # 创建包含列表和滚动条的内部容器
        list_inner.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scroll_pdf = tk.Scrollbar(list_inner)  # 创建垂直滚动条组件
        self.scroll_pdf.pack(side=tk.RIGHT, fill=tk.Y)  # 放置在右侧，垂直方向填满 (fill=tk.Y)
        
        # 创建单选模式 (selectmode=tk.SINGLE) 的列表框，并将其上下滚动动作与滚动条绑定
        self.listbox_pdf = tk.Listbox(list_inner, yscrollcommand=self.scroll_pdf.set, selectmode=tk.SINGLE)
        self.listbox_pdf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_pdf.config(command=self.listbox_pdf.yview)  # 反向绑定，拖动滚动条时更新列表框视图

        # 5.2 列表排序控制按钮区域
        frame_sort = tk.Frame(frame_list)  # 创建用于存放排序按钮的侧边容器
        frame_sort.pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        
        # 绑定上下移及移除功能的独立按钮
        tk.Button(frame_sort, text="上移", command=self.move_pdf_up, width=8).pack(pady=5)
        tk.Button(frame_sort, text="下移", command=self.move_pdf_down, width=8).pack(pady=5)
        tk.Button(frame_sort, text="移除", command=self.remove_pdf_file, width=8).pack(pady=20)

        # 5.3 底层文件添加按钮
        frame_ops = tk.Frame(self.tab_pdfmerge)
        frame_ops.pack(padx=10, fill=tk.X)
        tk.Button(frame_ops, text="添加 PDF", command=self.add_pdf_files).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_ops, text="清空列表", command=self.clear_pdf_files).pack(side=tk.LEFT, padx=5)

        # 开始合并按钮
        self.btn_start_merge = tk.Button(self.tab_pdfmerge, text="开始合并", command=self.start_pdf_merge, bg="lightblue")
        self.btn_start_merge.pack(pady=20, ipadx=30, ipady=5)

    def add_pdf_files(self):
        # 限制文件选择器只能选中 PDF 格式
        files = filedialog.askopenfilenames(title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")])
        for f in files:
            path = os.path.abspath(f)
            if path not in self.pdf_files:
                self.pdf_files.append(path)  # 将绝对路径写入后台列表
                self.listbox_pdf.insert(tk.END, os.path.basename(path))  # 将文件名写入前台 UI 列表

    def remove_pdf_file(self):
        idx = self.listbox_pdf.curselection()  # 获取当前 UI 列表框被选中的项的索引值（返回的是一个元组）
        if idx:
            self.pdf_files.pop(idx[0])  # 从后台数据列表中剔除该索引对应的数据
            self.listbox_pdf.delete(idx)  # 从前台 UI 中删除该条目

    def move_pdf_up(self):
        idx = self.listbox_pdf.curselection()
        # 判断是否有选中项，且选中项不是第一项（索引 0 无法再上移）
        if idx and idx[0] > 0:
            i = idx[0]
            # 核心算法：利用 Python 特性交换后台列表中当前元素与上一个元素的位置
            self.pdf_files[i], self.pdf_files[i-1] = self.pdf_files[i-1], self.pdf_files[i]
            # 同步更新 UI：获取当前文本，删除原位置条目，在上一位置重新插入，并保持高亮选中状态
            text = self.listbox_pdf.get(i)
            self.listbox_pdf.delete(i)
            self.listbox_pdf.insert(i-1, text)
            self.listbox_pdf.select_set(i-1)

    def move_pdf_down(self):
        idx = self.listbox_pdf.curselection()
        # 判断是否有选中项，且选中项不是最后一项
        if idx and idx[0] < len(self.pdf_files) - 1:
            i = idx[0]
            # 交换后台列表中当前元素与下一个元素的位置
            self.pdf_files[i], self.pdf_files[i+1] = self.pdf_files[i+1], self.pdf_files[i]
            text = self.listbox_pdf.get(i)
            self.listbox_pdf.delete(i)
            self.listbox_pdf.insert(i+1, text)
            self.listbox_pdf.select_set(i+1)

    def clear_pdf_files(self):
        self.pdf_files.clear()
        self.listbox_pdf.delete(0, tk.END)

    def start_pdf_merge(self):
        if len(self.pdf_files) < 2:  # 合并操作至少需要存在两个文件
            messagebox.showwarning("提示", "至少需要添加 2 个 PDF 文件进行合并")
            return
        
        # 呼出“另存为”对话框，要求用户确认最终合并文件的保存路径和名称
        output_path = filedialog.asksaveasfilename(
            title="保存合并后的文件",
            initialfile="检测报告_合并输出.pdf",  # 提供默认文件名
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        
        if output_path:  # 用户确立了保存路径
            self.btn_start_merge.config(state='disabled')  # 禁用按钮防止重复点击
            self.log("--- 启动 PDF 合并任务 ---")
            # 启动子线程执行合并写入 IO 任务，将 output_path 作为参数传递给 target 函数
            threading.Thread(target=self.process_pdf_merge, args=(output_path,), daemon=True).start()

    def process_pdf_merge(self, output_path):
        merger = PdfWriter()  # 实例化 pypdf 提供的合并器对象，负责在内存中管理 PDF 字节树
        try:
            for pdf in self.pdf_files:  # 按当前的顺序遍历后台列表
                self.log(f"装载文件: {os.path.basename(pdf)}")
                merger.append(pdf)  # 调用 API 将目标 PDF 文件的所有页面与书签数据追加至合并器序列中
            
            # 使用上下文管理器以二进制写模式 (wb) 打开目标输出路径的空文件
            with open(output_path, "wb") as f:
                merger.write(f)  # 将内存中组装完毕的 PDF 字节流集中写入磁盘
            
            self.log(f"--- 合并成功 | 已保存至: {output_path} ---")
        except Exception as e:
            self.log(f"合并失败: {str(e)}")
        finally:
            merger.close()  # 无论成功失败，显式关闭合并器实例释放内存流
            # 使用 after 方法指挥主线程将按钮解禁
            self.root.after(0, lambda: self.btn_start_merge.config(state='normal'))

# ==================== 6. 脚本入口点 ====================
# 当文件作为独立脚本被执行时，执行以下块内的代码
if __name__ == "__main__":
    root = tk.Tk()  # 初始化底层的 Tkinter 解释器，创建主循环窗口实例
    app = EngineeringDocTool(root)  # 实例化自建的业务类，将窗口对象挂载进去
    root.mainloop()  # 进入无限循环监听事件（鼠标点击、按键等），维持窗口展示直至程序被关闭