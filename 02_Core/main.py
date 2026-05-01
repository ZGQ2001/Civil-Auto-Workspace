"""
工程自动化主控制台 (main.py) - V2 双栏导航版

布局：
    ┌─ 顶部：标题 ─────────────────────────────────────────────┐
    ├─ 左：分类侧边栏  ├─ 右：工具说明卡片                      ┤
    │                 │   + 启动 / 停止按钮                   │
    │                 ├─ 实时日志面板（捕获子进程 stdout）     │
    └─────────────────┴────────────────────────────────────┘

核心改进：
    1. 工具分类，可折叠分组，避免一长溜竖排
    2. 子进程的 stdout / stderr 实时流回到日志面板，不再"黑箱"
    3. 同时只跑一个工具，左侧任何时候都可以切换查看其他工具说明
    4. 工具自身的对话框 / 进度条仍然由各脚本自己管，dashboard 只做调度 + 日志
"""
import os
import queue
import subprocess
import sys
import threading
from typing import Dict, List, Optional, Tuple

import customtkinter as ctk

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# 工具菜单：分类 + (显示名, 脚本文件名, 简介)
# ============================================================
ToolEntry = Tuple[str, str, str]
CATEGORIES: List[Tuple[str, List[ToolEntry]]] = [
    ("📝 报告排版", [
        ("正文排版引擎", "body_format.py",
         "扫描 Word 当前活动文档，按 04_Config/report_style_config.json 自动套用字体、间距、缩进、大纲级别。"),
        ("表格排版引擎", "table_format.py",
         "对 Word 文档里所有表格统一字号、行高、表名样式，并把空单元格高亮。"),
        ("括号半全角纠偏", "bracket_format.py",
         "通过 Word 通配符引擎全局规范括号：技术参数转半角、国标代号锁全角、第N回半角。"),
        ("交叉引用修复", "fix_cross_ref.py",
         "为所有 REF 域追加 \\* MERGEFORMAT 开关，避免后续编辑丢字号字体。"),
    ]),
    ("📷 照片附录", [
        ("照片流水线（排序+重编号）", "pipeline_sort_renumber.py",
         "一键完成：按 Excel 缺陷清单顺序重排 Word 表格里的图片+题注 → 题注重编号 → 同步改 Excel 引用。"),
    ]),
    ("📊 数据 / 绘图", [
        ("批量绘图（Excel→PNG）", "plot_curves.py",
         "从 Excel 某个 Sheet 按行读取，套用 04_Config/curve_templates.json 模板批量出 PNG，运行前会做列名预检。"),
    ]),
    ("📄 文档转换 / 工具", [
        ("Word ↔ PDF 转换", "word2pdf.py",
         "Word → PDF 与 PDF → Word 双向转换。"),
        ("PNG 坐标拾取器", "coord_picker.py",
         "在 PNG 底图上拖拽框选，输出 100% 像素坐标 JSON，给手写模拟工具用。"),
        ("手写模拟生成器", "auto_filler.py",
         "读 Excel 数据 + PNG 底图 + JSON 坐标 → 仿生手写体填表，导出 PDF。"),
    ]),
    ("⚙ 配置编辑器", [
        ("报告样式配置（字体/间距）", "config_editor.py",
         "可视化编辑 04_Config/report_style_config.json，给「正文排版引擎」/「表格排版引擎」用。"),
        ("曲线模板（绘图）", "curve_template_editor.py",
         "可视化编辑 04_Config/curve_templates.json，给「批量绘图」用；支持挂载 Excel 让列名变下拉选。"),
    ]),
]


# ============================================================
# 主控制台
# ============================================================
class MainDashboard:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title("工程自动化主控制台 V2")
        self.root.geometry("1180x760")
        self.root.minsize(960, 600)

        # 当前选中的工具与正在跑的子进程
        self.current_tool: Optional[ToolEntry] = None
        self.proc: Optional[subprocess.Popen] = None
        self.log_queue: queue.Queue = queue.Queue()
        self._reader_thread: Optional[threading.Thread] = None

        self._build_layout()
        self._select_tool(CATEGORIES[0][1][0])  # 默认选第一个工具
        self._poll_log_queue()

    # ------------------------------------------------------------
    # 布局
    # ------------------------------------------------------------
    def _build_layout(self) -> None:
        # 顶栏
        header = ctk.CTkFrame(self.root, height=64, corner_radius=0, fg_color=("#0078d4", "#1f3a5f"))
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header, text="工程自动化主控制台",
            font=("微软雅黑", 20, "bold"), text_color="white",
        ).pack(side="left", padx=24, pady=14)
        ctk.CTkLabel(
            header, text="Automation Dashboard V2",
            font=("Consolas", 11), text_color=("#cce4ff", "#aac5e0"),
        ).pack(side="left", pady=18)

        # 主体：左右分栏
        body = ctk.CTkFrame(self.root, fg_color="transparent")
        body.pack(fill="both", expand=True)

        # === 左侧：侧边栏 ===
        sidebar = ctk.CTkFrame(body, width=280, corner_radius=0, fg_color=("gray92", "gray18"))
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        ctk.CTkLabel(
            sidebar, text="工具分类", font=("微软雅黑", 12, "bold"), text_color="gray55", anchor="w",
        ).pack(fill="x", padx=20, pady=(16, 6))

        self.sidebar_scroll = ctk.CTkScrollableFrame(sidebar, fg_color="transparent")
        self.sidebar_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self._tool_buttons: Dict[str, ctk.CTkButton] = {}  # script_name -> btn
        self._build_sidebar()

        # === 右侧：内容区 ===
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True, padx=14, pady=14)

        # 工具说明卡片
        info_card = ctk.CTkFrame(right, corner_radius=12)
        info_card.pack(fill="x")

        info_inner = ctk.CTkFrame(info_card, fg_color="transparent")
        info_inner.pack(fill="x", padx=22, pady=18)

        self.info_title = ctk.CTkLabel(
            info_inner, text="—", font=("微软雅黑", 18, "bold"), anchor="w",
        )
        self.info_title.pack(fill="x", anchor="w")

        self.info_script = ctk.CTkLabel(
            info_inner, text="", font=("Consolas", 10), text_color="gray55", anchor="w",
        )
        self.info_script.pack(fill="x", anchor="w", pady=(4, 8))

        self.info_desc = ctk.CTkLabel(
            info_inner, text="", font=("微软雅黑", 12), text_color=("gray25", "gray85"),
            anchor="w", justify="left", wraplength=760,
        )
        self.info_desc.pack(fill="x", anchor="w")

        # 启动 / 停止按钮区
        btn_row = ctk.CTkFrame(info_inner, fg_color="transparent")
        btn_row.pack(fill="x", pady=(14, 0))

        self.btn_run = ctk.CTkButton(
            btn_row, text="▶ 启动该工具", height=38, width=160,
            font=("微软雅黑", 13, "bold"),
            command=self._launch_current,
        )
        self.btn_run.pack(side="left", padx=(0, 8))

        self.btn_stop = ctk.CTkButton(
            btn_row, text="■ 停止运行", height=38, width=120,
            font=("微软雅黑", 13), fg_color="#aa3333", hover_color="#cc4444",
            command=self._stop_current, state="disabled",
        )
        self.btn_stop.pack(side="left", padx=4)

        self.btn_clear = ctk.CTkButton(
            btn_row, text="🧹 清空日志", height=38, width=110,
            font=("微软雅黑", 12), fg_color="gray45", hover_color="gray55",
            command=self._clear_log,
        )
        self.btn_clear.pack(side="left", padx=4)

        self.run_status = ctk.CTkLabel(
            btn_row, text="● 空闲", font=("微软雅黑", 12, "bold"), text_color="gray60",
        )
        self.run_status.pack(side="right", padx=8)

        # 日志面板
        log_card = ctk.CTkFrame(right, corner_radius=12)
        log_card.pack(fill="both", expand=True, pady=(14, 0))

        log_header = ctk.CTkFrame(log_card, fg_color="transparent", height=36)
        log_header.pack(fill="x", padx=18, pady=(12, 0))
        log_header.pack_propagate(False)
        ctk.CTkLabel(
            log_header, text="📡 实时日志", font=("微软雅黑", 13, "bold"), anchor="w",
        ).pack(side="left")

        self.log_text = ctk.CTkTextbox(
            log_card, font=("Consolas", 11), wrap="word",
            fg_color=("gray97", "gray12"), text_color=("gray10", "gray88"),
        )
        self.log_text.pack(fill="both", expand=True, padx=14, pady=10)
        self.log_text.insert("end", "就绪。点击左侧工具，再按「▶ 启动」开始执行。\n")
        self.log_text.configure(state="disabled")

    def _build_sidebar(self) -> None:
        for cat_name, tools in CATEGORIES:
            # 分组标题（始终展开，简单起见不做折叠交互）
            cat_label = ctk.CTkLabel(
                self.sidebar_scroll, text=cat_name,
                font=("微软雅黑", 12, "bold"), anchor="w",
                text_color=("gray35", "gray70"),
            )
            cat_label.pack(fill="x", padx=8, pady=(10, 4))

            for entry in tools:
                display_name, script, _desc = entry
                btn = ctk.CTkButton(
                    self.sidebar_scroll, text=display_name, anchor="w",
                    font=("微软雅黑", 12), height=34, corner_radius=6,
                    fg_color="transparent", text_color=("gray10", "gray90"),
                    hover_color=("gray85", "gray28"),
                    command=lambda e=entry: self._select_tool(e),
                )
                btn.pack(fill="x", padx=4, pady=2)
                self._tool_buttons[script] = btn

    # ------------------------------------------------------------
    # 选中 / 高亮
    # ------------------------------------------------------------
    def _select_tool(self, entry: ToolEntry) -> None:
        self.current_tool = entry
        display_name, script, desc = entry

        # 高亮当前
        for s, btn in self._tool_buttons.items():
            if s == script:
                btn.configure(fg_color=("#cce4ff", "#1f3a5f"), text_color=("#003a73", "#cce4ff"))
            else:
                btn.configure(fg_color="transparent", text_color=("gray10", "gray90"))

        # 更新右侧说明卡片
        self.info_title.configure(text=display_name)
        self.info_script.configure(text=f"脚本：{script}")
        self.info_desc.configure(text=desc)

    # ------------------------------------------------------------
    # 启动 / 停止子进程
    # ------------------------------------------------------------
    def _launch_current(self) -> None:
        if self.current_tool is None:
            return
        if self.proc is not None and self.proc.poll() is None:
            self._append_log("⚠️ 已有工具在运行，请先停止或等待结束。\n")
            return

        _, script, _ = self.current_tool
        script_path = os.path.join(_THIS_DIR, script)
        if not os.path.exists(script_path):
            self._append_log(f"❌ 缺失模块：{script_path}\n")
            return

        self._clear_log()
        self._append_log(f"🚀 启动 {script} ...\n\n")

        # Windows 下隐藏黑色 cmd 窗口；强制 -u 让子进程立即刷出 stdout
        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        try:
            # 不传 bufsize=1：默认 8KB 块缓冲在二进制流上更稳；子进程 -u 已经保证 print 即时 flush。
            self.proc = subprocess.Popen(
                [sys.executable, "-u", script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                cwd=_THIS_DIR,
                env=env,
                creationflags=creationflags,
            )
        except Exception as e:
            self._append_log(f"❌ 启动失败：{e}\n")
            self.proc = None
            return

        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.run_status.configure(text="● 运行中", text_color="#1aaa55")

        self._reader_thread = threading.Thread(
            target=self._reader_worker, args=(self.proc,), daemon=True,
        )
        self._reader_thread.start()

    def _reader_worker(self, proc: subprocess.Popen) -> None:
        """后台线程：阻塞读 stdout 字节流，逐行扔进 log_queue。"""
        assert proc.stdout is not None
        try:
            while True:
                chunk = proc.stdout.readline()
                if not chunk:
                    break
                # 子进程 stdout 是 bytes（bufsize=1 + 没传 text=True）
                if isinstance(chunk, bytes):
                    line = chunk.decode("utf-8", errors="replace")
                else:
                    line = chunk
                self.log_queue.put(line)
        finally:
            proc.stdout.close()
            rc = proc.wait()
            self.log_queue.put(f"\n[进程结束 returncode={rc}]\n")
            self.log_queue.put(("__DONE__", rc))

    def _stop_current(self) -> None:
        if self.proc is None or self.proc.poll() is not None:
            return
        self._append_log("\n⛔ 用户请求停止...\n")
        try:
            self.proc.terminate()
        except Exception as e:
            self._append_log(f"   terminate 失败: {e}\n")

    # ------------------------------------------------------------
    # 日志面板
    # ------------------------------------------------------------
    def _poll_log_queue(self) -> None:
        """主线程定时轮询 queue，把后台线程读到的输出贴到 textbox。"""
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    self._on_proc_done(item[1])
                else:
                    self._append_log(item)
        except queue.Empty:
            pass
        self.root.after(80, self._poll_log_queue)

    def _on_proc_done(self, returncode: int) -> None:
        self.proc = None
        self.btn_run.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        if returncode == 0:
            self.run_status.configure(text="● 已完成", text_color="#1aaa55")
        else:
            self.run_status.configure(text=f"● 异常退出 (rc={returncode})", text_color="#aa3333")

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


def _main() -> None:
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    MainDashboard(root)
    root.mainloop()


if __name__ == "__main__":
    _main()
