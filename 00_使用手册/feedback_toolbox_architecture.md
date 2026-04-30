---
name: 工具箱式代码架构偏好
description: 用户在 02_Core 工程自动化项目里偏好的"工具箱"式代码组织 —— 业务/UI 分离、common 共享、dataclass、动词命名、main 菜单
type: feedback
originSessionId: c9988218-6cb4-4f0c-81d3-1dca18da21b6
---
写新工具或重构现有工具时，遵循以下 5 条原则。这是用户在 2026-04-30 重构 sort_photos.py + renumber_photos.py 时确认并要求作为长期标准的结构。

**Why:** 用户在做"工程检测报告自动化"项目，工具会越积越多（已有 9+ 个独立脚本），如果继续各写各的、复制粘贴公共逻辑，下次要复用时会重写、改正则要改一堆地方。

**How to apply:** 任何新加到 `02_Core/` 的工具，或重构旧工具时，都按这 5 条来：

1. **业务逻辑 vs UI 流程，物理分离**
   - 所有"读文件 / 弹窗 / print 业务消息"只能在 `if __name__ == "__main__":` 包裹的 `_main()` 函数里
   - 业务函数（`run_sort` / `rebuild_word_by_order` 这种）只接收参数、返回数据、抛异常
   - 验收标准：`import xxx` 不能触发任何对话框或文件读取

2. **公共能力沉淀到 `02_Core/common/`**
   - `common/io_helpers.py` —— pick_excel_file / read_sheet_names / unblock_file / kill_winword_processes / enable_line_buffered_stdout / ensure_extension
   - `common/word_helpers.py` —— scan_photo_pairs / build_caption_renumber_mapping / replace_in_caption_rows / make_caption_substitutor
   - `common/excel_helpers.py` —— get_excel_sort_order / find_column_index / replace_in_excel_column
   - `common/ui_helpers.py` —— field_sheet_select / field_text / field_word_file / field_dir 这种 form schema 工厂
   - `common/patterns.py` —— FIG_PATTERN 等全项目复用的正则
   - `common/types.py` —— PhotoPair 等 dataclass
   - 第二个工具用到同一段代码 = 该上提到 common/

3. **跨模块数据用 dataclass，不用裸 dict**
   - `PhotoPair(num, img_row_idx, txt_row_idx, img_col_idx, txt_col_idx)` 取代 `{"num": ..., "img_row_idx": ...}`
   - IDE 能跳转/补全/重命名

4. **配置参数化，不要全局 Config 类属性**
   - 旧版 `Config.EXCEL_PATH = ...` 是全局态污染，函数没法被另一个 main 用不同参数调用
   - 改成 `def run_xxx(excel_path, sheet_name, col_name, ...)`，参数显式

5. **一个工具 = 一个动词，顶层 main.py 做菜单**
   - 文件名是动词：sort_photos / renumber_photos / check_consistency / export_images
   - 每个工具暴露一个 `run_xxx(...)` 入口函数 + `_main()` 走 UI 流程
   - 新增工具 = 在 `main.py` 的 `modules` 列表加一行 `(显示名, 文件名.py)`，main.py 用 subprocess.Popen 启动，不用改其他东西

**新工具的模板（30 行骨架）：**
```python
import os, sys
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from common.io_helpers import enable_line_buffered_stdout, pick_excel_file, read_sheet_names
from common.ui_helpers import field_sheet_select, field_text, field_word_file, field_dir
from ui_components import ModernDynamicFormDialog

def run_xxx(excel_path, sheet_name, col_name, word_path, ...):
    """核心业务逻辑 —— 纯参数，可被其他工具 import 调用"""
    ...

def _request_params(excel_path, sheet_names):
    schema = [field_sheet_select(sheet_names), field_text(...), ...]
    return ModernDynamicFormDialog(title="...", form_schema=schema).show()

def _main():
    enable_line_buffered_stdout()
    excel_path = pick_excel_file()
    if not excel_path: return
    sheets = read_sheet_names(excel_path)
    if not sheets: return
    params = _request_params(excel_path, sheets)
    if not params: return
    run_xxx(excel_path, params["sheet_name"], ...)

if __name__ == "__main__":
    _main()
```

**额外约定：**
- 每个工具脚本文件顶部都要有 `sys.path.insert(0, _THIS_DIR)` 的 shim，保证 main.py 用 subprocess 启动时能解析 `from common.xxx import ...`
- print 业务日志用 emoji 分级：🚀 启动 / 📊 数据统计 / 🔍 扫描 / 📂 打开 / 💾 保存 / ✅ 成功 / ⚠️ 警告 / ❌ 错误 / 🎉 全部完成 ——  与现有 sort_photos.py 风格一致
- 新工具上线必须能从 `main.py` 主面板的按钮点击启动
