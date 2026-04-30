import os
import re
import copy
import pandas as pd
from tkinter import filedialog
from docx import Document

# 【架构解耦】：直接从你的通用 UI 库中引入动态表单组件
from ui_components import ModernDynamicFormDialog

# ==========================================
# 模块 1：配置中心
# ==========================================
class Config:
    EXCEL_PATH = '缺陷清单.xlsx'
    EXCEL_SHEET = None              # None = 第一个 sheet
    EXCEL_COL_NAME = '照片'
    WORD_PATH = '待排序_附录1.docx'
    OUTPUT_DIR = ''                 # 空 = 与 Word 同目录
    OUTPUT_NAME = '已排序_附录1.docx'
    OUTPUT_PATH = '已排序_附录1.docx'   # 由 OUTPUT_DIR + OUTPUT_NAME 拼出
    MATCH_PATTERN = r'图\s*(\d+)'

# ==========================================
# 模块 2：解析器（提取排序规则）
# ==========================================
def get_excel_sort_order(path, col_name, sheet_name=None):
    try:
        df = pd.read_excel(path, sheet_name=sheet_name) if sheet_name else pd.read_excel(path)
        if col_name not in df.columns:
            print(f"❌ Sheet [{sheet_name or '默认'}] 中找不到列：{col_name}")
            return []
        raw_list = df[col_name].dropna().astype(str).tolist()
        order = []
        for item in raw_list:
            match = re.search(Config.MATCH_PATTERN, item)
            if match:
                order.append(int(match.group(1)))
        return order
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return []

# ==========================================
# 模块 3：文档重构器（核心排序逻辑）
# ==========================================
class WordTableProcessor:
    def __init__(self, doc_path):
        self.doc = Document(doc_path)
        self.pairs = {}
        self.unmatched = []

    def extract_pairs(self, excel_order):
        if not self.doc.tables:
            raise ValueError("❌ Word 文档中没有找到任何表格！")

        table = self.doc.tables[0]

        for i in range(0, len(table.rows), 2):
            if i + 1 >= len(table.rows):
                break

            img_row = table.rows[i]
            txt_row = table.rows[i+1]

            for j in range(len(img_row.cells)):
                img_cell = img_row.cells[j]
                txt_cell = txt_row.cells[j]

                text = txt_cell.text.strip()
                if not text:
                    continue

                match = re.search(Config.MATCH_PATTERN, text)
                if match:
                    num = int(match.group(1))
                    pair_data = [copy.deepcopy(img_cell._tc), copy.deepcopy(txt_cell._tc)]

                    if num in excel_order:
                        self.pairs[num] = pair_data
                    else:
                        self.unmatched.append(pair_data)

    def rebuild(self, excel_order):
        new_doc = Document()
        new_table = new_doc.add_table(rows=0, cols=2)
        new_table.style = 'Table Grid'

        final_list = []
        for num in excel_order:
            if num in self.pairs:
                final_list.append(self.pairs[num])
        final_list.extend(self.unmatched)

        for k in range(0, len(final_list), 2):
            r1 = new_table.add_row().cells
            r2 = new_table.add_row().cells

            r1[0]._element.getparent().replace(r1[0]._element, final_list[k][0])
            r2[0]._element.getparent().replace(r2[0]._element, final_list[k][1])

            if k + 1 < len(final_list):
                r1[1]._element.getparent().replace(r1[1]._element, final_list[k+1][0])
                r2[1]._element.getparent().replace(r2[1]._element, final_list[k+1][1])

        new_doc.save(Config.OUTPUT_PATH)
        print(f"✅ 处理完成！已生成新文件: {Config.OUTPUT_PATH}")

# ==========================================
# 模块 4：执行入口（纯粹的业务调用）
# ==========================================
def pick_excel_first():
    """先单独弹一个原生文件框选 Excel，目的是为了能预读 Sheet 列表"""
    path = filedialog.askopenfilename(
        title="第一步：选择缺陷清单 Excel",
        filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")]
    )
    return path

def read_sheet_names(excel_path):
    try:
        return pd.ExcelFile(excel_path).sheet_names
    except Exception as e:
        print(f"❌ 读取 Sheet 列表失败: {e}")
        return []

def request_sort_params(excel_path, sheet_names):
    """弹出主参数窗口，让用户挑选 Sheet、列名、Word、输出位置"""
    default_dir = os.path.dirname(excel_path) or os.getcwd()
    schema = [
        {"key": "sheet_name", "label": "Excel 工作表:", "type": "select",
         "options": sheet_names, "default": sheet_names[0] if sheet_names else ""},
        {"key": "excel_col", "label": "照片列表头:", "type": "text",
         "default": Config.EXCEL_COL_NAME},
        {"key": "word_path", "label": "待排序 Word:", "type": "file",
         "file_types": [("Word 文件", "*.docx"), ("所有文件", "*.*")],
         "default": Config.WORD_PATH},
        {"key": "output_dir", "label": "输出目录:", "type": "dir",
         "default": default_dir},
        {"key": "output_name", "label": "输出文件名:", "type": "text",
         "default": Config.OUTPUT_NAME},
    ]
    dialog = ModernDynamicFormDialog(title="照片排序 - 参数配置", form_schema=schema, width=620)
    return dialog.show()


if __name__ == "__main__":
    # 阶段一：先选 Excel 文件，才能读出可选的 Sheet 列表
    excel_path = pick_excel_first()
    if not excel_path:
        print("⚠️ 已取消：未选择 Excel 文件。")
        raise SystemExit

    sheets = read_sheet_names(excel_path)
    if not sheets:
        print("⚠️ 终止：该 Excel 没有可读取的工作表。")
        raise SystemExit

    # 阶段二：动态表单收集其余参数（Sheet 下拉框 + 输出位置）
    params = request_sort_params(excel_path, sheets)

    if not params or not params.get("word_path"):
        print("⚠️ 已取消：未选择 Word 文件或直接关闭了窗口。")
    else:
        Config.EXCEL_PATH = excel_path
        Config.EXCEL_SHEET = params.get("sheet_name") or sheets[0]
        Config.EXCEL_COL_NAME = params.get("excel_col") or Config.EXCEL_COL_NAME
        Config.WORD_PATH = params["word_path"]
        Config.OUTPUT_DIR = params.get("output_dir") or os.path.dirname(Config.WORD_PATH)
        Config.OUTPUT_NAME = params.get("output_name") or Config.OUTPUT_NAME

        # 文件名安全兜底：如果用户没写后缀，自动补 .docx
        if not Config.OUTPUT_NAME.lower().endswith(".docx"):
            Config.OUTPUT_NAME += ".docx"

        os.makedirs(Config.OUTPUT_DIR, exist_ok=True)
        Config.OUTPUT_PATH = os.path.join(Config.OUTPUT_DIR, Config.OUTPUT_NAME)

        print(f"🚀 正在读取数据... [Sheet: {Config.EXCEL_SHEET}]")
        order = get_excel_sort_order(Config.EXCEL_PATH, Config.EXCEL_COL_NAME, Config.EXCEL_SHEET)

        if not order:
            print("⚠️ 终止：Excel 排序数据为空。")
        else:
            print(f"📊 从 Excel 获取到 {len(order)} 个排序指令，开始重构 Word...")
            processor = WordTableProcessor(Config.WORD_PATH)
            processor.extract_pairs(order)
            processor.rebuild(order)
