import os
import re
import copy
import pandas as pd
from docx import Document

# ==========================================
# 模块 1：配置中心（运行前只需修改这里）
# ==========================================
class Config:
    # 你的 Excel 文件名（确保和代码在同一个文件夹）
    EXCEL_PATH = '缺陷清单.xlsx'        
    # 表格中包含“图2”那一列的表头名称
    EXCEL_COL_NAME = '照片'            
    
    # 你单独提出来的、只包含照片表格的 Word 文件名
    WORD_PATH = '待排序_附录1.docx'        
    # 程序处理完后输出的新文件名
    OUTPUT_PATH = '已排序_附录1.docx'  
    
    # 匹配规则：匹配“图 2”或“图2”并提取数字
    MATCH_PATTERN = r'图\s*(\d+)'      

# ==========================================
# 模块 2：解析器（提取排序规则）
# ==========================================
def get_excel_sort_order(path, col_name):
    """读取 Excel，返回目标顺序列表，例如 [2, 4, 13, 18...]"""
    try:
        df = pd.read_excel(path)
        # 提取目标列，过滤掉空行
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
        self.pairs = {}      # 存储能匹配上的：{图号: [图片XML, 文字XML]}
        self.unmatched = []  # 存储Excel里没提到的多余图片

    def extract_pairs(self, excel_order):
        """将 Word 表格按“一行图+一行字”拆解为独立内存块"""
        if not self.doc.tables:
            raise ValueError("❌ Word 文档中没有找到任何表格！请确保放入了正确的表格。")
            
        table = self.doc.tables[0]
        
        # 步长为 2 遍历表格（i 为图片行，i+1 为图名行）
        for i in range(0, len(table.rows), 2):
            if i + 1 >= len(table.rows): 
                break # 防止表格行数为奇数导致越界
            
            img_row = table.rows[i]
            txt_row = table.rows[i+1]
            
            # 遍历每一列（左图右图）
            for j in range(len(img_row.cells)):
                img_cell = img_row.cells[j]
                txt_cell = txt_row.cells[j]
                
                text = txt_cell.text.strip()
                # 如果这个单元格是空的，跳过
                if not text: 
                    continue
                
                # 正则匹配图名编号
                match = re.search(Config.MATCH_PATTERN, text)
                if match:
                    num = int(match.group(1))
                    # deepcopy 是核心：把图片和格式的底层 XML 完整复制到内存
                    pair_data = [copy.deepcopy(img_cell._tc), copy.deepcopy(txt_cell._tc)]
                    
                    if num in excel_order:
                        self.pairs[num] = pair_data
                    else:
                        self.unmatched.append(pair_data)

    def rebuild(self, excel_order):
        """按照 Excel 顺序组装新的 Word 表格"""
        new_doc = Document()
        # 创建一个 2 列的新表格
        new_table = new_doc.add_table(rows=0, cols=2)
        new_table.style = 'Table Grid' # 加上标准网格线
        
        # 1. 组装最终的数据流（先排匹配的，再把没匹配上的追加在最后）
        final_list = []
        for num in excel_order:
            if num in self.pairs:
                final_list.append(self.pairs[num])
        final_list.extend(self.unmatched)

        # 2. 将数据流重新写入新表格
        # 步长为 2（因为每次填入左右两个图）
        for k in range(0, len(final_list), 2):
            # 新增两行：一行放图片，一行放图名
            r1 = new_table.add_row().cells
            r2 = new_table.add_row().cells
            
            # 填充左列 (数据流中的第 k 个)
            r1[0]._element.getparent().replace(r1[0]._element, final_list[k][0])
            r2[0]._element.getparent().replace(r2[0]._element, final_list[k][1])
            
            # 填充右列 (数据流中的第 k+1 个，如果有的话)
            if k + 1 < len(final_list):
                r1[1]._element.getparent().replace(r1[1]._element, final_list[k+1][0])
                r2[1]._element.getparent().replace(r2[1]._element, final_list[k+1][1])
        
        new_doc.save(Config.OUTPUT_PATH)
        print(f"✅ 处理完成！已生成新文件: {Config.OUTPUT_PATH}")

# ==========================================
# 模块 4：执行入口
# ==========================================
if __name__ == "__main__":
    print("🚀 正在读取数据...")
    
    order = get_excel_sort_order(Config.EXCEL_PATH, Config.EXCEL_COL_NAME)
    
    if not order:
        print("⚠️ 终止：Excel 排序数据为空。")
    else:
        print(f"📊 从 Excel 获取到 {len(order)} 个排序指令，开始重构 Word...")
        processor = WordTableProcessor(Config.WORD_PATH)
        processor.extract_pairs(order)
        processor.rebuild(order)