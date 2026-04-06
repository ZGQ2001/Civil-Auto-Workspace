"""
===============================================================================
脚本名称：仿生手写记录表生成器 (auto_filler.py)
核心修复：
    1. 解决空白问题：修复了 V9.4 图层叠加逻辑错误导致的文字不显示。
    2. 小数点归位：放弃单字拆分，回归整行渲染，确保小数点老老实实待在数字下方。
    3. 拒绝裁切：大幅增加贴纸缓冲区，并修正粘贴偏移坐标，确保长文本不再被切断。
    4. 紧凑间距：通过 word_spacing 参数调节，实现自然紧凑的手写感。
===============================================================================
"""
import os           # 处理文件路径
import json         # 读取坐标映射
import random       # 生成手写随机抖动
import pandas as pd  # 处理 Excel 数据
from PIL import Image, ImageFont, ImageOps  # 图像处理核心库
from handright import Template, handwrite    # 仿生手写核心引擎

# ---------------- 板块 1：全局路径配置 ----------------

# 自动定位项目主文件夹
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = os.path.join(BASE_DIR, "01_Input")
OUTPUT_DIR = os.path.join(BASE_DIR, "03_Output")

# 文件路径
JSON_PATH = os.path.join(INPUT_DIR, "record_mapping.json")
BASE_IMG_PATH = os.path.join(INPUT_DIR, "记录表.png")
EXCEL_PATH = os.path.join(INPUT_DIR, "平均值生成检测值.xlsm")
FONT_PATH = os.path.join(INPUT_DIR, "font.ttf") 
OUTPUT_PDF = os.path.join(OUTPUT_DIR, "自动生成_检测记录表.pdf")

# 【目标 Sheet】：直接在这里改名字就行了，不要动其他代码了
TARGET_SHEET = "Sheet2"

# --- 仿生视觉参数微调 ---
VERTICAL_DRIFT_FIX = -1.5   # 全局上下移动（负数上移，正数下移）
GLOBAL_SIZE_LIMIT = 1.48    # 全局字号放大倍数
ITEMS_PER_PAGE = 8          # 每张纸放 8 组数据

# ---------------- 板块 2：仿生引擎核心函数 ----------------

def load_config():
    """读取 JSON 里的坐标和字号信息"""
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    
    # 将 Excel 字段映射到 JSON 里的框选区域
    base_boxes = {
        "name": cfg["构件名称"]["box"],
        "val1": cfg["测点1"]["box"],
        "val2": cfg["测点2"]["box"],
        "val3": cfg["测点3"]["box"],
        "avg": cfg["平均值"]["box"]
    }
    # 计算向右(dx)和向下(dy)跨步的距离
    dx = cfg["右"]["box"][0] - cfg["构件名称"]["box"][0]
    dy = (cfg["下"]["box"][1] - cfg["构件名称"]["box"][1]) + VERTICAL_DRIFT_FIX
    
    # 算出最终字号
    font_size = int(cfg["构件名称"].get("font_size", 35) * GLOBAL_SIZE_LIMIT)
    return base_boxes, dx, dy, font_size

def create_handwritten_sticker(text, box, font_path, font_size, fatigue_idx=0, total_items=1):
    """
    仿生渲染核心：负责把文字变成一张高清、抗锯齿、带随机抖动的透明贴纸。
    """
    w, h = int(box[2]), int(box[3])
    clean_text = str(text).strip()
    
    # 如果没数据，返回空图
    if clean_text in ['nan', 'None', '']:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0))

    # 计算“写累了”的疲劳系数
    fatigue_boost = min(1.3, 1.0 + (fatigue_idx / total_items))
    
    current_size = font_size
    final_wrapped_text = ""
    
    # --- 阶段 A：智能排版逻辑 (挤一挤 -> 换行 -> 缩字号) ---
    while current_size > 10:
        font = ImageFont.truetype(font_path, current_size)
        # 如果长度超出不多（允许 20% 溢出），强行一行写完，这最像真人
        if font.getlength(clean_text) <= (w * 1.20):
            final_wrapped_text = clean_text
            lines_count = 1
            break
        
        # 否则尝试换行
        lines = []
        cur_line = ""
        for char in clean_text:
            if font.getlength(cur_line + char) <= (w * 1.1):
                cur_line += char
            else:
                if cur_line: lines.append(cur_line)
                cur_line = char
        lines.append(cur_line)
        
        # 检查换行后的高度
        line_h = int(current_size * 1.15)
        if (len(lines) * line_h) <= (h + 10):
            final_wrapped_text = "\n".join(lines)
            lines_count = len(lines)
            break
        current_size -= 2 # 实在塞不下才把字写小
    else:
        final_wrapped_text, lines_count, current_size = clean_text, 1, 12
        font = ImageFont.truetype(font_path, 12)

    # --- 阶段 B：高清抗锯齿渲染 (解决像素低、锯齿重的问题) ---
    # 创建超大画布，防止笔画被切断
    canvas_w, canvas_h = w + 200, h + 150
    bg = Image.new("L", (canvas_w, canvas_h), 255) # 灰度画布
    
    # 计算文字在格子里的垂直居中位置
    total_text_h = lines_count * int(current_size * 1.15)
    top_pos = (canvas_h - total_text_h) // 2

    # 设置手写引擎参数
    template = Template(
        background=bg, font=font, line_spacing=int(current_size * 1.15),
        fill=0, # 渲染黑字
        left_margin=100, top_margin=top_pos, # 100 是为了在超大画布中心书写
        word_spacing=-2, # 【关键】：设置负间距，让字迹更紧凑自然
        line_spacing_sigma=1.0, font_size_sigma=1.1, word_spacing_sigma=1.2,
        perturb_x_sigma=1.5 * fatigue_boost, 
        perturb_y_sigma=1.2 * fatigue_boost, 
        perturb_theta_sigma=0.04 * fatigue_boost
    )
    
    try:
        # 渲染原始笔触
        raw_img = list(handwrite(final_wrapped_text, template))[0]
        
        # 保留每一个边缘像素，彻底消除低像素锯齿感
        mask = ImageOps.invert(raw_img) 
        ink_layer = Image.new("RGBA", (canvas_w, canvas_h), (35, 35, 35, 255))
        sticker = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
        sticker.paste(ink_layer, (0, 0), mask)
        
        return sticker
    except:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0))

# ---------------- 板块 3：主程序自动化流程 ----------------

def main():
    # 准备输出环境
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    
    print(f"🚀 正在启动 ...")
    base_boxes, dx, dy, final_font_size = load_config()
    
    # 读取 Excel 数据，指定 Sheet2
    df = pd.read_excel(EXCEL_PATH, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
    df[0] = df[0].ffill() # 处理合并单元格
    
    # 1. 准备一个空的“总清单”，用来存放后面整理出来的构件数据包。
    items = []

# 2. 开启核心循环：从第 0 行开始，到表格结束，每次“跨步”跳过 4 行走。
#    为什么要跳 4 行？因为 Excel 里每 4 行才构成一个完整的检测单元。
    for i in range(0, len(df), 4):

    # 3. 安全卫士：检查剩下的行数还够不够 4 行。
    #    如果剩下的行数凑不齐 4 行了，说明数据已经拿完了，直接跳出循环，防止电脑报错。
        if i + 3 >= len(df): break

    # 4. 裁切动作：从总表里精准地把这连着的 4 行数据整体抠出来，存入临时变量 chunk（小块）中。
        chunk = df.iloc[i:i+4]

    # 5. 组装动作：从抠出来的这 4 行里，按照坐标位置提取精华，打包成一个“身份卡”存入清单。
        items.append({
        "name": chunk.iloc[0, 0], # 拿第 1 行、第 1 列的数据，也就是“构件名称”
        "val1": chunk.iloc[0, 2], # 拿第 1 行、第 3 列的数据，也就是“实测值 1”
        "val2": chunk.iloc[1, 2], # 拿第 2 行、第 3 列的数据，也就是“实测值 2”
        "val3": chunk.iloc[2, 2], # 拿第 3 行、第 3 列的数据，也就是“实测值 3”
        "avg":  chunk.iloc[3, 2]  # 拿第 4 行、第 3 列的数据，也就是最后的“平均值”
        })
    
    total_count = len(items)
    pages = []
    current_page_img = None
    
    # 循环填表
    for idx, item in enumerate(items):
        pos_index = idx % ITEMS_PER_PAGE
        
        # 换页逻辑
        if pos_index == 0:
            if current_page_img: pages.append(current_page_img.convert('RGB'))
            current_page_img = Image.open(BASE_IMG_PATH).convert('RGBA')
            print(f"📄 正在合成第 {len(pages) + 1} 页 ...")
            
        col, row = pos_index % 2, pos_index // 2
        cur_dx, cur_dy = col * dx, row * dy
        
        for key in ["name", "val1", "val2", "val3", "avg"]:
            bx, by, bw, bh = base_boxes[key]
            tx, ty = int(bx + cur_dx), int(by + cur_dy)
            
            # 生成完美的贴纸
            sticker = create_handwritten_sticker(item[key], [tx, ty, bw, bh], FONT_PATH, final_font_size, idx, total_count)
            
            # 【重要修正】：粘贴坐标偏移对齐
            # 因为贴纸画布左右多加了 100 像素，上下多加了 75 像素（中心点对齐）
            # 所以粘贴时要往回挪 100 和 75，才能让字落在格子正中
            current_page_img.paste(sticker, (tx - 100, ty - 75), sticker)
            
    # 保存成品 PDF
    if current_page_img: pages.append(current_page_img.convert('RGB'))
    if pages:
        pages[0].save(OUTPUT_PDF, save_all=True, append_images=pages[1:], resolution=300)
        print(f"结束！请查收：{OUTPUT_PDF}")

if __name__ == "__main__":
    main()