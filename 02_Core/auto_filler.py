"""
===============================================================================
脚本名称：仿生手写记录表生成器
功能通俗解释：
    1. 像人一样“预判”：它不是数几个字，而是量每一个字有多少毫米（像素），
       如果剩下的字很短，它会像人一样在格子边缘“挤一挤”写完。
    2. 高清不模糊：采用工业级的“蒙版”技术，保证印出来的字边缘锐利，没有锯齿。
    3. 自动填表：读取 Excel，自动换行，自动缩放，最后直接吐出一份排版完美的 PDF。
===============================================================================
"""

import os           # 操作系统工具：用来处理电脑里的文件夹和文件路径
import json         # JSON工具：用来读取你之前框选保存的坐标文件
import random       # 随机工具：用来产生手写时的随机抖动，让每一笔都不一样
import pandas as pd  # 数据工具：大名鼎鼎的表格处理库，专门用来读 Excel
from PIL import Image, ImageFont, ImageOps  # 图像工具：用来贴图、写字、反转颜色
from handright import Template, handwrite    # 手写引擎：核心技术，把电脑字变成手写体

# ---------------- 板块 1：告诉电脑文件都在哪 ----------------

# 自动定位项目的主文件夹路径
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 拼接各个文件夹的路径，这样写代码在任何人的电脑上都能跑通，不会因为路径报错
INPUT_DIR = os.path.join(BASE_DIR, "01_Input")   # 放入底图和 Excel 的地方
OUTPUT_DIR = os.path.join(BASE_DIR, "03_Output") # 生成成品 PDF 的地方

# 具体的各个文件身份证（路径）
JSON_PATH = os.path.join(INPUT_DIR, "record_mapping.json")       # 坐标配置文件
BASE_IMG_PATH = os.path.join(INPUT_DIR, "记录表.png")            # 你的空白底图
EXCEL_PATH = os.path.join(INPUT_DIR, "平均值生成检测值.xlsm")    # 你的数据源
FONT_PATH = os.path.join(INPUT_DIR, "font.ttf")                  # 你的手写字体文件
OUTPUT_PDF = os.path.join(OUTPUT_DIR, "自动生成_检测记录表.pdf") # 最终生成的成品名字

# --- 核心物理微调参数 ---
VERTICAL_DRIFT_FIX = -1.2    # 垂直修正：如果发现字越往后越压线，就把这个数调大（如-1.5）
GLOBAL_SIZE_LIMIT = 1.45     # 全局字号放大：想让全表的字整体变大，就改这个倍数
ITEMS_PER_PAGE = 8           # 每张纸放几个构件的数据：你的表是左4右4，所以是8个

# ---------------- 板块 2：功能函数（电脑的“大脑”逻辑） ----------------

def load_config():
    """这个函数负责把 JSON 里的坐标读出来，并计算出每个格子该挪多远"""
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        cfg = json.load(f)  # 打开并读取 JSON 文件内容
    
    # 将 Excel 里的列名和你框选的坐标名字一一对应
    base_boxes = {
        "name": cfg["构件名称"]["box"], # 对应构件名称的大格子
        "val1": cfg["测点1"]["box"],   # 对应实测值1的小格子
        "val2": cfg["测点2"]["box"],   # 对应实测值2的小格子
        "val3": cfg["测点3"]["box"],   # 对应实测值3的小格子
        "avg": cfg["平均值"]["box"]     # 对应平均值格子
    }
    
    # 计算“跨步”距离：第1个构件到右边、下边格子的像素差
    dx = cfg["右"]["box"][0] - cfg["构件名称"]["box"][0] # 向右挪多远
    dy = (cfg["下"]["box"][1] - cfg["构件名称"]["box"][1]) + VERTICAL_DRIFT_FIX # 向下挪多远
    
    # 计算最终要用的基准字号
    font_size = int(cfg["构件名称"].get("font_size", 35) * GLOBAL_SIZE_LIMIT)
    
    return base_boxes, dx, dy, font_size

def create_handwritten_image(text, box, font_path, font_size):
    """这是全脚本最聪明的地方：它负责把一段话变成一张带手写感的透明贴纸"""
    
    # 获取格子的宽度和高度像素
    w, h = int(box[2]), int(box[3])
    clean_text = str(text).strip() # 去掉文字前后的空格
    
    # 如果 Excel 里这格是空的，就直接返回一张透明的空图
    if clean_text in ['nan', 'None', '']:
        return Image.new("RGBA", (w, h), (0, 0, 0, 0))

    current_size = font_size # 先用标准字号试试
    final_wrapped_text = ""  # 准备存放换行后的文字
    
    # --- 阶段 A：物理长度排版决策（核心 AI 逻辑） ---
    # 如果字号太大塞不下，这个循环会让字号一点点变小，直到塞进去为止
    while current_size > 10:
        font = ImageFont.truetype(font_path, current_size) # 加载当前字号的字体
        
        # 1. 测量这串字的总物理长度（单位是像素）
        full_width = font.getlength(clean_text)
        
        # 【仿生挤字逻辑】：如果全长只比格子宽了 20% 以内，人通常会选择“挤一挤”不换行
        if full_width <= (w * 1.20):
            final_wrapped_text = clean_text
            final_lines_count = 1
            break # 决定了，不换行！
            
        # 2. 如果实在太长，开始计算在哪里折行
        lines = []
        current_line = ""
        chars = list(clean_text)
        
        for i, char in enumerate(chars):
            test_line = current_line + char
            test_w = font.getlength(test_line) # 量一量加上这个字后的长度
            
            # 预判：看看剩下的字还有多长？
            remaining_text = "".join(chars[i+1:])
            remaining_w = font.getlength(remaining_text)
            
            # 如果当前行写满了
            if test_w > w:
                # 【预判挤字】：如果剩下的字总长度小于 40 像素（大概 1-2 个字），人会硬挤在这一行写完
                if remaining_w < 40: 
                    current_line = test_line
                    continue
                else:
                    # 剩下的字还挺多，老老实实换行吧
                    if current_line: lines.append(current_line)
                    current_line = char
            else:
                current_line = test_line # 还没满，继续写
                
        if current_line: lines.append(current_line)
        
        # 3. 检查换行后的总高度有没有超出格子
        line_h = int(current_size * 1.15)
        if (len(lines) * line_h) <= (h + 6):
            final_wrapped_text = "\n".join(lines)
            final_lines_count = len(lines)
            break # 换行成功，高度也合适！
            
        current_size -= 2 # 换了行还太高？那就把字号变小一点再重来
    else:
        # 万一字号缩到最小还不行，就强制用 12 号字单行显示
        final_wrapped_text, final_lines_count, current_size = clean_text, 1, 12
        font = ImageFont.truetype(font_path, 12)

    # --- 阶段 B：高保真去锯齿渲染（让字变清晰） ---
    
    # 为了防止手写的笔画勾到格子外面被切断，我们先把贴纸画布做大一点
    canvas_w, canvas_h = w + 200, h + 40
    bg = Image.new("L", (canvas_w, canvas_h), 255) # 创建一张纯白底色图
    
    # 计算文字在格子里的垂直居中位置
    total_h = final_lines_count * int(current_size * 1.15)
    top_pos = (h - total_h) // 2
    
    # 设置手写模板参数：让每一格的字都有微小的旋转、抖动和间距变化
    template = Template(
        background=bg, font=font, line_spacing=int(current_size * 1.15),
        fill=0, # 这里的 0 代表纯黑墨水
        left_margin=8, top_margin=top_pos + 20, # 给左边留点缝，模拟人写字不贴边
        word_spacing=-1, # 字间距紧凑一点更自然
        line_spacing_sigma=1, font_size_sigma=1, word_spacing_sigma=1, # 抖动参数
        perturb_x_sigma=1.5, perturb_y_sigma=1, perturb_theta_sigma=0.03 # 笔画偏移和旋转
    )
    
    try:
        # 调用核心引擎，把字“写”出来
        raw_img = list(handwrite(final_wrapped_text, template))[0]
        
        # 【关键技术：蒙版粘贴】：
        # 这一步不是简单的抠图，而是把文字的深浅直接变成透明度。
        # 这样能保留字体边缘最细微的“羽化”效果，印出来才像真的墨水，不模糊。
        mask = ImageOps.invert(raw_img) # 把黑字白底反转成白字黑底
        ink_layer = Image.new("RGBA", (canvas_w, canvas_h), (35, 35, 35, 255)) # 准备深灰色的墨水层
        sticker = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))        # 准备全透明的底图
        sticker.paste(ink_layer, (0, 0), mask) # 隔着蒙版把墨水“刷”上去
        return sticker # 返回这张做好的“手写贴纸”
    except:
        # 出错时的兜底方案，返回空白
        return Image.new("RGBA", (w, h), (0, 0, 0, 0))

# ---------------- 板块 3：主程序流程（开始干活） ----------------

def main():
    # 如果还没建输出文件夹，就建一个
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    
    # 1. 第一步：加载所有的坐标和配置
    base_boxes, dx, dy, final_font_size = load_config()
    
    # 2. 第二步：读取 Excel 数据，专门指定读取 Sheet2
    df = pd.read_excel(EXCEL_PATH, sheet_name="Sheet2", header=None, engine='openpyxl')
    # 【高亮】：ffill() 会处理 Excel 里的合并单元格，把空缺的名字自动补齐
    df[0] = df[0].ffill() 
    
    # 将 Excel 里的原始数据每 4 行打包成一个“构件对象”
    items = []
    for i in range(0, len(df), 4):
        if i + 3 >= len(df): break # 防止表格末尾有残缺行
        chunk = df.iloc[i:i+4] # 截取这 4 行
        items.append({
            "name": chunk.iloc[0, 0], # 构件名字
            "val1": chunk.iloc[0, 2], # 数据1
            "val2": chunk.iloc[1, 2], # 数据2
            "val3": chunk.iloc[2, 2], # 数据3
            "avg":  chunk.iloc[3, 2]  # 平均值
        })
    
    pages = [] # 用来存放每一页做好的图片
    current_page_img = None # 当前正在画的那一页
    
    # 遍历处理每一个构件数据
    for idx, item in enumerate(items):
        # 确定这个构件在这一页的哪个位置 (0到7)
        pos_index = idx % ITEMS_PER_PAGE
        
        # 如果当前位置是 0，说明要新开一张纸了
        if pos_index == 0:
            if current_page_img: pages.append(current_page_img.convert('RGB')) # 把画好的前一页存起来
            current_page_img = Image.open(BASE_IMG_PATH).convert('RGBA') # 拿出一张新的空白底图
            print(f"📄 正在生成第 {len(pages) + 1} 页...")
            
        # 计算当前格子相对于基准格子的平移量
        col, row = pos_index % 2, pos_index // 2 # 算出它是第几列、第几行
        cur_dx, cur_dy = col * dx, row * dy     # 算出要挪动的像素值
        
        # 遍历需要填写的 5 个格子（名称、三个实测、平均值）
        for key in ["name", "val1", "val2", "val3", "avg"]:
            bx, by, bw, bh = base_boxes[key] # 获取基础位置
            tx, ty = int(bx + cur_dx), int(by + cur_dy) # 计算在这一页的具体坐标
            
            # 调用之前的“大脑”，做出这张手写贴纸
            sticker = create_handwritten_image(item[key], [tx, ty, bw, bh], FONT_PATH, final_font_size)
            
            # 把贴纸啪地一下贴到底图上。ty-20 是因为我们做贴纸时画布多留了 20 像素的余量
            current_page_img.paste(sticker, (tx, ty - 20), sticker)
            
    # 处理最后一页数据
    if current_page_img: pages.append(current_page_img.convert('RGB'))
    
    # 3. 第三步：把所有画好的图片叠在一起，保存成多页 PDF
    if pages:
        pages[0].save(OUTPUT_PDF, save_all=True, append_images=pages[1:], resolution=300)
        print(f"🎉 任务完美结束！请去这里看结果：{OUTPUT_PDF}")

# 程序启动入口
if __name__ == "__main__":
    main()