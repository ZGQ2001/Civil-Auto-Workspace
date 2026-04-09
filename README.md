# 🏗️ Civil Auto Workspace (CAW)
> **工程报告自动化排版矩阵 V2.0**  
> *Transforming Engineering Reports from Manual Labor to Digital Precision.*

---

## 🌟 项目愿景 | Project Vision
在传统的工程检测与鉴定工作中，报告排版往往占据了工程师大量的时间成本。**Civil Auto Workspace (CAW)** 旨在通过一套高度模块化的自动化体系，将机械化的排版工作转化为秒级的数字化生成，让工程师回归专业分析，而非陷于格式泥潭。

### 核心哲学：解耦与适配
CAW 采用了**「配置-执行」解耦架构**。通过将排版标准（字体、字号、间距等）完全定义在 `JSON` 配置文件中，实现了逻辑代码与业务标准的彻底分离。这意味着：**无需修改一行代码，即可通过调整配置适配任何一套新的报告标准。**

---

## 🛠️ 技术栈 | Tech Stack
- **Language:** Python 3.11+
- **Core Libraries:** 
  - `pywin32` (COM Interface for Word/WPS)
  - `pypdf` & `PyMuPDF (fitz)` (High-precision PDF processing)
  - `Pillow` (Visual coordinate mapping)
  - `tkinter` (Lightweight GUI Control Panel)

---

## 📂 模块矩阵 | Module Matrix

### 1. 统一调度中心 (`main.py`)
项目唯一的入口。采用异步唤醒机制，通过图形化面板一键调度所有子模块，实现任务的流水线操作。

### 2. 智能排版引擎 (`02_Core`)
- **正文引擎 (`body_format.py`)**: 基于语义层级的智能识别，自动处理多级标题、图表标注及缩进，确保文档结构绝对一致。
- **表格引擎 (`table_format.py`)**: 针对工程表格的专项优化。实现全量宽度规整、单元格对齐，并集成**空值自动标红**的质量校验功能。

### 3. 专项纠偏套件 (`02_Core`)
- **规范校对 (`bracket_format.py`)**: 针对工程文档中极易出错的括号全半角进行全局纠偏，同时智能保护技术参数，确保规范代号不被误改。
- **引用修复 (`fix_cross_ref.py`)**: 解决 Word 交叉引用在更新后易崩坏的痛点，通过强制追加 `\\* MERGEFORMAT` 确保格式稳定性。

### 4. PDF 处理链路 (`02_Core`)
- **高效转换 (`word2pdf.py`)**: 集成多线程批量转换与合并逻辑，自带引擎重启机制，确保在处理大规模文档时的鲁棒性。
- **坐标拾取 (`pdf_coordinate_picker.py`)**: 可视化精准定位，为自动化填充提供高精度的坐标映射基准。

---

## 🚀 快速部署 | Deployment

### 环境准备
建议使用 `uv` 或 `venv` 构建隔离环境：
```bash
# 快速安装依赖
pip install -r requirements.txt
```

### 启动流程
1. 确保目标文档已保存并处于 **打开状态**（程序通过 COM 接口实时通信）。
2. 运行根目录下的 `运行主程序.bat`。
3. 在 GUI 面板中选择所需模块 $\to$ 点击执行。

---

## 🛡️ 安全机制 | Safety First
本项目内置 **`file_utils_backup.py`** 模块。在任何破坏性写入操作执行前，系统将自动进行**静默克隆备份**。即使发生意外，也能在毫秒级内恢复至操作前状态，确保工程数据的绝对安全。

---

## ✍️ 作者 | Author
**ZGQ**  
*Building the future of engineering automation.*
