这是一份关于 VS Code 核心插件与 Python 常用第三方库的基础配置手册。内容按工具环境和应用方向分板块整理。

### 一、 VS Code 常用扩展插件 (Extensions)

VS Code 本身是一个轻量级文本编辑器，其对 Python 的支持完全依赖于扩展插件。

**下载方式：**
在 VS Code 左侧活动栏点击“扩展”图标（或使用快捷键 `Ctrl+Shift+X`），在搜索框输入插件名称，点击“安装” (Install) 即可。

**核心插件列表：**
1. **Python (由 Microsoft 提供)**
   * **简介**：VS Code 写 Python 必装的官方插件。提供调试环境配置、环境切换（选择不同的 Python 解释器）以及基础的语法检查支持。
2. **Pylance (由 Microsoft 提供)**
   * **简介**：高性能的 Python 语言服务器。与 Python 插件配合使用，提供极速的自动补全、参数提示、类型检查和代码跳转功能。通常在安装 Python 插件时会自动打包安装。
3. **Jupyter (由 Microsoft 提供)**
   * **简介**：允许在 VS Code 中直接运行 Jupyter Notebook (`.ipynb` 文件)，适合进行数据分析或分步测试代码。
4. **Code Runner**
   * **简介**：代码一键运行工具。右上角会出现一个“播放”按钮，选中一段代码即可快速执行，支持包括 Python 在内的多种语言，省去每次在终端输入命令的麻烦。
5. **Chinese (Simplified) (简体中文) Language Pack**
   * **简介**：VS Code 官方中文语言包，安装后重启软件即可汉化界面。

---

### 二、 Python 常用第三方库 (Libraries)

Python 的强大在于其庞大的第三方生态。这里主要梳理数据处理、文档自动化及日常开发的常用库。

**下载方式：**
由于国内网络环境限制，直接使用 `pip install` 经常会遇到下载缓慢或报错。建议配置国内镜像源（此处以清华源为例）。
在电脑终端（如 CMD、PowerShell 或 VS Code 终端）输入以下命令进行**全局配置**：
```bash
pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
```
配置成功后，以后下载任何库只需输入：`pip install 库名`

#### 1. 数据处理与分析核心库
* **pandas**
  * **简介**：Python 数据处理的基础设施。可以将它理解为“没有图形界面的超强 Excel”。擅长对千万级数据进行筛选、透视、合并和数学计算。
  * **下载**：`pip install pandas`
* **numpy**
  * **简介**：提供高性能的多维数组对象及科学计算工具。通常作为 pandas 的底层支撑库，较少单独用于常规业务逻辑，但在处理大规模矩阵运算时必不可少。
  * **下载**：`pip install numpy`

#### 2. 文件与报表自动化生成库
* **openpyxl**
  * **简介**：专门用于读写 `.xlsx` 格式的 Excel 文件。如果需要精细控制表格样式（如合并单元格、修改字体颜色、画边框）或者向固定的复杂表格模板中填入数据，此库是首选。
  * **下载**：`pip install openpyxl`
* **python-docx**
  * **简介**：专门用于操作 Word 文档 (`.docx`)。可以通过代码自动生成报告段落、插入表格和图片，结合数据处理库可实现从数据清洗到文档生成的全链路自动化。
  * **下载**：`pip install python-docx`

#### 3. 效率与工具库
* **os / pathlib**
  * **简介**：Python **内置库**（无需下载）。用于文件和目录路径操作，比如批量读取一个文件夹下的所有表格名称，或者自动新建当天日期的文件夹。
* **pyinstaller**
  * **简介**：打包工具。如果你写了一个自动化脚本，想要分享给没有安装 Python 环境的同事使用，这个库可以将你的 Python 脚本打包成独立的 `.exe` 可执行文件。
  * **下载**：`pip install pyinstaller`

---

Python 库的生态覆盖非常广，上述仅为基础梳理。