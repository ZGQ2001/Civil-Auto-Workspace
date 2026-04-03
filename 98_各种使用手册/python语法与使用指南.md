这是一份 Python 基础语法与核心指令的全面指南。内容按编程逻辑划分为核心板块，注重实用性与结构化。



### 一、 基础变量与数据类型 (Data Types)

Python 是动态类型语言，声明变量时无需指定类型。

* **整数 (`int`)**: 表示没有小数部分的数字。
    ```python
    age = 25
    ```
* **浮点数 (`float`)**: 表示带有小数部分的数字。
    ```python
    pi = 3.14159
    ```
* **字符串 (`str`)**: 用于表示文本，使用单引号或双引号包裹。
    ```python
    name = "Python"
    ```
* **布尔值 (`bool`)**: 表示逻辑上的真或假，常用于条件判断。
    ```python
    is_valid = True  # 或 False
    ```

### 二、 核心数据结构 (Data Structures)

用于在单个变量中存储多个元素的容器。

* **列表 (`list`)**: 有序、可变的数据集合，用方括号 `[]` 表示。
    ```python
    fruits = ["apple", "banana", "cherry"]
    fruits.append("orange")  # 在末尾添加元素
    ```
* **元组 (`tuple`)**: 有序、**不可变**的数据集合，用圆括号 `()` 表示。适合存储固定不变的数据。
    ```python
    coordinates = (10.0, 20.0)
    ```
* **字典 (`dict`)**: 无序的键值对 (Key-Value) 集合，用大括号 `{}` 表示。常用于存储结构化属性。
    ```python
    user = {"name": "Alice", "age": 30}
    print(user["name"])  # 读取值
    ```
* **集合 (`set`)**: 无序且**元素唯一**的集合，用大括号 `{}` 表示。常用于去重和关系测试（交集、并集）。
    ```python
    unique_numbers = {1, 2, 2, 3}  # 实际存储为 {1, 2, 3}
    ```

### 三、 基础运算符 (Operators)

* **算术运算符**: `+` (加), `-` (减), `*` (乘), `/` (除, 返回浮点数), `//` (整除, 向下取整), `%` (取模/求余), `**` (幂运算)。
* **比较运算符**: `==` (等于), `!=` (不等于), `>` (大于), `<` (小于), `>=` (大于等于), `<=` (小于等于)。返回布尔值。
* **逻辑运算符**: `and` (与, 两者皆真为真), `or` (或, 一者真即真), `not` (非, 取反)。

### 四、 控制流 (Control Flow)

用于控制代码的执行顺序。

* **条件判断 (`if-elif-else`)**: 根据条件执行不同代码块。注意 Python 依赖缩进划分代码块。
    ```python
    score = 85
    if score >= 90:
        print("优秀")
    elif score >= 60:
        print("及格")
    else:
        print("不及格")
    ```
* **`for` 循环**: 用于遍历序列（如列表、字符串或指定范围）。
    ```python
    for i in range(3):  # range(3) 生成 0, 1, 2
        print(i)
    ```
* **`while` 循环**: 当条件为真时持续执行代码块。
    ```python
    count = 0
    while count < 3:
        print(count)
        count += 1
    ```
* **循环控制 (`break` / `continue`)**:
    * `break`: 立即终止并跳出整个循环。
    * `continue`: 跳过当前循环的剩余部分，直接进入下一次循环。

### 五、 函数 (Functions)

用于封装可重复使用的代码块。

* **常规函数 (`def`)**:
    ```python
    def greet(name):
        return f"Hello, {name}!"  # f-string 用于格式化字符串
    
    result = greet("World")
    ```
* **匿名函数 (`lambda`)**: 适用于简短的、单行的逻辑操作，常配合高阶函数（如 map, filter）使用。
    ```python
    square = lambda x: x ** 2
    print(square(5))  # 输出 25
    ```

### 六、 文件操作 (File Handling)

推荐使用 `with` 上下文管理器，它会在操作完成后自动关闭文件，避免资源泄漏。

* **读取文件 (`r` 模式)**:
    ```python
    with open("data.txt", "r", encoding="utf-8") as file:
        content = file.read()
    ```
* **写入文件 (`w` 模式, 覆盖 / `a` 模式, 追加)**:
    ```python
    with open("output.txt", "w", encoding="utf-8") as file:
        file.write("这是新写入的一行文本。")
    ```

### 七、 异常处理 (Exception Handling)

用于捕获并处理运行时错误，防止程序崩溃。

* **`try-except-finally` 结构**:
    ```python
    try:
        result = 10 / 0
    except ZeroDivisionError:
        print("错误：除数不能为零")
    except Exception as e:
        print(f"发生未知错误: {e}")
    finally:
        print("无论是否报错，这句都会执行")
    ```

### 八、 模块与导入 (Modules & Imports)

Python 通过模块化管理代码。你可以导入内置库、第三方库或自己写的其他 `.py` 文件。

* **导入整个模块**:
    ```python
    import math
    print(math.sqrt(16))  # 输出 4.0
    ```
* **导入模块中的特定功能**:
    ```python
    from datetime import datetime
    print(datetime.now())
    ```
* **导入并重命名 (起别名)**:
    ```python
    import pandas as pd
    ```
    以下是 Python 基础语法在实际开发环境（以 VS Code 为例）中的具体使用方法和落地运行指南。要让代码真正跑起来，通常分为“环境操作”和“逻辑组织”两个层面。

### 一、 在 VS Code 中运行 Python 的三种标准方法

在将代码变为现实之前，你需要先在 VS Code 中新建一个文件，并将文件名后缀命名为 `.py`（例如 `main.py`）。

**方法 1：使用右上角“运行”按钮（最直观）**
* **适用场景**：日常写完脚本直接看结果。
* **操作**：在安装了 Microsoft 官方 Python 插件后，打开 `.py` 文件，点击 VS Code 界面右上角的 **▶ (Run Python File)** 按钮。结果会在下方的终端面板中直接输出。

**方法 2：使用终端命令（最底层、最通用）**
* **适用场景**：需要传递参数，或者在服务器环境、无界面环境下运行。
* **操作**：在 VS Code 中按 `Ctrl + \` 打开终端，确保终端路径与代码文件所在路径一致，输入以下命令并回车：
    ```bash
    python main.py
    ```
    *(注：如果是 Mac 系统，通常输入 `python3 main.py`)*

**方法 3：使用交互式代码块 (Jupyter 风格)**
* **适用场景**：数据分析、分步调试，不需要每次从头运行整个文件。
* **操作**：在 Python 文件中，输入 `#%%` 即可定义一个代码块。点击代码块上方出现的 **Run Cell**，代码会在右侧的交互式窗口中分段执行并保留变量状态（需提前安装 Jupyter 插件）。

---

### 二、 综合使用示范：如何将零散的语法组合成实际工具

单独的语法像零件，实际开发中需要将它们组装成完整的流水线。以下是一个模拟“自动化数据筛查并生成日志”的综合脚本，它融合了**变量、数据结构、控制流、函数、异常处理与文件操作**。

你可以将这段代码完整复制到 `main.py` 中运行：

```python
import os  # 导入内置模块

# 1. 使用函数封装核心处理逻辑
def filter_abnormal_data(data_list, threshold):
    """
    筛选出超出设定阈值的异常数据
    :param data_list: 包含字典的列表 (数据结构)
    :param threshold: 整数或浮点数，报警阈值 (变量)
    :return: 异常数据的名称列表
    """
    abnormal_items = []  # 初始化空列表
    
    # 遍历数据集 (for 循环控制流)
    for item in data_list:
        # 判断数值是否超标 (if 条件判断控制流)
        if item["value"] > threshold:
            abnormal_items.append(item["name"])  # 列表的增加操作
            
    return abnormal_items

# 2. 准备基础数据 (组合使用列表与字典)
# 实际业务中，这些数据通常来自读取 Excel 或数据库
raw_data = [
    {"name": "监测点A", "value": 45.2},
    {"name": "监测点B", "value": 89.5},
    {"name": "监测点C", "value": 30.1},
    {"name": "监测点D", "value": 112.0}
]

# 3. 执行主程序并处理潜在异常 (try-except 异常处理)
try:
    print("开始处理数据...")
    
    # 调用函数，设定阈值为 80
    warning_list = filter_abnormal_data(raw_data, 80)
    
    # 4. 根据处理结果执行文件写入逻辑
    if warning_list:  # 如果列表不为空（存在异常）
        print(f"发现异常数据：{warning_list}，正在生成日志...")
        
        # 写入文本文件 (with 上下文管理器 + 文件操作)
        with open("warning_log.txt", "w", encoding="utf-8") as file:
            file.write("【数据异常警告日志】\n")
            for name in warning_list:
                file.write(f"- {name} 触发了超标警报\n")
                
        print("日志生成完毕。请在当前文件夹查看 warning_log.txt")
    else:
        print("数据全部正常，无需生成日志。")

except Exception as e:
    # 如果运行过程中发生任何错误（例如数据格式不对导致无法对比），防止程序直接崩溃
    print(f"程序运行中断，发生错误：{e}")
```

**运行后的执行流拆解：**
1.  Python 从上到下读取代码，先记住 `filter_abnormal_data` 这个函数的规则，但不执行它。
2.  内存中建立 `raw_data` 变量，装入四个字典。
3.  进入 `try` 块，程序正式开始干活，把 `raw_data` 扔进函数里跑一遍。
4.  函数内部通过 `for` 循环和 `if` 筛选出 "监测点B" 和 "监测点D"，把它们装进新列表退还给主程序。
5.  主程序发现新列表有内容，触发文件生成动作，在你的电脑上实际创建一个 `warning_log.txt` 文件并写入汉字。