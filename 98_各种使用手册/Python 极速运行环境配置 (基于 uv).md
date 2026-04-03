# Python 极速运行环境配置 (基于 uv)

本项目废弃传统 `pip`，全面采用 `uv` 作为包管理器。
* **基本概念**：`uv` 是极速构建环境的“包工头”，`.venv` 文件夹才是真正隔离全局系统的“无菌沙盒”。

**步骤 1：解除 Windows 脚本执行限制（新电脑必执行）**
由于激活虚拟环境需要运行 `.ps1` 脚本，必须先放开 PowerShell 权限：
```powershell
Set-ExecutionPolicy RemoteSigned -scope CurrentUser
```

**步骤 2：全局安装 uv 引擎**
推荐直接使用系统自带的 pip 进行极速安装：
```bash
pip install uv
```
*安装完成后重启终端，输入 `uv --version` 确认安装成功。*

**步骤 3：构建并激活项目虚拟沙盒（核心步骤）**
在项目根目录下执行以下命令，让 uv 圈出一块独立运行地基：
```bash
uv venv
```
激活环境（每次重新打开 VS Code 终端时，若未自动挂载则需手动执行）：
```powershell
.venv\Scripts\activate
```
*成功标志：命令行最左侧出现绿色的 `(.venv)` 前缀。*

**步骤 4：项目依赖管理（必须在带有绿字前缀时执行）**
此处根据你的实际开发场景，分为两种操作流：

**▶ 场景 A：从零开发新功能（进货并造册）**
当你在这个项目里需要用到新的第三方库时：
1. **安装新库**（例如操作 Word、Excel、PDF 的库）：
    ```bash
    uv pip install python-docx pandas openpyxl pypdf
    ```
2. **生成环境清单**（安装完必须立刻执行，生成或更新 `requirements.txt`）：
    ```bash
    uv pip freeze > requirements.txt
    ```

**▶ 场景 B：跨设备克隆旧项目（按清单一键还原）**
当你在三检所或宿舍的新电脑上，刚从 GitHub 拉取了本项目，且文件夹内**已有** `requirements.txt` 时：
* 直接运行以下命令，包工头会根据清单，秒速复刻一模一样的开发环境：
    ```bash
    uv pip install -r requirements.txt
    ```
