# Python 极速运行环境配置 (基于 uv)

本项目废弃传统 `pip`，全面采用 `uv` 作为包管理器，以实现毫无污染的虚拟环境隔离和毫秒级依赖安装。

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
在项目根目录下执行以下命令，创建独立运行环境：
```bash
uv venv
```
激活环境（每次重新打开 VS Code 终端时，若未自动挂载则需手动执行）：
```powershell
.venv\Scripts\activate
```
*成功标志：命令行最左侧出现绿色的 `(.venv)` 前缀。*

**步骤 4：项目依赖管理（必须在激活状态下执行）**
* **恢复项目环境**（拉取代码后首选）：
    ```bash
    uv pip install -r requirements.txt
    ```
* **安装新依赖**（如操作 Word 或 Excel）：
    ```bash
    uv pip install python-docx pandas openpyxl
    ```
* **更新环境清单**（安装新库后必须执行）：
    ```bash
    uv pip freeze > requirements.txt
    ```