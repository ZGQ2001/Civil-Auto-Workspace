# Python 极速运行环境配置 (基于 uv)

本项目废弃传统 `pip`，全面采用 `uv` 作为包管理器，以实现毫秒级依赖安装。

**步骤 1：解除 Windows 脚本执行限制（若报错）**
```powershell
Set-ExecutionPolicy RemoteSigned -scope CurrentUser
```

**步骤 2：安装 uv 引擎**
```powershell
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
```
*安装完成后重启终端，输入 `uv --version` 确认安装成功。*

**步骤 3：项目依赖管理**
* **恢复项目环境**（新设备首选）：
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