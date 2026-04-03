# SSH 模式同步核心手册
## 1. 底层基建（仅需配置一次）
针对“张德帅”中文路径 Bug，必须在终端执行此命令，强行指定私钥路径并跳过身份校验：

```Bash
git config --global core.sshCommand "ssh -i ~/.ssh/id_ed25519 -o UserKnownHostsFile=/dev/null -o StrictHostKeyChecking=no"
```

---
## 2. 每日操作闭环
为了保证多端（单位/宿舍）代码不冲突，严格执行以下“三部曲”：

- 开工前 - 拉取 (Pull)：同步云端最新进展。

- 改动后 - 提交 (Commit)：写好备注，点击“提交”按钮（存档本地）。

- 收工前 - 推送 (Push)：点击“同步更改”或蓝色按钮（备份云端）。

---

## 3. 常用排查指令
检查连接状态：`ssh -T git@github.com`（看到 `Hi ZGQ2001!` 即为成功）。

查看远程地址：`git remote -v`（确保是 `git@github.com` 开头的 SSH 地址）。

强制推送测试：`git push -u origin main`。

---

## 4. 新电脑冷启动流程
### 1. 软件环境安装
- 安装 VS Code。

- 安装 Git for Windows（下载时一路回车即可）。

- 安装 Python（记得勾选 `Add Python to PATH`）。

### 2. 申领新电脑的SSH Key

- 在新电脑的终端输入：
```bash
ssh-keygen -t ed25519 -C "z13623133797@gmail.com"
```
一路回车到底，不要设密码。
### 3. 将新房卡备案到 GitHub
- 终端输入：`cat ~/.ssh/id_ed25519.pub`

- 登录 GitHub SSH Settings。

- 点击 New SSH key，标题起名为“宿舍笔记本”或“家里电脑”，把那一长串 `ssh-ed25519` 开头的内容贴进去保存。

### 4. 建立信任与拉取代码
在新电脑你想存放代码的地方（比如 D:\Work），右键打开终端，输入：

```Bash
# 1. 建立信任（第一次会问 yes/no，输入 yes）
ssh -T git@github.com

# 2. 克隆仓库（瞬间移动你的所有代码）
git clone git@github.com:ZGQ2001/Civil_Auto_Workspace.git
```
### 5. 特殊补丁（仅当新电脑用户名也是中文时）
如果新电脑的路径也是 `C:\Users\张德帅\...` 这种中文，请务必执行你在单位电脑运行过的那行“指路命令”：

```Bash
git config --global core.sshCommand "ssh -i ~/.ssh/id_ed25519 -o UserKnownHostsFile=/dev/null -o StrictHostKeyChecking=no"
```
如果新电脑用户名是英文（如 C:\Users\Administrator），则不需要这一步，Git 会自动处理。

-------------

## 5. 在新电脑上开启一个全新的办公自动化项目（比如：`Sanjia_Auto_V2`）。
只需要执行以下三个阶段：

---

## 第一阶段：在新电脑配置 SSH

每台新电脑都需要向 GitHub 申领一张属于自己的“通行证”。

1.  **生成密钥**：在终端输入：
    ```bash
    ssh-keygen -t ed25519 -C "z13623133797@gmail.com"
    ```
    *一路回车，不要设密码。*
2.  **报备 GitHub**：
    * 运行 `cat ~/.ssh/id_ed25519.pub`。
    * 复制那一长串以 `ssh-ed25519` 开头的字符。
    * 贴到 GitHub 网页的 [SSH Settings](https://github.com/settings/keys) 里，取名“新电脑”。
3.  **握手测试**：
    ```bash
    ssh -T git@github.com
    ```
    *看到 `Hi ZGQ2001!` 就代表这台电脑已经具备“合法身份”了。*



---

## 第二阶段：本地新建项目

不要先在网页建仓库，先在本地把结构搭好。

1.  **新建文件夹**：比如 `D:\Civil-Auto-V2`，用 VS Code 打开。
2.  **初始化 Git**：在 VS Code 终端输入：
    ```bash
    git init
    ```
    *这会在文件夹里安装“时光机”。*
3.  **职业习惯：建立忽略名单**：
    新建一个 `.gitignore` 文件，写上：
    ```text
    *.docx
    *.xlsx
    03_doc/
    __pycache__/
    ```
    *确保你的检测数据不被误传。*
4.  **第一次存档**：
    新建一个 `main.py`，随便写点代码。然后在左侧“源代码管理”输入 `init project`，点**提交 (Commit)**。

---

## 第三阶段：关联 GitHub

1.  **网页端建库**：
    去 GitHub 首页点 **"New"**。
    * **Repository name**: 填 `Sanjia_Auto_V2`。
    * **Public/Private**: 必选 **Private**（保护你的 JSA 逻辑）。
    * **不要**勾选 "Initialize this repository with a README"（因为你本地已经建好了）。
2.  **复制 SSH 地址**：
    在新建好的页面，点击 **SSH** 选项卡，复制那串 `git@github.com:ZGQ2001/Sanjia_Auto_V2.git`。
3.  **建立连接并推送**：
    回到 VS Code 终端，输入以下两行：
    ```bash
    git remote add origin git@github.com:ZGQ2001/Sanjia_Auto_V2.git
    git push -u origin main
    ```

---

## 💡 特别提醒 

* **路径补丁**：如果你的新电脑 Windows 用户名还是中文（比如“张德帅”），记得运行我们之前那个 `core.sshCommand` 的长指令，否则 `git push` 依然会报路径乱码错。
* **README 文件**：建议在根目录建一个 `README.md`。用中文写清楚这个项目是解决三检所哪个具体问题的（比如：*“本工具用于自动处理 2026 版房屋检测规范中的倾斜率计算”*）。这会让你看起来像个真正的系统架构师。
* **多项目切换**：当你以后有了 5 个项目，VS Code 左下角的名称会提醒你当前在哪个仓库。记得“收工 Push，开工 Pull”的原则，逻辑就不会乱。
--------