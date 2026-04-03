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
检查连接状态：`ssh -T git@github.com`（看到 Hi ZGQ2001! 即为成功）。

查看远程地址：`git remote -v`（确保是 git@github.com 开头的 SSH 地址）。

强制推送测试：`git push -u origin main`。

---

## 新电脑冷启动流程
### 1. 软件环境安装
- 安装 VS Code。

- 安装 Git for Windows（下载时一路回车即可）。

- 安装 Python（记得勾选 Add Python to PATH）。

### 2. 申领新电脑的SSH Key

- 在新电脑的终端输入：
```bash
ssh-keygen -t ed25519 -C "z13623133797@gmail.com"
```
一路回车到底，不要设密码。
### 3. 将新房卡备案到 GitHub
- 终端输入：`cat ~/.ssh/id_ed25519.pub`

- 登录 GitHub SSH Settings。

- 点击 New SSH key，标题起名为“宿舍笔记本”或“家里电脑”，把那一长串 ssh-ed25519 开头的内容贴进去保存。

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