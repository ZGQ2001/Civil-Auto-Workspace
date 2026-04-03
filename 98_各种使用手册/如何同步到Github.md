这是一份针对代码初学者的 VS Code 连接 GitHub (基于 SSH) 操作手册。内容将严格按照操作逻辑分板块逐步展开。

### 〇、 前置准备与中文路径排雷

1. **安装必要软件**：确保已安装 **VS Code** 和 **Git**，并拥有一个 **GitHub** 账号。
2. **中文路径问题处理**：
   * **项目存储路径**：**严禁**将任何代码项目存放在包含中文的路径下。请在非系统盘新建纯英文路径，例如 `D:\CodeProjects`。
   * **系统用户名中文问题**：如果你的 Windows 用户名是中文，默认的 SSH 密钥会生成在 `C:\Users\你的中文名\.ssh`。虽然新版 Git 通常能识别，但容易引发玄学报错。
   * **建议**：在此次配置中，严格遵循下面的代码操作，如遇公钥读取失败，需在 Git Bash 中通过命令强制指定密钥路径（后文会提及）。

---

### 一、 配置 SSH 密钥（本地连接 GitHub 的核心）

由于要求全部走 SSH 模式，必须先在本地生成物理密钥并绑定到 GitHub。

**1. 生成密钥**
在任意位置右键选择 **"Open Git Bash here"**（或在 VS Code 终端中选择 Git Bash），输入以下命令并回车（替换为你的 GitHub 邮箱）：
```bash
ssh-keygen -t ed25519 -C "your_email@example.com"
```
连续按三次回车，保持默认路径且不设置密码。

**2. 获取公钥内容**
继续在终端输入：
```bash
cat ~/.ssh/id_ed25519.pub
```
终端会输出一串以 `ssh-ed25519` 开头的长字符，将其完整复制。

**3. 将公钥添加到 GitHub**
* 登录 GitHub，点击右上角头像 -> **Settings**。
* 左侧导航栏选择 **SSH and GPG keys** -> 点击 **New SSH key**。
* **Title** 随意填（如 "MyLaptop"），**Key** 框内粘贴刚才复制的公钥代码，点击 **Add SSH key**。

**4. 测试连接**
在 Git Bash 中输入：
```bash
ssh -T git@github.com
```
如果提示 `Are you sure you want to continue connecting (yes/no/[fingerprint])?`，输入 `yes` 并回车。最终看到 `Hi [你的用户名]! You've successfully authenticated...` 即代表连接成功。

---

### 二、 基础配置：告诉 Git 你是谁

在本地终端或 VS Code 终端执行一次即可（替换为你自己的信息）：
```bash
git config --global user.name "你的GitHub用户名"
git config --global user.email "你的GitHub邮箱"
```

---

### 三、 场景实战

#### 场景 1：下载 GitHub 上的已有仓库到本地 (Clone)
1. 在 GitHub 上打开目标仓库，点击绿色的 **"<> Code"** 按钮。
2. **务必选择 SSH 选项卡**，复制以 `git@github.com:` 开头的链接。
3. 在本地纯英文路径下（如 `D:\CodeProjects`），右键打开 Git Bash 或 VS Code 终端。
4. 输入以下命令：
```bash
git clone 刚才复制的SSH链接
```
5. 下载完成后，在 VS Code 中点击“文件” -> “打开文件夹”，选中刚刚下载的仓库文件夹即可开始写代码。

#### 场景 2：将本地现有的文件夹上传为新仓库
1. 在 GitHub 网页端点击 **"+" -> New repository**，填写仓库名，直接点击 **Create repository**（不要勾选添加 README 等任何选项，保持空仓库）。
2. 在 VS Code 中打开你的本地纯英文文件夹。
3. 打开 VS Code 终端 (Ctrl + `)，依次执行：
```bash
# 初始化本地 Git 仓库
git init

# 将所有文件添加到暂存区
git add .

# 提交更改并添加备注
git commit -m "首次提交项目"

# 切换主分支名称为 main (GitHub 默认标准)
git branch -M main

# 关联远程仓库 (复制 GitHub 提示的 SSH 链接)
git remote add origin git@github.com:你的用户名/你的仓库名.git

# 推送代码到 GitHub
git push -u origin main
```

#### 场景 3：改名操作同步
**A. 本地文件夹改名**
Git 追踪的是文件夹内部的文件变化，而不是根文件夹的名字。
1. 直接在电脑系统中把文件夹重命名。
2. 在 VS Code 中重新打开这个改名后的文件夹。
3. 终端执行常规的 `git add .` -> `git commit -m "update"` -> `git push`，不影响任何同步。

**B. GitHub 仓库改名**
1. 在 GitHub 网页端进入仓库 -> **Settings** -> **General** -> 修改 **Repository name** 并点击 Rename。
2. 此时本地仓库关联的旧地址已失效。在本地 VS Code 终端执行以下命令更新 SSH 地址：
```bash
# 获取新的 SSH 链接并在本地重新设置
git remote set-url origin git@github.com:用户名/新仓库名.git

# 验证是否修改成功
git remote -v
```

#### 场景 4：两台电脑同步协作 (A电脑更新，B电脑同步)

假设 A 电脑和 B 电脑都已经配置了各自的 SSH 密钥并添加到了同一个 GitHub 账号下。

**A电脑操作（上传新文件）：**
1. 在 VS Code 中写完新代码/添加新文件。
2. 终端执行“提交三连”：
```bash
git add .
git commit -m "A电脑新增了某个文件"
git push
```

**B电脑操作（获取最新同步）：**
1. 在 B 电脑用 VS Code 打开同一个项目文件夹。
2. **在开始写代码前**，养成先拉取最新代码的习惯。在终端执行：
```bash
git pull
```
3. 此时 B 电脑的本地文件会自动更新为与 GitHub 仓库一致的状态。随后 B 电脑即可继续开发，开发完后同样执行 `add`、`commit`、`push`。

---
**排错建议**：如果在上述任何阶段遇到 VS Code 提示权限拒绝 (Permission denied)，100% 是由于 SSH 密钥未正确配置或未使用 SSH 格式的链接（误用了 HTTPS）。请回到第一步检查 `ssh -T git@github.com` 的连通性。