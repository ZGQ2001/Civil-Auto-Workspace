# Civil-Auto-Workspace (CAW)
## 前言
本人为土木工程检测相关专业从业人员，业余时间利用 Claude 开发了这款自动化工具箱。本文档总结了我在开发过程中遵循的核心原则、验收标准以及工具介绍，旨在确保每一个功能都能真正解决工程师的痛点，并且在实际应用中表现出色。
所有开发都必须围绕两个核心判断：
1. 是否真正减少工程师重复劳动
2. 是否能在真实项目中稳定运行
---
任何新功能上线前，必须能够通过以下六项测试:

1. 新手测试:
一个刚入职的应届生，不看说明书，能否靠“示例模板”和“悬停提示”在 5 分钟内走通核心流程？

2. 极限测试:
导入 500 页报告、200 张图表或上万条数据时，UI 是否保持流畅，不假死，进度是否准确？

3. 变更测试:
甲方要求把所有图表的“红色实线”改为“蓝色虚线”，能否通过修改配置文件在 1 分钟内全局生效？

4. 溯源测试:
最终生成的 Word 报告，能否追溯到具体哪一天、哪个原始文件、哪次处理动作？

5. 断电测试:
软件运行中途被强制关闭后，再次打开时，能否从临时快照或备份中恢复到灾难前状态？

6. 升级测试:
国家或地方规范更新后，能否只替换底层算法模块，而不改动 UI 框架与主体流程？
---
## 目录结构
```
[Project-Root-Name]/                       # 项目根目录
├── .venv/                                 # Python 虚拟环境
├── .vscode/                               # VS Code 配置
├── .env                                   # 本地环境变量，不提交 Git
├── .gitignore                             # Git 忽略规则
├── config.yaml                            # 外部配置文件：业务参数、路径、阈值等
├── pyproject.toml                         # 项目元数据、依赖、构建配置
├── requirements.txt                       # 运行依赖清单
├── README.md                              # 项目说明文档
├── run.bat                                # Windows 一键启动脚本
│
├── scripts/                               # 维护脚本目录
│   ├── build_exe.bat                      # 打包脚本
│   └── init_project.py                    # 自动初始化目录/文件脚本
│
├── data/                                  # 数据仓库，不建议纳入 Git
│   ├── raw/                               # 原始输入数据
│   └── output/                            # 输出结果、生成报告
│
├── logs/                                  # 日志目录
│
├── templates/                             # 模板中心
│   ├── docx/                              # Word 模板
│   └── xlsx/                              # Excel 模板
│
├── tests/                                 # 自动化测试
│   ├── test_core/                         # 业务逻辑测试
│   ├── test_io/                           # IO 层测试
│   └── test_config/                       # 配置加载测试
│
└── src/                                   # 源代码根目录
    └── [package_name]/                    # 项目主包名，必须使用 snake_case
        ├── __init__.py
        ├── main.py                        # 程序唯一入口：初始化配置、日志、主窗口

        ├── app/                           # 应用装配层：负责启动与依赖组织
        │   ├── __init__.py
        │   └── bootstrap.py               # 启动初始化逻辑：配置、日志、窗口挂载

        ├── ui/                            # 视图层：只放界面相关代码
        │   ├── __init__.py
        │   ├── windows/                   # 主窗口、详情窗口等
        │   │   ├── __init__.py
        │   │   ├── main_window.py         # 主界面窗口
        │   │   └── detail_window.py       # 详情/子功能窗口
        │   ├── components/                # 可复用控件
        │   │   ├── __init__.py
        │   │   ├── search_box.py          # 搜索框
        │   │   ├── parameter_panel.py      # 参数配置面板
        │   │   ├── log_panel.py           # 日志面板
        │   │   ├── preview_panel.py       # 预览面板
        │   │   ├── table_editor.py        # 表格编辑器
        │   │   └── progress_bar.py        # 进度条/状态显示组件
        │   ├── models/                    # Qt Model 层
        │   │   ├── __init__.py
        │   │   └── table_models.py        # QAbstractTableModel 等
        │   └── dialogs/                   # 弹窗/设置页
        │       ├── __init__.py
        │       └── settings_dialog.py     # 设置对话框
        │
        ├── core/                          # 核心业务层：纯算法、纯规则、无 UI、无 IO
        │   ├── __init__.py
        │   ├── evaluator.py               # 评定算法
        │   ├── rules.py                   # 业务规则
        │   └── scheduler.py               # 任务调度逻辑
        │
        ├── io/                            # 数据访问层：读取/写出/文件操作
        │   ├── __init__.py
        │   ├── input_handler.py           # 统一输入入口
        │   ├── output_handler.py          # 统一输出入口
        │   ├── excel_reader.py            # Excel 读取
        │   ├── docx_writer.py             # Word 输出
        │   └── file_manager.py            # 文件与目录操作
        │
        ├── models/                        # 数据契约层：统一内部数据结构
        │   ├── __init__.py
        │   └── schema.py                  # dataclass 定义
        │
        ├── config/                        # 配置加载层：读取、校验、合并默认值
        │   ├── __init__.py
        │   └── loader.py                  # config.yaml 加载器
        │
        ├── infra/                         # 基础设施层：通用能力，不放业务
        │   ├── __init__.py
        │   ├── logger.py                 # 日志系统
        │   ├── paths.py                  # 路径统一管理
        │   ├── time_utils.py             # 时间处理（建议带时区）
        │   ├── exceptions.py             # 自定义异常
        │   ├── validators.py             # 通用校验
        │   └── helpers.py                # 零碎通用小工具
        │
        └── resources/                     # 静态资源目录
            ├── icons/                     # 图标资源
            ├── images/                    # 图片资源
            └── styles/                    # 样式资源
                └── qss/                   # Qt 样式表
 
 ```
 