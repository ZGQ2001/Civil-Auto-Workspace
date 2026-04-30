---
name: 工程检测报告自动化项目背景
description: 用户在做的"工程自动化"项目的领域背景与代码组织 —— 桥梁/土建检测报告批量处理工具集
type: project
originSessionId: c9988218-6cb4-4f0c-81d3-1dca18da21b6
---
用户在做"市政工程检测报告自动化"项目，主要业务是：把现场拍的照片 + Excel 缺陷清单，自动生成符合检测公司规范的 Word 附录文档。已实现工具集中放在 `D:\01_VC CODE\Civil_Auto_Workspace\02_Core\` 目录，主入口是 `main.py`（customtkinter dashboard，subprocess 启动各工具）。

**Why:** 用户是检测院/检测公司的从业者（北京市建设工程质量第三检测所相关），工作中要反复处理同类报告 —— 排照片、改题注、修引用、转 PDF、加水印等等，纯人工很耗时。这套工具集的目标是让单次报告制作从"半天"压缩到"分钟级"。

**How to apply:**
- 涉及"图 N"题注、缺陷清单、附录排版的需求，优先考虑能否在现有 `sort_photos.py` / `renumber_photos.py` 之上扩展，而不是另起炉灶
- 现有工具已覆盖：报告正文排版（body_format）、表格排版（table_format）、括号纠偏（bracket_format）、交叉引用修复（fix_cross_ref）、Word↔PDF（word2pdf）、PNG 坐标选取（coord_picker）、手写模拟（auto_filler_v2）、照片排序（sort_photos）、照片重编号（renumber_photos）
- 用户在 2026-04-30 完成了第一次架构层重构，沉淀了 `common/` 公共包；后续新工具按 `feedback_toolbox_architecture.md` 的 5 条原则写
- 用户已提到的下一批候选工具：①Word/Excel 一致性校核 ②图片批量导出归档 ③sort+renumber 一键流水线 —— 没排优先级，要做时再问

**项目环境：**
- Python 3.12.13 + uv 管理的虚拟环境（`D:\01_VC CODE\Civil_Auto_Workspace\.venv\`）
- Windows 11，git 仓库，主分支 main
- 重度依赖 win32com（Word COM 自动化）、python-docx、openpyxl、pandas、customtkinter
- 编辑器是 VS Code + Pylance（类型存根偶尔会误报，比如 sys.stdout.reconfigure）
