"""项目内复用的正则模式。

把"图 N"这种全项目通用的匹配模式集中在这里，避免在多个工具里抄写同一个字符串、
将来要改格式（比如要支持"图 1-1"）只改一处。
"""
import re

# 匹配中文报告里"图 N / 图N"的题注编号，捕获组 1 是数字部分
FIG_PATTERN_STR: str = r'图\s*(\d+)'
FIG_PATTERN: re.Pattern = re.compile(FIG_PATTERN_STR)
