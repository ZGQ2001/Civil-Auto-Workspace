"""跨模块流转的数据契约。

用 dataclass 取代裸 dict，IDE 能补全/跳转，字段改名一键全改。
"""
from dataclasses import dataclass


@dataclass
class PhotoPair:
    """已排序 Word 表格里一对"图 + 题注"在源表中的位置坐标（0-indexed）。

    img/txt 行列索引以 python-docx 的 row/cell 索引为准；
    传给 win32com 时要 +1（COM 是 1-indexed）。
    """
    num: int            # "图 N" 中的 N
    img_row_idx: int
    txt_row_idx: int
    img_col_idx: int
    txt_col_idx: int
