"""
===============================================================================
脚本名称：Word 运行环境保护模块 (word_env_utils.py)
作者：ZGQ
功能概述：
    提供用于 Word 自动化操作的上下文管理器，实现底层执行效率优化与状态安全兜底。
    
    这个脚本相当于排版引擎的“加速器”和“安全锁”。它会在排版开始前，自动关掉 
    Word 耗时的后台计算（如拼写检查、自动分页等）让程序全速运行；并在排版结
    束或中途报错被掐断时，强制把 Word 恢复到正常的亮屏可用状态，彻底杜绝 Word 
    假死或卡顿报错。所有涉及 Word 批量处理的脚本都可以直接套用它。
===============================================================================
"""

from contextlib import contextmanager
import win32com.client

@contextmanager
def word_optimized_environment(app):
    """
    Word 排版加速与安全恢复的上下文管理器。
    
    作用：
    1. 进入该环境时，自动关闭 Word 的后台计算和屏幕刷新。
    2. 退出该环境时（无论是正常结束还是抛出异常/被手动中断），强制恢复 Word 原有状态。
    """
    # ====== 进入环境：挂起 Word ======
    try:
        # 保存原有的状态（虽然通常都是 True，但更严谨的做法是保存当前状态）
        original_updating = app.ScreenUpdating
        original_pagination = app.Options.Pagination
        original_spelling = app.Options.CheckSpellingAsYouType
        original_grammar = app.Options.CheckGrammarAsYouType
        
        # 强制关闭，全速运行
        app.ScreenUpdating = False
        app.Options.Pagination = False
        app.Options.CheckSpellingAsYouType = False
        app.Options.CheckGrammarAsYouType = False
        
        # 把控制权交还给主程序（即 yield 后面的代码开始执行）
        yield 
        
    # ====== 退出环境：安全恢复 ======
    finally:
        try:
            app.ScreenUpdating = True # 强制亮屏
            app.Options.Pagination = original_pagination
            app.Options.CheckSpellingAsYouType = original_spelling
            app.Options.CheckGrammarAsYouType = original_grammar
        except Exception as e:
            print(f"【环境恢复异常】Word 状态恢复失败: {e}")