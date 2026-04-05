# word_env_utils.py
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