"""
===============================================================================
脚本名称：Word 运行环境保护模块 (word_env_utils.py)
作者：ZGQ
功能概述：
    提供用于 Word 自动化操作的上下文管理器，实现底层执行效率优化与状态安全兜底。
    V3.0 工业级优化版：引入防御性状态存取与全量弹窗静默，完美适配 WPS 复杂接口。
===============================================================================
"""

from contextlib import contextmanager

@contextmanager
def word_optimized_environment(app):
    """
    Word 排版加速与安全恢复的上下文管理器。
    """
    # 使用字典记录成功获取到的原始状态，避免获取失败导致变量未定义
    states = {}
    
    # ====== 进入环境：挂起 Word ======
    try:
        # 1. 最关键的基础挂起
        try:
            states['ScreenUpdating'] = app.ScreenUpdating
            app.ScreenUpdating = False
        except: pass
        
        # 2. 【核心优化】弹窗静默 (极其重要：防止后台系统级弹窗导致假死)
        try:
            states['DisplayAlerts'] = app.DisplayAlerts
            app.DisplayAlerts = 0  # 0 = wdAlertsNone
        except: pass

        # 3. 【核心优化】耗时选项挂起 (采用防御性读取，防止部分精简版 WPS 接口报错中断引擎)
        try:
            states['Pagination'] = app.Options.Pagination
            app.Options.Pagination = False
        except: pass
        
        try:
            states['CheckSpelling'] = app.Options.CheckSpellingAsYouType
            app.Options.CheckSpellingAsYouType = False
        except: pass

        try:
            states['CheckGrammar'] = app.Options.CheckGrammarAsYouType
            app.Options.CheckGrammarAsYouType = False
        except: pass

        # 把控制权交还给主程序
        yield 
        
    # ====== 退出环境：安全恢复 ======
    finally:
        # 【核心优化】必须将所有状态的恢复独立进行
        # 防止某一个不兼容属性恢复失败时，抛出异常导致后面的 ScreenUpdating 无法亮屏
        try:
            if 'CheckGrammar' in states: app.Options.CheckGrammarAsYouType = states['CheckGrammar']
        except: pass
        
        try:
            if 'CheckSpelling' in states: app.Options.CheckSpellingAsYouType = states['CheckSpelling']
        except: pass
        
        try:
            if 'Pagination' in states: app.Options.Pagination = states['Pagination']
        except: pass
        
        try:
            if 'DisplayAlerts' in states: app.DisplayAlerts = states['DisplayAlerts']
        except: pass
        
        # 最核心的屏幕刷新必须独立保证执行
        try:
            app.ScreenUpdating = True
        except: pass