import os
import time

def backup_current_document(app):
    """
    通用备份模块：
    1. 验证目标文档是否已在本地磁盘驻留（已保存）
    2. 强制同步当前内存数据至本地
    3. 生成附带时间戳（_yyyymmdd_HH时MM分）的安全副本
    4. 仅返回布尔值，不触发任何前端 UI 阻断
    """
    try:
        doc = app.ActiveDocument
        
        # 安全读取属性，拦截 COM 底层异常
        try:
            doc_path = doc.Path
            doc_fullname = doc.FullName
            doc_name = doc.Name
        except Exception:
            doc_path = ""
            doc_fullname = ""
            doc_name = ""

        # 校验物理文件落盘状态
        if not doc_path or doc_fullname == doc_name:
            print("【备份阻断】源文件尚未执行本地存储。")
            return False
        
        # 1. 强制写入当前进度
        doc.Save()
        
        # 2. 构建副本物理地址 
        base, ext = os.path.splitext(doc_fullname)
        timestamp = time.strftime("%Y%m%d_%H时%M分")
        new_path = os.path.abspath(f"{base}_{timestamp}{ext}")
        
        # 3. 后台静默克隆并落盘 (不干涉 ActiveDocument 的焦点)
        backup_doc = app.Documents.Add(Template=doc_fullname)
        backup_doc.SaveAs2(new_path)
        backup_doc.Close(0)
        
        print(f"【系统】文件安全克隆完毕: {new_path}")
        return True
        
    except Exception as e:
        print(f"【系统】克隆进程抛出异常: {e}")
        return False