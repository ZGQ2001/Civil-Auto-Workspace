from ui_components import ModernHandwriteDialog

def main():
    # 唤起我们刚做好的第一阶段配置面板
    dialog = ModernHandwriteDialog()
    result = dialog.show()
    
    if result:
        print("用户配置收集完毕！这是传给下一阶段的数据：")
        print(result)

if __name__ == "__main__":
    main()