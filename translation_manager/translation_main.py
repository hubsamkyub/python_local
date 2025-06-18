# translation_main.py

import sys
from PyQt5.QtWidgets import QApplication

# =====================================================================================
# 메인 UI 클래스의 경로가 변경되었으므로 import 경로를 수정합니다.
# from tools.translate_tool_main import TranslationTool -> from ui.main_window import TranslationTool
# =====================================================================================
from ui.main_window import TranslationTool

def main():
    """
    애플리케이션의 메인 진입점입니다.
    QApplication 인스턴스를 생성하고 메인 윈도우를 표시합니다.
    """
    app = QApplication(sys.argv)
    
    main_window = TranslationTool()
    main_window.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()