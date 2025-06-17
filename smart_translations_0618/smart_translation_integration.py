import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

# 모듈 경로 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from smart_translation_manager import SmartTranslationManager, run_smart_translation_manager
from translation_helpers import TranslationHelper


class SmartTranslationIntegration:
    """기존 프로그램과의 통합을 위한 래퍼 클래스"""
    
    def __init__(self, parent_window=None):
        self.parent = parent_window
        self.manager = None
        
    def launch(self):
        """스마트 번역 매니저 실행"""
        if self.parent:
            # 부모 창이 있으면 자식 창으로 실행
            self.manager = run_smart_translation_manager(self.parent)
        else:
            # 독립 실행
            self.manager = run_smart_translation_manager()
            
        return self.manager
    
    def launch_as_tab(self, notebook_widget):
        """노트북 위젯의 탭으로 추가"""
        tab_frame = ttk.Frame(notebook_widget)
        notebook_widget.add(tab_frame, text="스마트 번역")
        
        # 탭 내에 매니저 생성
        self.manager = SmartTranslationManager(tab_frame)
        
        return self.manager
    
    @staticmethod
    def quick_translate(kr_texts, target_langs=None, use_cache=True):
        """빠른 번역 API (다른 모듈에서 사용)"""
        if target_langs is None:
            target_langs = ["EN", "JP"]
            
        helper = TranslationHelper()
        results = {}
        
        # 간단한 캐시 확인 (실제로는 DB 연동 필요)
        for text in kr_texts:
            results[text] = {}
            for lang in target_langs:
                # 여기서는 예시로 빈 값 반환
                # 실제로는 DB 조회 및 API 호출 구현 필요
                results[text][lang] = ""
                
        return results


def add_to_main_menu(menubar, root):
    """메인 프로그램의 메뉴바에 추가"""
    tools_menu = None
    
    # 기존 도구 메뉴 찾기
    for i in range(menubar.index('end') + 1):
        try:
            label = menubar.entryconfig(i, 'label')[-1]
            if '도구' in label:
                tools_menu = menubar.nametowidget(menubar.entryconfig(i, 'menu')[-1])
                break
        except:
            continue
            
    # 도구 메뉴가 없으면 새로 생성
    if tools_menu is None:
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도구", menu=tools_menu)
        
    # 스마트 번역 메뉴 추가
    tools_menu.add_separator()
    tools_menu.add_command(
        label="스마트 번역 관리자",
        command=lambda: SmartTranslationIntegration(root).launch()
    )
    

# 사용 예시 코드
if __name__ == "__main__":
    # 독립 실행 테스트
    print("스마트 번역 관리자 독립 실행 모드")
    
    # 테스트용 메인 창
    test_root = tk.Tk()
    test_root.title("스마트 번역 시스템 테스트")
    test_root.geometry("300x200")
    
    # 버튼들
    ttk.Label(test_root, text="스마트 번역 시스템 테스트", font=('', 14, 'bold')).pack(pady=20)
    
    ttk.Button(
        test_root, 
        text="스마트 번역 관리자 실행",
        command=lambda: SmartTranslationIntegration(test_root).launch()
    ).pack(pady=10)
    
    ttk.Button(
        test_root,
        text="종료",
        command=test_root.quit
    ).pack(pady=10)
    
    # 메뉴바 테스트
    menubar = tk.Menu(test_root)
    test_root.config(menu=menubar)
    add_to_main_menu(menubar, test_root)
    
    test_root.mainloop()


# 기존 프로그램에 통합하는 방법:
"""
1. 메인 프로그램에서 import:
   from smart_translation_integration import SmartTranslationIntegration, add_to_main_menu

2. 메뉴에 추가:
   add_to_main_menu(main_window_menubar, main_window)

3. 또는 버튼으로 실행:
   integration = SmartTranslationIntegration(main_window)
   integration.launch()

4. 탭으로 추가:
   integration = SmartTranslationIntegration()
   integration.launch_as_tab(notebook_widget)
"""