import tkinter as tk
from tkinter import ttk, messagebox

class LanguageSelectionDialog:
    """업데이트할 언어를 선택하는 모달 다이얼로그 (개선된 버전)"""
    def __init__(self, parent, all_langs, title="언어 선택"):  # title 매개변수 추가
        self.top = tk.Toplevel(parent)
        self.top.title(title)  # 동적 제목 설정
        self.top.geometry("400x300")
        self.top.transient(parent)
        self.top.grab_set()

        self.all_langs = all_langs
        self.selected_lang = None  # 단일 선택을 위한 변수 추가
        self.selected_langs = []   # 기존 다중 선택 유지
        self.vars = {}

        main_frame = ttk.Frame(self.top, padding="10")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="분석할 언어를 선택하세요.", wraplength=380).pack(pady=10)

        lang_frame = ttk.Frame(main_frame)
        lang_frame.pack(pady=5)
        
        # 라디오 버튼으로 변경 (단일 선택)
        self.selected_var = tk.StringVar()
        
        # 3열로 배치
        cols = 3
        for i, lang in enumerate(self.all_langs):
            rb = ttk.Radiobutton(lang_frame, text=lang, variable=self.selected_var, value=lang)
            rb.grid(row=i//cols, column=i%cols, padx=10, pady=5, sticky="w")
        
        # 첫 번째 언어를 기본 선택
        if self.all_langs:
            self.selected_var.set(self.all_langs[0])

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side="bottom", fill="x", pady=10)

        ttk.Button(button_frame, text="취소", command=self.cancel).pack(side="right", padx=10)
        ttk.Button(button_frame, text="확인", command=self.ok).pack(side="right")

    def ok(self):
        self.selected_lang = self.selected_var.get()
        if not self.selected_lang:
            messagebox.showwarning("선택 오류", "언어를 선택해야 합니다.", parent=self.top)
            return
        
        # 기존 호환성을 위해 selected_langs도 설정
        self.selected_langs = [self.selected_lang]
        self.top.destroy()

    def cancel(self):
        self.selected_lang = None
        self.selected_langs = []
        self.top.destroy()