"""
편집 관련 다이얼로그들
"""
import tkinter as tk
from tkinter import ttk, messagebox


class InlineEditDialog:
    """번역 내용 직접 편집 다이얼로그"""
    
    def __init__(self, parent, trans_item, visible_langs, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("번역 직접 편집")
        self.top.geometry("800x600")
        self.top.transient(parent)
        self.top.grab_set()
        
        self.trans_item = trans_item
        self.visible_langs = visible_langs
        self.callback = callback
        self.lang_vars = {}
        
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.top, padding="15")
        main_frame.pack(fill="both", expand=True)
        
        # 기본 정보 표시
        info_frame = ttk.LabelFrame(main_frame, text="기본 정보")
        info_frame.pack(fill="x", pady=5)
        
        ttk.Label(info_frame, text=f"STRING_ID: {self.trans_item['STRING_ID']}", 
                 font=("맑은 고딕", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        
        kr_label = ttk.Label(info_frame, text=f"KR: {self.trans_item['KR']}", 
                           wraplength=700, justify="left")
        kr_label.pack(anchor="w", padx=10, pady=5)
        
        # 번역 편집 섹션
        edit_frame = ttk.LabelFrame(main_frame, text="번역 편집")
        edit_frame.pack(fill="both", expand=True, pady=10)
        
        # 스크롤 가능한 프레임
        canvas = tk.Canvas(edit_frame)
        scrollbar = ttk.Scrollbar(edit_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 각 언어별 편집 필드
        for i, lang in enumerate(self.visible_langs):
            lang_frame = ttk.Frame(scrollable_frame)
            lang_frame.pack(fill="x", padx=10, pady=5)
            
            ttk.Label(lang_frame, text=f"{lang}:", font=("맑은 고딕", 10, "bold")).pack(anchor="w")
            
            # 텍스트 입력 위젯 (여러 줄 지원)
            text_widget = tk.Text(lang_frame, height=3, wrap="word", font=("맑은 고딕", 9))
            text_widget.pack(fill="x", pady=2)
            
            # 기존 번역 내용 설정
            current_text = self.trans_item["translations"].get(lang, "")
            text_widget.insert("1.0", current_text)
            
            self.lang_vars[lang] = text_widget
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="취소", command=self.cancel).pack(side="right", padx=5)
        ttk.Button(button_frame, text="저장", command=self.save).pack(side="right")
        ttk.Button(button_frame, text="원본 복원", command=self.restore_original).pack(side="left")
    
    def save(self):
        """편집 내용 저장"""
        # 각 언어별 입력 내용 수집
        for lang, text_widget in self.lang_vars.items():
            content = text_widget.get("1.0", "end-1c").strip()
            self.trans_item["translations"][lang] = content
        
        # 상태 업데이트
        self.trans_item["method"] = "직접편집"
        self.trans_item["status"] = "[수정완료]"
        
        # 콜백 호출
        if self.callback:
            self.callback(self.trans_item)
        
        self.top.destroy()
    
    def cancel(self):
        """편집 취소"""
        self.top.destroy()
    
    def restore_original(self):
        """TM의 원본 내용으로 복원"""
        if not messagebox.askyesno("복원 확인", "TM의 원본 내용으로 복원하시겠습니까?"):
            return
        
        # TM에서 원본 데이터 가져오기 (부모 클래스의 translation_memory 접근 필요)
        # 이 부분은 부모 클래스 참조가 필요하므로 생략하고 현재 내용 유지
        messagebox.showinfo("알림", "현재 버전에서는 지원되지 않습니다.")



class GlossaryEditDialog:
    def __init__(self, parent, title, initial_data=None):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()

        self.result = None
        self.vars = {}
        # DB 컬럼 순서 (description 제외, 추후 추가 가능)
        self.db_cols = ["string_id", "kr", "en", "cn", "tw", "th", "pt", "es", "de", "fr", "jp", "engine", "contributor", "update_at", "verified"]

        main_frame = ttk.Frame(self.top, padding="15")
        main_frame.pack(fill="both", expand=True)

        # 2열 그리드로 필드 배치
        for i, field in enumerate(self.db_cols):
            row, col = i // 2, (i % 2) * 2
            ttk.Label(main_frame, text=f"{field.upper()}:").grid(row=row, column=col, sticky="w", padx=5, pady=2)
            
            if field == "verified":
                var = tk.BooleanVar()
                ttk.Checkbutton(main_frame, variable=var).grid(row=row, column=col+1, sticky="w", padx=5, pady=2)
            else:
                var = tk.StringVar()
                entry = ttk.Entry(main_frame, textvariable=var, width=40)
                entry.grid(row=row, column=col+1, sticky="ew", padx=5, pady=2)
                if field == "string_id" and initial_data: # 편집 모드일 때 string_id는 수정 불가
                    entry.config(state="readonly")

            self.vars[field] = var
        
        # 초기 데이터 설정 (편집 모드)
        if initial_data:
            for field, value in initial_data.items():
                if field in self.vars:
                    if field == "verified":
                        self.vars[field].set(True if value == "Y" else False)
                    else:
                        self.vars[field].set(value)

        # 저장/취소 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=(len(self.db_cols)//2) + 1, column=0, columnspan=4, pady=(20, 0))
        ttk.Button(button_frame, text="저장", command=self.save).pack(side="left", padx=10)
        ttk.Button(button_frame, text="취소", command=self.top.destroy).pack(side="left")

    def save(self):
        data = {}
        for field, var in self.vars.items():
            val = var.get()
            if field == "verified":
                data[field] = 1 if val else 0
            else:
                data[field] = val
        
        if not data.get("string_id"):
            messagebox.showwarning("입력 오류", "string_id는 필수 항목입니다.", parent=self.top)
            return
        
        self.result = data
        self.top.destroy()