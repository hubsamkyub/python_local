import tkinter as tk
from tkinter import ttk

# --- 번역 결과 리포트 다이얼로그 클래스 ---
class TranslationReportDialog:
    def __init__(self, parent, title, report_data):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("500x400")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = ttk.Frame(self.top, padding="15")
        main_frame.pack(fill="both", expand=True)

        # 1. 요약 정보 프레임
        summary_frame = ttk.LabelFrame(main_frame, text="✅ 작업 요약")
        summary_frame.pack(fill="x", pady=5)
        
        summary_labels = [
            ("총 처리 항목:", report_data.get("total_processed", 0)),
            ("TM 사용:", report_data.get("from_tm", 0)),
            ("API 번역:", report_data.get("via_api", 0)),
            ("LLM 후편집:", report_data.get("with_llm", 0)),
            ("실패:", len(report_data.get("failures", [])))
        ]
        
        for i, (text, value) in enumerate(summary_labels):
            ttk.Label(summary_frame, text=text, font=('Malgun Gothic', 10, 'bold')).grid(row=i, column=0, sticky="w", padx=10, pady=3)
            ttk.Label(summary_frame, text=str(value)).grid(row=i, column=1, sticky="e", padx=10, pady=3)
        summary_frame.grid_columnconfigure(1, weight=1)

        # 2. 실패 항목 프레임 (실패 항목이 있을 경우에만 표시)
        failures = report_data.get("failures", [])
        if failures:
            failure_frame = ttk.LabelFrame(main_frame, text="⚠️ 실패 항목")
            failure_frame.pack(fill="both", expand=True, pady=(10, 5))

            text_area = tk.Text(failure_frame, height=8, wrap="word", relief="solid", borderwidth=1)
            text_area.pack(fill="both", expand=True, padx=5, pady=5)
            
            for fail_item in failures:
                text_area.insert(tk.END, f"- {fail_item}\n")
            text_area.config(state="disabled")

        # 3. 확인 버튼
        ttk.Button(main_frame, text="확인", command=self.top.destroy).pack(pady=(15, 0))


class UpdatePreviewDialog:
    def __init__(self, parent, new_entries, updated_entries, visible_langs):
        self.top = tk.Toplevel(parent)
        self.top.title("업데이트 미리보기")
        self.top.geometry("1200x700")
        self.top.transient(parent)
        self.top.grab_set()

        self.new_entries = new_entries
        self.updated_entries = updated_entries
        self.visible_langs = visible_langs
        
        self.confirmed = False
        self.new_to_apply = new_entries
        self.updates_to_apply = updated_entries

        # UI 구성
        main_frame = ttk.Frame(self.top, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        ttk.Label(main_frame, text="엑셀 파일과 마스터 TM을 비교한 결과입니다. 아래 변경사항을 확인하고 '적용' 버튼을 누르세요.").pack(pady=5)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True, pady=10)

        # 신규 항목 탭
        new_tab = ttk.Frame(notebook)
        notebook.add(new_tab, text=f"신규 항목 ({len(new_entries)}개)")
        self.setup_new_entries_tab(new_tab)

        # 변경 항목 탭
        update_tab = ttk.Frame(notebook)
        notebook.add(update_tab, text=f"변경 항목 ({len(updated_entries)}개)")
        self.setup_updated_entries_tab(update_tab)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side="bottom", fill="x", pady=(10,0))
        ttk.Button(button_frame, text="취소", command=self.cancel).pack(side="right", padx=10)
        ttk.Button(button_frame, text="변경사항 적용", command=self.apply).pack(side="right")

    def setup_new_entries_tab(self, parent):
        tree = ttk.Treeview(parent, columns=['KR'] + self.visible_langs, show='headings')
        tree.heading('KR', text='KR')
        tree.column('KR', width=200)
        for lang in self.visible_langs:
            tree.heading(lang, text=lang)
            tree.column(lang, width=120)
        
        for entry in self.new_entries:
            values = [entry.get('KR', '')] + [str(entry.get(lang, '')) for lang in self.visible_langs]
            tree.insert('', 'end', values=values)
        
        tree.pack(fill='both', expand=True)

    def setup_updated_entries_tab(self, parent):
        columns = ['KR', '언어', '기존 번역', '새 번역']
        tree = ttk.Treeview(parent, columns=columns, show='headings')
        for col in columns:
            tree.heading(col, text=col)
        tree.column('KR', width=200)
        tree.column('언어', width=80, anchor='center')
        tree.column('기존 번역', width=300)
        tree.column('새 번역', width=300)
        
        for entry in self.updated_entries:
            kr = entry['kr']
            for lang, change in entry['changes'].items():
                tree.insert('', 'end', values=[kr, lang, change['old'], change['new']])
        
        tree.pack(fill='both', expand=True)

    def apply(self):
        self.confirmed = True
        self.top.destroy()

    def cancel(self):
        self.confirmed = False
        self.top.destroy()
      
  