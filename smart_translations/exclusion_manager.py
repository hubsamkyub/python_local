import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os
import re # 정규식 유효성 검사를 위해 임포트
import datetime # 생성일자 저장을 위해 임포트

class ExclusionManager:
    """
    제외 패턴을 관리하는 독립적인 Tkinter 창입니다.
    KR 및 번역(LANG) 패턴을 추가, 삭제, 조회할 수 있습니다.
    패턴은 SQLite 데이터베이스에 저장됩니다.
    """
    def __init__(self, parent_root, db_path="exclusion_patterns.db"):
        self.parent_root = parent_root
        self.db_path = db_path
        self.conn = None
        
        # DB 연결 및 테이블 초기화
        self._initialize_db()

        # Tkinter Toplevel 창 생성
        self.root = tk.Toplevel(self.parent_root)
        self.root.title("제외 패턴 관리")
        self.root.geometry("800x600")
        self.root.transient(self.parent_root) # 부모 창 위에 항상 있도록 설정
        self.root.grab_set() # 부모 창과 상호작용하지 못하도록 설정하여 모달처럼 동작

        self.patterns = {"KR": [], "LANG": []} # 메모리에 로드된 패턴 저장
        self._setup_ui() # UI 구성
        self._load_patterns_from_db() # DB에서 초기 패턴 로드
        self.populate_treeviews() # Treeview에 패턴 표시

        # 창이 닫힐 때 DB 연결을 안전하게 종료합니다.
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _initialize_db(self):
        """
        DB 파일이 없으면 생성하고, 필요한 테이블을 초기화합니다.
        """
        try:
            self.conn = sqlite3.connect(self.db_path)
            cursor = self.conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS exclusion_patterns (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    pattern_type TEXT NOT NULL, -- 'KR' 또는 'LANG'
                    pattern TEXT NOT NULL UNIQUE, -- 정규식 패턴
                    description TEXT, -- 패턴에 대한 설명
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP -- 생성일자
                )
            """)
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("DB 오류", f"제외 패턴 DB 초기화 중 오류 발생: {e}")
            self.conn = None # 연결 실패 시 conn을 None으로 설정

    def _setup_ui(self):
        """
        제외 패턴 관리 창의 UI를 구성합니다.
        KR 패턴과 번역 패턴을 위한 탭을 포함합니다.
        """
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 탭을 위한 Notebook 위젯 생성
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=10)

        # KR 제외 패턴 탭
        self.kr_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.kr_tab, text="KR 제외 패턴")
        self._setup_pattern_tab(self.kr_tab, "KR")

        # 번역 제외 패턴 탭
        self.lang_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.lang_tab, text="번역 제외 패턴")
        self._setup_pattern_tab(self.lang_tab, "LANG")

    def _setup_pattern_tab(self, parent_frame, pattern_type):
        """
        각 패턴 유형(KR 또는 LANG) 탭의 UI를 구성합니다.
        패턴 입력, 설명 입력, 추가 버튼, Treeview (목록), 삭제 버튼을 포함합니다.
        """
        # 패턴 추가 입력 프레임
        input_frame = ttk.LabelFrame(parent_frame, text=f"{pattern_type} 패턴 추가", padding=10)
        input_frame.pack(fill="x", pady=5)

        ttk.Label(input_frame, text="패턴 (정규식):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        pattern_entry = ttk.Entry(input_frame, width=50)
        pattern_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(input_frame, text="설명:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        description_entry = ttk.Entry(input_frame, width=50)
        description_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        add_button = ttk.Button(input_frame, text="추가", 
                                command=lambda: self.add_pattern(pattern_type, pattern_entry.get(), description_entry.get()))
        add_button.grid(row=0, column=2, rowspan=2, padx=10, ipady=5)

        input_frame.grid_columnconfigure(1, weight=1) # 패턴 입력 필드가 확장되도록 설정

        # 패턴 목록을 표시할 Treeview 프레임
        tree_frame = ttk.Frame(parent_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)

        columns = ("ID", "패턴", "설명", "생성일")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        tree.heading("ID", text="ID")
        tree.column("ID", width=50, stretch=tk.NO) # ID 컬럼 너비 고정
        tree.heading("패턴", text="패턴")
        tree.column("패턴", width=250, stretch=tk.YES)
        tree.heading("설명", text="설명")
        tree.column("설명", width=300, stretch=tk.YES)
        tree.heading("생성일", text="생성일")
        tree.column("생성일", width=150, stretch=tk.NO) # 생성일 컬럼 너비 고정

        # 스크롤바 추가
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        # 삭제 버튼
        delete_button = ttk.Button(parent_frame, text="선택 항목 삭제", 
                                   command=lambda: self.delete_selected_pattern(pattern_type, tree))
        delete_button.pack(pady=5)

        # 각 패턴 유형에 맞는 입력 필드와 Treeview 참조 저장
        if pattern_type == "KR":
            self.kr_pattern_entry = pattern_entry
            self.kr_description_entry = description_entry
            self.kr_tree = tree
        else: # LANG
            self.lang_pattern_entry = pattern_entry
            self.lang_description_entry = description_entry
            self.lang_tree = tree

    def _load_patterns_from_db(self):
        """
        SQLite DB에서 제외 패턴을 로드하여 self.patterns 딕셔너리에 저장합니다.
        """
        if not self.conn:
            return

        self.patterns = {"KR": [], "LANG": []} # 기존 패턴 클리어
        cursor = self.conn.cursor()
        try:
            cursor.execute("SELECT id, pattern_type, pattern, description, created_at FROM exclusion_patterns")
            for row in cursor.fetchall():
                id, p_type, pattern, desc, created_at = row
                self.patterns[p_type].append({"id": id, "pattern": pattern, "description": desc, "created_at": created_at})
        except sqlite3.Error as e:
            messagebox.showerror("DB 로드 오류", f"제외 패턴 로드 중 오류 발생: {e}")

    def populate_treeviews(self):
        """
        현재 self.patterns에 로드된 패턴을 Treeview 위젯에 표시합니다.
        """
        # KR Treeview 업데이트
        self.kr_tree.delete(*self.kr_tree.get_children())
        for p in self.patterns["KR"]:
            self.kr_tree.insert("", "end", iid=p["id"], values=(p["id"], p["pattern"], p["description"], p["created_at"]))

        # LANG Treeview 업데이트
        self.lang_tree.delete(*self.lang_tree.get_children())
        for p in self.patterns["LANG"]:
            self.lang_tree.insert("", "end", iid=p["id"], values=(p["id"], p["pattern"], p["description"], p["created_at"]))

    def add_pattern(self, pattern_type, pattern_text, description):
        """
        새로운 패턴을 DB에 추가하고 UI를 갱신합니다.
        정규식 유효성 검사를 수행합니다.
        """
        pattern_text = pattern_text.strip()
        if not pattern_text:
            messagebox.showwarning("입력 오류", "패턴은 비워둘 수 없습니다.", parent=self.root)
            return

        # 정규식 유효성 검사
        try:
            re.compile(pattern_text)
        except re.error as e:
            messagebox.showerror("정규식 오류", f"유효하지 않은 정규식 패턴입니다:\n{e}", parent=self.root)
            return

        if not self.conn:
            messagebox.showerror("DB 오류", "데이터베이스 연결이 초기화되지 않았습니다.", parent=self.root)
            return

        cursor = self.conn.cursor()
        try:
            cursor.execute("INSERT INTO exclusion_patterns (pattern_type, pattern, description) VALUES (?, ?, ?)",
                           (pattern_type, pattern_text, description.strip()))
            self.conn.commit()
            messagebox.showinfo("추가 성공", "패턴이 성공적으로 추가되었습니다.", parent=self.root)
            
            # UI 및 메모리 갱신
            self._load_patterns_from_db() # 최신 DB 상태를 메모리에 반영
            self.populate_treeviews() # Treeview를 다시 그려줌
            
            # 입력 필드 초기화
            if pattern_type == "KR":
                self.kr_pattern_entry.delete(0, tk.END)
                self.kr_description_entry.delete(0, tk.END)
            else: # LANG
                self.lang_pattern_entry.delete(0, tk.END)
                self.lang_description_entry.delete(0, tk.END)

        except sqlite3.IntegrityError:
            messagebox.showwarning("중복 오류", "동일한 패턴이 이미 존재합니다.", parent=self.root)
        except sqlite3.Error as e:
            messagebox.showerror("DB 오류", f"패턴 추가 중 오류 발생: {e}", parent=self.root)

    def delete_selected_pattern(self, pattern_type, tree):
        """
        선택된 패턴을 DB에서 삭제하고 UI를 갱신합니다.
        """
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("선택 오류", "삭제할 패턴을 선택하세요.", parent=self.root)
            return

        if not messagebox.askyesno("삭제 확인", "선택된 패턴을 정말 삭제하시겠습니까?", parent=self.root):
            return

        if not self.conn:
            messagebox.showerror("DB 오류", "데이터베이스 연결이 초기화되지 않았습니다.", parent=self.root)
            return

        cursor = self.conn.cursor()
        try:
            for item_id in selected_items:
                # Treeview의 item id는 DB의 id와 동일하게 설정했으므로 바로 사용 가능
                pattern_id = tree.item(item_id, "iid")
                cursor.execute("DELETE FROM exclusion_patterns WHERE id = ?", (pattern_id,))
            self.conn.commit()
            messagebox.showinfo("삭제 완료", "선택된 패턴이 성공적으로 삭제되었습니다.", parent=self.root)
            
            # UI 및 메모리 갱신
            self._load_patterns_from_db()
            self.populate_treeviews()

        except sqlite3.Error as e:
            messagebox.showerror("DB 오류", f"패턴 삭제 중 오류 발생: {e}", parent=self.root)

    def get_kr_patterns(self):
        """현재 로드된 KR 제외 패턴 목록(정규식 문자열 리스트)을 반환합니다."""
        # _load_patterns_from_db()를 호출하여 최신 상태를 보장할 수 있습니다.
        # 그러나 성능을 위해 process_source_dbs에서 한 번만 호출하는 것이 좋습니다.
        # 여기서는 이미 메모리에 로드된 패턴을 반환합니다.
        return [p["pattern"] for p in self.patterns["KR"]]

    def get_lang_patterns(self):
        """현재 로드된 LANG 제외 패턴 목록(정규식 문자열 리스트)을 반환합니다."""
        return [p["pattern"] for p in self.patterns["LANG"]]
        
    def _on_closing(self):
        """
        제외 패턴 관리 창이 닫힐 때 호출되며, DB 연결을 종료합니다.
        """
        if self.conn:
            self.conn.close()
            self.conn = None
        self.root.destroy()

# ExclusionManager 단독 테스트를 위한 코드 (main 앱에서 실행 시에는 사용되지 않음)
if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw() # 메인 루트 창을 숨김 (ExclusionManager는 Toplevel로 뜨므로)
    app = ExclusionManager(root)
    root.mainloop()