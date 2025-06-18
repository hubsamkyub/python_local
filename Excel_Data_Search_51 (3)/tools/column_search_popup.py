import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os

class ColumnSearchPopup:
    def __init__(self, master, excel_cache):
        self.excel_cache = excel_cache
        self.top = tk.Toplevel(master)
        self.top.title("🔍 컬럼명 검색기")
        self.top.geometry("900x600")

        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar()

        # 검색 입력 UI
        input_frame = tk.Frame(self.top)
        input_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(input_frame, text="컬럼명:").pack(side="left")
        tk.Entry(input_frame, textvariable=self.search_var, width=30).pack(side="left", padx=5)
        tk.Button(input_frame, text="검색", command=self.search_columns).pack(side="left", padx=5)

        # 결과 테이블
        self.tree = ttk.Treeview(self.top, columns=("파일명", "시트명", "컬럼목록"), show="headings")
        self.tree.heading("파일명", text="파일명")
        self.tree.heading("시트명", text="시트명")
        self.tree.heading("컬럼목록", text="컬럼목록")
        self.tree.column("파일명", width=180)
        self.tree.column("시트명", width=150)
        self.tree.column("컬럼목록", width=500)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree.bind("<Double-1>", self.open_excel_file)

        # 상태바
        tk.Label(self.top, textvariable=self.status_var, anchor="w", relief="sunken").pack(fill="x")

    def search_columns(self):
        keyword = self.search_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        if not keyword:
            messagebox.showwarning("입력 필요", "검색할 컬럼명을 입력해주세요.")
            return

        count = 0
        for file, info in self.excel_cache.items():
            for sheet, meta in info.get("sheets", {}).items():
                columns = meta.get("columns", [])
                if any(keyword in str(col).lower() for col in columns if col is not None):
                    self.tree.insert("", "end", values=(file, sheet, ", ".join(columns)), tags=(info["path"],))
                    count += 1

        self.status_var.set(f"🔍 검색 완료: {count}개 시트에서 일치하는 컬럼 발견")

    def open_excel_file(self, event):
        item = self.tree.selection()[0]
        file_path = self.tree.item(item, "tags")[0]
        try:
            subprocess.Popen(['start', '', file_path], shell=True)
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 열기 실패: {e}")
