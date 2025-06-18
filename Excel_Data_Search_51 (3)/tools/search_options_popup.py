import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

class SearchOptionsPopup:
    def __init__(self, master, folder_path, files_manager, excel_cache):
        self.master = master
        self.folder_path = folder_path
        self.files_manager = files_manager
        self.excel_cache = excel_cache
        
        self.top = tk.Toplevel(master)
        self.top.title("검색 옵션 설정")
        self.top.geometry("600x500")
        self.top.transient(master)
        self.top.grab_set()
        
        self.build_ui()
        self.refresh_file_list()
        
    def build_ui(self):
        # 상단 설명 레이블
        tk.Label(self.top, text="검색할 파일 목록을 관리합니다").pack(pady=10)
        
        # 파일 추가 프레임
        add_frame = tk.Frame(self.top)
        add_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(add_frame, text="파일 추가:").pack(side="left")
        self.file_entry = tk.Entry(add_frame, width=40)
        self.file_entry.pack(side="left", padx=5)
        
        tk.Button(add_frame, text="추가", command=self.add_file).pack(side="left")
        tk.Button(add_frame, text="찾기", command=self.browse_file).pack(side="left", padx=5)
        
        # 폴더 추가 프레임 (새로 추가)
        folder_frame = tk.Frame(self.top)
        folder_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(folder_frame, text="폴더 추가:").pack(side="left")
        self.folder_entry = tk.Entry(folder_frame, width=40)
        self.folder_entry.pack(side="left", padx=5)
        
        tk.Button(folder_frame, text="추가", command=self.add_folder).pack(side="left")
        tk.Button(folder_frame, text="찾기", command=self.browse_folder).pack(side="left", padx=5)
        
        
        # 파일 목록 프레임
        list_frame = tk.Frame(self.top)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 목록 헤더와 상태
        header_frame = tk.Frame(list_frame)
        header_frame.pack(fill="x")
        
        tk.Label(header_frame, text="검색할 파일 목록").pack(side="left")
        self.status_label = tk.Label(header_frame, text="")
        self.status_label.pack(side="right")
        
        # 목록과 스크롤바
        self.listbox_frame = tk.Frame(list_frame)
        self.listbox_frame.pack(fill="both", expand=True)
        
        self.scrollbar = tk.Scrollbar(self.listbox_frame)
        self.scrollbar.pack(side="right", fill="y")
        
        self.file_listbox = tk.Listbox(self.listbox_frame, selectmode="extended", 
                                      yscrollcommand=self.scrollbar.set)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        self.scrollbar.config(command=self.file_listbox.yview)
        
        # 버튼 프레임
        btn_frame = tk.Frame(self.top)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Button(btn_frame, text="선택 삭제", command=self.remove_selected).pack(side="left")
        tk.Button(btn_frame, text="전체 삭제", command=self.clear_all).pack(side="left", padx=5)
        tk.Button(btn_frame, text="폴더 내 모든 Excel 추가", 
                 command=self.add_all_excel).pack(side="left")
        
        # 닫기 버튼
        tk.Button(self.top, text="저장 후 닫기", command=self.save_and_close).pack(pady=10)
        
    def refresh_file_list(self):
        """파일 목록을 새로고침합니다."""
        self.file_listbox.delete(0, tk.END)
        files = self.files_manager.get_search_files()
        
        for file in sorted(files):
            self.file_listbox.insert(tk.END, file)
            
        self.status_label.config(text=f"총 {len(files)}개 파일")
        
    def add_file(self):
        """사용자 입력으로 파일을 추가합니다."""
        filename = self.file_entry.get().strip()
        if not filename:
            messagebox.showwarning("입력 오류", "파일명을 입력하세요.")
            return
            
        success, message = self.files_manager.add_search_file(filename)
        if success:
            self.file_entry.delete(0, tk.END)
            self.refresh_file_list()
        
        messagebox.showinfo("알림", message)
        

    def browse_file(self):
        """여러 파일 찾기 다이얼로그를 엽니다."""
        from tkinter import filedialog
        file_paths = filedialog.askopenfilenames(
            initialdir=self.folder_path,
            title="Excel 파일 선택 (여러 개 선택 가능)",
            filetypes=(("Excel 파일", "*.xlsx"), ("모든 파일", "*.*"))
        )
        
        if file_paths:
            # 선택한 파일 모두 추가
            added_count = 0
            skipped_count = 0
            
            for file_path in file_paths:
                # 상대 경로로 변환
                if file_path.startswith(self.folder_path):
                    rel_path = os.path.relpath(file_path, self.folder_path)
                    success, _ = self.files_manager.add_search_file(rel_path)
                    if success:
                        added_count += 1
                    else:
                        skipped_count += 1
                else:
                    skipped_count += 1
            
            self.refresh_file_list()
            
            if added_count > 0:
                messagebox.showinfo("파일 추가 완료", 
                                f"총 {added_count}개의 Excel 파일이 추가되었습니다.\n"
                                f"(중복/오류로 건너뛴 파일: {skipped_count}개)")

    def remove_selected(self):
        """선택한 파일을 목록에서 제거합니다."""
        selected = self.file_listbox.curselection()
        if not selected:
            messagebox.showinfo("선택 없음", "삭제할 항목을 선택하세요.")
            return
            
        # 역순으로 제거 (인덱스 변화 방지)
        for i in sorted(selected, reverse=True):
            filename = self.file_listbox.get(i)
            self.files_manager.remove_search_file(filename)
            
        self.refresh_file_list()
        
    def clear_all(self):
        """모든 파일을 목록에서 제거합니다."""
        if messagebox.askyesno("확인", "모든 파일을 목록에서 제거하시겠습니까?"):
            for file in self.files_manager.get_search_files()[:]:
                self.files_manager.remove_search_file(file)
            self.refresh_file_list()
            
    def add_all_excel(self):
        """폴더 내 모든 Excel 파일을 추가합니다."""
        if messagebox.askyesno("확인", "폴더 내 모든 Excel 파일을 목록에 추가하시겠습니까?"):
            added = 0
            for file in os.listdir(self.folder_path):
                if file.endswith('.xlsx'):
                    success, _ = self.files_manager.add_search_file(file)
                    if success:
                        added += 1
                        
            self.refresh_file_list()
            messagebox.showinfo("완료", f"{added}개의 Excel 파일이 추가되었습니다.")
            
    def save_and_close(self):
        """설정을 저장하고 창을 닫습니다."""
        self.files_manager.save_search_files()
        self.top.destroy()
        
    
    def browse_folder(self):
        """여러 폴더 찾기 다이얼로그를 엽니다."""
        from tkinter import filedialog
        
        # 폴더 선택 다이얼로그를 반복적으로 열기
        folder_paths = []
        while True:
            folder_path = filedialog.askdirectory(
                initialdir=self.folder_path,
                title="Excel 파일이 있는 폴더 선택 (취소 버튼으로 완료)"
            )
            
            if not folder_path:  # 취소 버튼을 누르면
                break
                
            folder_paths.append(folder_path)
            
            # 사용자에게 계속할지 물어보기
            if not messagebox.askyesno("폴더 추가", "폴더가 추가되었습니다. 더 추가하시겠습니까?"):
                break
        
        if folder_paths:
            # 모든 선택된 폴더 처리
            self._process_selected_folders(folder_paths)


    def _process_selected_folders(self, folder_paths):
        """선택한 여러 폴더의 Excel 파일을 추가합니다."""
        total_added = 0
        total_skipped = 0
        
        for folder_path in folder_paths:
            # 상대 경로로 변환할 수 있는지 확인
            if not folder_path.startswith(self.folder_path):
                messagebox.showwarning("경로 오류", 
                                    f"선택한 폴더가 기준 폴더({self.folder_path}) 내에 있지 않아 건너뜁니다:\n{folder_path}")
                continue
            
            # 폴더 내 모든 Excel 파일 찾기 (하위 폴더 포함)
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.endswith('.xlsx') and not file.startswith('~$'):
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, self.folder_path)
                        
                        success, _ = self.files_manager.add_search_file(rel_path)
                        if success:
                            total_added += 1
                        else:
                            total_skipped += 1
        
        self.refresh_file_list()
        
        if total_added > 0:
            messagebox.showinfo("폴더 추가 완료", 
                            f"총 {total_added}개의 Excel 파일이 추가되었습니다.\n"
                            f"(중복/오류로 건너뛴 파일: {total_skipped}개)")
        else:
            messagebox.showinfo("알림", f"추가할 Excel 파일을 찾지 못했습니다.")


    def add_folder(self):
        """사용자가 지정한 폴더 내의 모든 Excel 파일을 추가합니다."""
        folder_name = self.folder_entry.get().strip()
        if not folder_name:
            messagebox.showwarning("입력 오류", "폴더명을 입력하세요.")
            return
        
        # 쉼표나 세미콜론으로 구분된 여러 폴더 처리
        folder_names = folder_name.replace(';', ',').split(',')
        folder_paths = []
        
        for name in folder_names:
            name = name.strip()
            if not name:
                continue
                
            target_folder = os.path.join(self.folder_path, name)
            
            if not os.path.exists(target_folder):
                messagebox.showwarning("폴더 오류", f"'{name}' 폴더를 찾을 수 없습니다.")
                continue
            
            if not os.path.isdir(target_folder):
                messagebox.showwarning("폴더 오류", f"'{name}'은 폴더가 아닙니다.")
                continue
                
            folder_paths.append(target_folder)
        
        if folder_paths:
            self._process_selected_folders(folder_paths)
        
        # 입력 필드 초기화
        self.folder_entry.delete(0, tk.END)