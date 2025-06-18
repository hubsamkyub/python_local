import tkinter as tk
from tkinter import ttk
import os
import threading
from utils.config_utils import save_search_history

class SearchUI:
    """검색 관련 UI 및 기능을 담당하는 클래스"""
    
    def __init__(self, parent):
        """
        SearchUI 초기화
        
        Args:
            parent: 부모 클래스 (StringSearchPopup)
        """
        self.parent = parent
        self.search_input_frame = None
        self.search_button = None
        self.search_text = tk.StringVar()
        self.match_exact = tk.BooleanVar(value=True)
        
        # 언어 선택 체크박스
        self.selected_languages = {
            "KR": tk.BooleanVar(value=True),
            "EN": tk.BooleanVar(value=False),
            "CN": tk.BooleanVar(value=False),
            "TW": tk.BooleanVar(value=False),
            "STRING_ID": tk.BooleanVar(value=False),
        }
        
        # 현재 모드 추적
        self.current_mode = "search"
        
        # 검색 결과 필터링 관련 변수
        self.column_filter_vars = []
        self._detached_items = []
    
    def create_search_input_frame(self, parent_frame):
        """
        검색 입력 프레임 생성
        
        Args:
            parent_frame: 부모 프레임
        
        Returns:
            생성된 검색 입력 프레임
        """
        # 검색 입력 프레임
        self.search_input_frame = tk.Frame(parent_frame)
        
        tk.Label(self.search_input_frame, text="검색어: ").pack(side="left")
        entry = tk.Entry(self.search_input_frame, textvariable=self.search_text, width=40)
        entry.pack(side="left", padx=5)
        entry.bind("<Return>", lambda event: self.start_search())
        
        tk.Checkbutton(self.search_input_frame, text="완전 일치", variable=self.match_exact).pack(side="left", padx=10)

        # 언어 선택 체크박스 표시
        for lang in ["KR", "EN", "CN", "TW", "STRING_ID"]:
            tk.Checkbutton(self.search_input_frame, text=lang, variable=self.selected_languages[lang]).pack(side="left")
        
        return self.search_input_frame
    
    def create_filter_frame(self, parent_frame):
        """
        필터 입력 프레임 생성
        
        Args:
            parent_frame: 부모 프레임
            
        Returns:
            생성된 필터 프레임과 필터 변수 목록
        """
        # 필터 입력란 + 초기화 버튼
        filter_container = tk.Frame(parent_frame)
        
        filter_row = tk.Frame(filter_container)
        filter_row.pack(side="left")

        self.column_filter_vars = []
        headers = ["파일명", "시트명", "STRING_ID", "KR", "EN", "CN", "TW"]
        for i, header in enumerate(headers):
            var = tk.StringVar()
            entry = tk.Entry(filter_row, textvariable=var, width=15)
            entry.grid(row=0, column=i, padx=1)
            entry.bind("<KeyRelease>", self.filter_by_columns)
            self.column_filter_vars.append(var)

        tk.Button(filter_container, text="🔄 필터 초기화", command=self.clear_column_filters).pack(side="right", padx=10)
        
        return filter_container, self.column_filter_vars
    
    def start_search(self):
        """검색 실행"""
        # UI 부모 객체 접근
        parent = self.parent
        
        # 트리뷰 초기화
        parent.clear_result()
        
        # 키워드 확인
        keyword = self.search_text.get().strip()
        if not keyword:
            parent.show_warning("입력 필요", "검색할 키워드를 입력해주세요.")
            return

        # 선택 언어 수집
        columns = []
        for lang in ["STRING_ID", "KR", "EN", "CN", "TW"]:
            var = self.selected_languages.get(lang)
            if isinstance(var, tk.BooleanVar) and var.get():
                columns.append(lang.lower())

        if not columns:
            parent.show_warning("선택 필요", "검색할 언어(KR, EN 등)를 최소 1개 이상 선택해주세요.")
            return
        
        # DB 검색 옵션 설정
        search_options = {
            "match_exact": self.match_exact.get(),
            "match_case": parent.match_case.get() if hasattr(parent, "match_case") else False,
            "match_word": parent.match_word.get() if hasattr(parent, "match_word") else False,
            "use_regex": parent.use_regex.get() if hasattr(parent, "use_regex") else False
        }
        
        # DB 검색 실행
        results = parent.db_operations.search_string_db(
            keyword=keyword, 
            columns=columns, 
            **search_options
        )

        if not results:
            parent.show_info("검색 결과 없음", "조건에 맞는 결과가 없습니다.", parent=parent.top)
            return

        # 검색 히스토리에 추가
        if keyword not in parent.search_history:
            parent.search_history.insert(0, keyword)
            while len(parent.search_history) > 10:  # 최대 10개 항목 유지
                parent.search_history.pop()
            parent.refresh_history_listbox()
            save_search_history(parent.search_history, "string")  # 히스토리 저장
            
        # 결과 처리 및 트리뷰에 추가
        for row in results:
            # 파일명 처리 (표시용)
            file_path = row["file"]
            display_name = os.path.basename(file_path)  # 파일명만 표시
            
            values = [
                display_name,  # 파일명만 표시
                row["sheet"], 
                row["STRING_ID"],
                row.get("KR", ""), 
                row.get("EN", ""), 
                row.get("CN", ""), 
                row.get("TW", "")
            ]
            
            # 매칭된 컬럼 강조 표시
            matched = row.get("matched", [])
            if "KR" in matched: values[3] = f"{values[3]}"
            if "EN" in matched: values[4] = f"{values[4]}"
            if "CN" in matched: values[5] = f"{values[5]}"
            if "TW" in matched: values[6] = f"{values[6]}"
            
            # 원래 파일 경로를 태그로 저장 (열기 기능에서 사용)
            parent.tree.insert("", "end", values=values, tags=(file_path,))

        # 상태 업데이트
        if hasattr(parent.status_var, "set"):
            parent.status_var.set(f"🔍 총 {len(results)}건 검색됨")
    
    def start_search_thread(self):
        """백그라운드 스레드에서 검색 실행"""
        threading.Thread(target=self.start_search, daemon=True).start()
    
    def stop_search(self):
        """검색 중지 (필요시 구현)"""
        self.parent.stop_flag = True  # 중지 플래그 설정
        
    def clear_column_filters(self):
        """컬럼 필터 초기화"""
        for var in self.column_filter_vars:
            var.set("")
        self.filter_by_columns()  # 필터 적용
    
    def filter_by_columns(self, event=None):
        """컬럼별 필터링 실행"""
        parent = self.parent  # 부모 클래스 참조
        filters = [v.get().strip().lower() for v in self.column_filter_vars]

        # 필터 기준: 트리뷰 전체 항목 (숨겨진 것도 포함)
        all_items = list(parent.tree.get_children("")) + self._detached_items

        new_detached = []

        for item in all_items:
            values = parent.tree.item(item, "values")
            visible = True
            for i, f in enumerate(filters):
                if f and f not in str(values[i]).lower():
                    visible = False
                    break

            if visible:
                parent.tree.reattach(item, "", "end")
            else:
                parent.tree.detach(item)
                new_detached.append(item)

        self._detached_items = new_detached  # 최종적으로 덮어쓰기
    
    def open_excel_from_result(self, event):
        """검색 결과에서 엑셀 파일 열기"""
        parent = self.parent  # 부모 클래스 참조
        
        print("엑셀 파일 열기 시도")
        try:
            item = parent.tree.focus()
            if not item:
                print("선택된 항목 없음")
                return
            
            values = parent.tree.item(item, "values")
            tags = parent.tree.item(item, "tags")  # 태그 정보 가져오기
            print(f"값: {values}")
            
            if len(values) < 3:
                print("값 정보 부족")
                return
            
            file_name = values[0]  # values의 파일명
            sheet_name = values[1]
            
            # 파일 경로 구성 (태그 정보 활용)
            if tags and len(tags) > 0:
                # 태그에서 파일 경로 정보 사용
                file_rel_path = tags[0]
                if os.path.isabs(file_rel_path):
                    # 절대 경로인 경우 그대로 사용
                    file_path = file_rel_path
                else:
                    # 상대 경로인 경우 folder_path와 결합
                    file_path = os.path.join(parent.folder_path, file_rel_path)
            else:
                # 태그 정보가 없는 경우 기존 방식 사용
                file_path = os.path.join(parent.folder_path, file_name)
            
            # 백그라운드에서 엑셀 열기 및 시트 활성화 작업 수행
            parent.file_manager.open_excel_with_sheet(file_path, sheet_name)
            
        except Exception as e:
            error_msg = f"엑셀 파일을 열 수 없습니다: {str(e)}"
            print(f"엑셀 파일 열기 오류: {str(e)}")
            parent.show_error("파일 열기 오류", error_msg)
    
    def clear_detail_frame(self):
        """상세 정보 프레임 초기화"""
        parent = self.parent
        
        if hasattr(parent, 'detail_frame'):
            for widget in parent.detail_frame.winfo_children():
                widget.destroy()
    
    def on_row_selected(self, event):
        """행 선택 시 상세 정보 표시"""
        parent = self.parent
        
        self.clear_detail_frame()
        selected_item = parent.tree.selection()
        if not selected_item:
            return

        item = parent.tree.item(selected_item[0])
        values = item["values"]
        if not values or len(values) < 3:
            return

        file, sheet = values[0], values[1]
        string_id = values[2]
        data_by_lang = dict(zip(["KR", "EN", "CN", "TW"], values[3:]))

        row_frame = tk.Frame(parent.detail_frame)
        row_frame.pack(fill="x", padx=5, pady=5)

        tk.Label(row_frame, text="STRING_ID:", width=15, anchor="w").pack(side="left")
        id_text = tk.Text(row_frame, height=1, width=40)
        id_text.insert("1.0", string_id)
        id_text.pack(side="left", fill="x", expand=True)

        for lang in ["KR", "EN", "CN", "TW"]:
            val = data_by_lang.get(lang, "")
            line = tk.Frame(parent.detail_frame)
            line.pack(fill="x", padx=5, anchor="w")
            
            tk.Label(line, text=f"{lang}:", width=15, anchor="w").pack(side="left")
            text_widget = tk.Text(line, height=max(1, min(5, val.count("\n") + 1)), wrap="word")
            text_widget.insert("1.0", val)
            text_widget.config(state="normal")
            text_widget.pack(side="left", fill="both", expand=True)
            
            if len(val) > 100 or val.count("\n") > 0:
                scrollbar = tk.Scrollbar(line, command=text_widget.yview)
                text_widget.config(yscrollcommand=scrollbar.set)
                scrollbar.pack(side="right", fill="y")