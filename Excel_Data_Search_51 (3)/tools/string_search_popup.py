import os
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import threading
import time

# 리팩토링된 모듈들 임포트
from ui.progress_manager import ProgressManager
from ui.search_ui import SearchUI
from ui.replace_ui import ReplaceUI

# 기존 유틸리티 함수 임포트
from tools.search_settings import SearchFilesManager
from tools.search_options_popup import SearchOptionsPopup
from tools.file_operations import ExcelFileManager
from tools.db_operations import StringDBOperations

# 수정할 코드
from utils.common_utils import PathUtils, show_message
from utils.cache_utils import load_cached_data, hash_paths, update_excel_cache
from utils.config_utils import load_search_history, save_search_history


# 전역 변수 (필요시)
df_cache_global = {}

class StringSearchPopup:
    """문자열 검색 및 치환 기능을 제공하는 메인 클래스"""
    
    def __init__(self, master, folder_path, db_path, excel_cache=None):
        """
        StringSearchPopup 초기화
        
        Args:
            master: 부모 윈도우
            folder_path: 데이터 폴더 경로
            db_path: DB 파일 경로
            excel_cache: 엑셀 캐시 (선택 사항)
        """
        logging.debug("StringSearchPopup 초기화 시작")
        self.top = tk.Toplevel(master)
        self.top.title("🧾 String 검색/치환 도구")
        self.top.geometry("1200x800")
        
        # 경로 및 기본 변수 설정
        self.folder_path = folder_path
        self.db_path = db_path
        self.db_folder_path = tk.StringVar(value=self.folder_path)
        
        # 현재 모드 설정
        self.current_mode = "search"
        self.mode_lock = False
        
        # 검색 옵션 변수
        self.match_case = tk.BooleanVar(value=False)  # 대소문자 구분 없음
        self.use_regex = tk.BooleanVar(value=False)  # 정규식 사용 안함 
        self.match_word = tk.BooleanVar(value=False)  # 단어 단위 검색 안함
        self.keyword_var = tk.StringVar()
        
        # 캐시 처리
        if excel_cache and len(excel_cache) > 0:
            self.cache = excel_cache
        else:
            # 캐시 찾기 시도
            found_cache = self._find_existing_cache()
            self.cache = found_cache if found_cache else {}
            
        self.df_cache = df_cache_global
        
        # 검색 히스토리 및 기타 변수 초기화
        self.search_history = load_search_history("string")
        self.stop_flag = False
        self.file_to_items = {}  # 파일-항목 매핑
        
        # 상태 표시용 변수
        self.status_var = tk.StringVar(value="초기화 중...")
        
        # 모듈 초기화
        self.init_modules()
        
        # UI 구성
        self.build_ui()
        self.refresh_history_listbox()
        
        # 백그라운드에서 파일 매니저 초기화
        threading.Thread(target=self._async_initialize, daemon=True).start()
    
    def init_modules(self):
        """모듈 초기화"""
        # 파일 작업 관리자
        self.file_manager = ExcelFileManager(self.folder_path, self._update_progress)
        
        # DB 작업 관리자
        self.db_operations = StringDBOperations(self.folder_path, self.db_path, self._update_progress)
        
        # 진행 상태 관리자
        self.progress_manager = ProgressManager(self.top)
        
        # 검색 UI 모듈
        self.search_ui = SearchUI(self)
        
        # 치환 UI 모듈
        self.replace_ui = ReplaceUI(self)
        
        # 공통 콜백 설정
        self._set_common_callbacks()
    
    def _set_common_callbacks(self):
        """모듈 간 콜백 함수 설정"""
        # 각 모듈에 필요한 콜백 함수 등록
        pass
    
    def _find_existing_cache(self):
        """사용 가능한 캐시 찾기"""
        cache_base = ".cache"
        if os.path.exists(cache_base):
            for d in os.listdir(cache_base):
                cache_file = os.path.join(cache_base, d, "excel_cache.json")
                if os.path.exists(cache_file):
                    return load_cached_data(cache_file)
        return None
    
    def _async_initialize(self):
        """백그라운드에서 캐시 및 파일 매니저 초기화"""
        try:
            # 캐시가 비어있으면 새로 생성
            if not self.cache:
                self.top.after(0, lambda: self.status_var.set("엑셀 캐시 생성 중..."))
                self.cache = update_excel_cache(self.folder_path, self.db_path)
            
            # 파일 매니저 초기화
            self.top.after(0, lambda: self.status_var.set("검색 파일 목록 로드 중..."))
            self.files_manager = SearchFilesManager(self.folder_path)
            
            # 파일 목록 검증
            if self.cache:
                self.files_manager.validate_and_update_files(self.cache)
            
            # 완료 표시
            self.top.after(0, lambda: self.status_var.set("검색 준비 완료"))
            
        except Exception as e:
            error_msg = f"초기화 오류: {str(e)}"
            self.top.after(0, lambda: self.status_var.set(f"오류: {error_msg}"))
    
    def build_ui(self):
        """전체 UI 구성"""
        # 리본 UI 구성
        self.configure_ribbon_ui()
        
        # 검색 및 치환 입력 프레임 생성 (아직 배치하지 않음)
        self.search_input_frame = self.search_ui.create_search_input_frame(self.top)
        self.replace_input_frame = self.replace_ui.create_replace_input_frame(self.top)
        
        # 🔹 히스토리 제목 + 삭제 버튼
        title_frame = tk.Frame(self.top)
        title_frame.pack(fill="x", padx=10, pady=(5, 0))

        tk.Label(title_frame, text="🔍 최근 검색어").pack(side="left")
        clear_btn = tk.Button(title_frame, text="❌ 히스토리 삭제", command=self.clear_history)
        clear_btn.pack(side="left", padx=5)

        # 🔹 히스토리 리스트
        history_frame = tk.Frame(self.top)
        history_frame.pack(fill="x", padx=10, pady=(0, 5))

        self.history_listbox = tk.Listbox(history_frame, height=5)
        self.history_listbox.pack(fill="x")
        self.history_listbox.bind("<<ListboxSelect>>", self.research_from_history)

        # 창 포커스 이벤트 바인딩 추가
        self.top.bind("<FocusIn>", self.on_window_focus)
        self.top.bind("<FocusOut>", self.on_window_focus_out)

        # 필터 컨테이너
        filter_container = tk.Frame(self.top)
        filter_container.pack(fill="x", padx=10, pady=(5, 10))
        
        # 필터 UI 생성 (SearchUI 모듈에서)
        filter_frame, self.column_filter_vars = self.search_ui.create_filter_frame(filter_container)
        filter_frame.pack(fill="x", expand=True)

        # 트리뷰 프레임 (검색 결과)
        result_frame = tk.Frame(self.top)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree_scroll = tk.Scrollbar(result_frame)
        self.tree_scroll.pack(side="right", fill="y")

        # 초기 트리뷰는 검색 모드로 설정
        all_columns = ["KR", "EN", "CN", "TW"]
        self.tree = ttk.Treeview(result_frame, columns=("파일명", "시트명", "STRING_ID", *all_columns), show="headings", yscrollcommand=self.tree_scroll.set)

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=160 if col != "STRING_ID" else 200)

        self.tree.pack(fill="both", expand=True)
        self.tree_scroll.config(command=self.tree.yview)
        self.tree.bind("<<TreeviewSelect>>", self.search_ui.on_row_selected)
        self.tree.bind("<Double-1>", self.search_ui.open_excel_from_result)
        self.tree.bind("<Button-1>", self.replace_ui.handle_click)  # 체크박스 클릭 처리

        # 🔹 상세 보기 패널
        detail_outer = tk.Frame(self.top)
        detail_outer.pack(fill="both", padx=10, pady=5)

        self.detail_canvas = tk.Canvas(detail_outer, height=140)
        self.detail_scrollbar = tk.Scrollbar(detail_outer, orient="vertical", command=self.detail_canvas.yview)
        self.detail_frame = tk.Frame(self.detail_canvas)

        self.detail_frame.bind("<Configure>", lambda e: self.detail_canvas.configure(scrollregion=self.detail_canvas.bbox("all")))
        self.detail_canvas.create_window((0, 0), window=self.detail_frame, anchor="nw")
        self.detail_canvas.configure(yscrollcommand=self.detail_scrollbar.set)

        self.detail_canvas.pack(side="left", fill="both", expand=True)
        self.detail_scrollbar.pack(side="right", fill="y")

        # 🔹 하단 상태 표시
        tk.Label(self.top, textvariable=self.status_var).pack(fill="x")
        
        # 초기 상태 설정: 검색 모드
        self._show_search_controls()
        self._hide_replace_controls()
    
    def configure_ribbon_ui(self):
        """리본 UI 구성"""
        # 리본 컨테이너
        self.ribbon_frame = tk.Frame(self.top, relief="ridge", bd=1)
        self.ribbon_frame.pack(fill="x", padx=5, pady=5)
        
        # 1. 모드 섹션
        mode_section = tk.LabelFrame(self.ribbon_frame, text="모드")
        mode_section.pack(side="left", padx=5, pady=5, fill="y")
        
        self.search_mode_btn = tk.Button(mode_section, text="🔍 검색", width=10, 
                                        relief="sunken", bg="#e6f0ff",
                                        command=lambda: self.switch_to_search_mode(force=True))
        self.search_mode_btn.pack(side="left", padx=5, pady=2)
        
        self.replace_mode_btn = tk.Button(mode_section, text="🔄 치환", width=10,
                                        command=lambda: self.switch_to_replace_mode(force=True))
        self.replace_mode_btn.pack(side="left", padx=5, pady=2)
        
        # 2. 액션 섹션 (모드에 따라 동적 변경)
        self.action_section = tk.LabelFrame(self.ribbon_frame, text="액션")
        self.action_section.pack(side="left", padx=5, pady=5, fill="y")
        
        # 검색 모드 액션 버튼
        self.search_action_frame = tk.Frame(self.action_section)
        self.search_action_frame.pack(fill="both", expand=True)
        
        tk.Button(self.search_action_frame, text="🔍 검색", width=8,
                command=self.search_ui.start_search_thread).pack(side="left", padx=5, pady=2)
        tk.Button(self.search_action_frame, text="🛑 중지", width=8,
                command=self.search_ui.stop_search).pack(side="left", padx=5, pady=2)
        
        # 치환 모드 액션 버튼
        self.replace_action_frame = tk.Frame(self.action_section)
        # pack은 모드 전환 시 처리
        
        # 치환 모드 선택
        mode_box = tk.Frame(self.replace_action_frame)
        mode_box.pack(side="left", padx=5)
        
        tk.Radiobutton(mode_box, text="단어 치환", variable=self.replace_ui.replace_mode, value="replace", 
                    command=self.replace_ui.update_replace_ui).pack(side="top", anchor="w")
        tk.Radiobutton(mode_box, text="일괄 변경", variable=self.replace_ui.replace_mode, value="bulk", 
                    command=self.replace_ui.update_replace_ui).pack(side="top", anchor="w")
        tk.Radiobutton(mode_box, text="고유값 적용", variable=self.replace_ui.replace_mode, value="unique", 
                    command=self.replace_ui.update_replace_ui).pack(side="top", anchor="w")
        
        # 치환 실행 버튼
        tk.Button(self.replace_action_frame, text="✔️ 선택 항목 처리", width=15,
                command=self.replace_ui.process_selected).pack(side="left", padx=5, pady=2)
        
        # 3. 옵션 섹션
        option_section = tk.LabelFrame(self.ribbon_frame, text="옵션")
        option_section.pack(side="left", padx=5, pady=5, fill="y")
        
        # 공통 옵션
        tk.Checkbutton(option_section, text="대소문자 구분", variable=self.match_case).pack(side="left", padx=5)
        tk.Checkbutton(option_section, text="단어 단위 검색", variable=self.match_word).pack(side="left", padx=5)
        tk.Checkbutton(option_section, text="정규식 사용", variable=self.use_regex).pack(side="left", padx=5)
        
        # 4. 도구 섹션
        tools_section = tk.LabelFrame(self.ribbon_frame, text="도구")
        tools_section.pack(side="left", padx=5, pady=5, fill="y")
        
        # 검색 옵션 버튼
        tk.Button(tools_section, text="⚙️ 검색 옵션", 
                command=self.open_search_options).pack(side="left", padx=5, pady=2)
    
    def on_window_focus(self, event):
        """창이 포커스를 받았을 때 처리"""
        # 포커스를 받을 때 모드 잠금 해제
        self.mode_lock = False

    def on_window_focus_out(self, event):
        """창이 포커스를 잃었을 때 처리"""
        # 포커스를 잃을 때 현재 모드 잠금
        self.mode_lock = True
    
    def switch_to_search_mode(self, force=False):
        """검색 모드로 전환"""
        # 모드 잠금 상태이고, 강제 전환이 아니라면 무시
        if self.mode_lock and not force:
            return
            
        if self.current_mode == "search":
            return  # 이미 검색 모드면 변경 안함
            
        self.current_mode = "search"
        
        # 버튼 스타일 변경
        self.search_mode_btn.config(relief="sunken", bg="#e6f0ff")
        self.replace_mode_btn.config(relief="raised", bg="SystemButtonFace")
        
        # UI 업데이트
        self._show_search_controls()
        self._hide_replace_controls()
        
        # 체크박스 상태 초기화 (이제 필요 없음)
        self.replace_ui.row_checks = {}
        
        # ★ 트리뷰 컬럼 재구성 전에 데이터 완전 초기화
        self.tree.delete(*self.tree.get_children())
        
        # 트리뷰 설정 변경
        self.configure_treeview_for_mode("search")
        
        # ★ 현재 검색어로 검색 실행 (트리뷰 다시 채우기)
        keyword = self.search_ui.search_text.get().strip()
        if keyword:
            self.search_ui.start_search()

    def switch_to_replace_mode(self, force=False):
        """치환 모드로 전환"""
        # 모드 잠금 상태이고, 강제 전환이 아니라면 무시
        if self.mode_lock and not force:
            return
        
        # 검색 결과가 있는지 확인
        if not self.tree.get_children():
            # 메인 창을 부모로 지정하여 메시지 박스 표시
            self.show_warning("데이터 없음", "치환할 검색 결과가 없습니다.")
            # 메시지 박스 처리 후 메인 창에 포커스 다시 맞추기
            self.top.focus_force()
            return
            
        if self.current_mode == "replace":
            return  # 이미 치환 모드면 변경 안함
            
        self.current_mode = "replace"
        
        # 검색어를 치환어 기본값으로 설정
        if not self.replace_ui.replace_from.get():
            self.replace_ui.replace_from.set(self.search_ui.search_text.get())
        
        # 버튼 스타일 변경
        self.search_mode_btn.config(relief="raised", bg="SystemButtonFace")
        self.replace_mode_btn.config(relief="sunken", bg="#e6f0ff")
        
        # UI 업데이트
        self._hide_search_controls()
        self._show_replace_controls()
        
        # 체크박스 상태 초기화
        self.replace_ui.row_checks = {}
        self.replace_ui.select_all_var.set(True)
        
        # 트리뷰 설정 변경
        self.configure_treeview_for_mode("replace")
        
        # 모든 행 체크 상태로 설정
        for item in self.tree.get_children():
            self.replace_ui.row_checks[item] = True
    
    def _show_search_controls(self):
        """검색 관련 컨트롤 표시"""
        self.search_input_frame.pack(fill="x", padx=10, pady=5, after=self.ribbon_frame)
        self.search_action_frame.pack(fill="both", expand=True)

    def _hide_search_controls(self):
        """검색 관련 컨트롤 숨김"""
        self.search_input_frame.pack_forget()
        self.search_action_frame.pack_forget()

    def _show_replace_controls(self):
        """치환 관련 컨트롤 표시"""
        self.replace_input_frame.pack(fill="x", padx=10, pady=5, after=self.ribbon_frame)
        self.replace_action_frame.pack(fill="both", expand=True)
        
        # 현재 치환 모드에 따라 적절한 프레임 표시
        self.replace_ui.update_replace_ui()

    def _hide_replace_controls(self):
        """치환 관련 컨트롤 숨김"""
        self.replace_input_frame.pack_forget()
        self.replace_action_frame.pack_forget()
    
    def configure_treeview_for_mode(self, mode):
        """모드에 따라 트리뷰 설정 변경"""
        # 현재 트리뷰 내용 저장
        current_items = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            tags = self.tree.item(item, "tags")
            current_items.append((values, tags))
        
        # 트리뷰 초기화
        self.tree.delete(*self.tree.get_children())
        
        # 모드에 따라 컬럼 구성 변경
        if mode == "search":
            columns = ("파일명", "시트명", "STRING_ID", "KR", "EN", "CN", "TW", "JP")
        else:  # replace 모드
            columns = ("선택", "파일명", "시트명", "STRING_ID", "KR", "EN", "CN", "TW", "JP")
        
        # 트리뷰 재구성
        self.tree["columns"] = columns
        
        for col in columns:
            if col == "선택":
                self.tree.column(col, width=50, anchor="center")
                self.tree.heading(col, text=col)
            else:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=160 if col != "STRING_ID" else 200)
        
        # 저장된 아이템 다시 추가
        for values, tags in current_items:
            if mode == "search":
                # 치환 모드 -> 검색 모드
                if len(values) > 0 and len(columns) == 8:  # 검색 모드용 컬럼 개수 확인
                    if len(values) > 8:  # 체크박스 열이 포함된 경우
                        # 체크박스 열 제거 (첫 번째 열)
                        new_values = values[1:9]  # 검색 모드에 필요한 8개 컬럼만 사용
                        self.tree.insert("", "end", values=new_values, tags=tags)
                    else:
                        # 열 개수가 정확히 맞지 않으면 원본 데이터 사용
                        self.tree.insert("", "end", values=values[:8], tags=tags)
            else:  # replace 모드
                # 검색 모드 -> 치환 모드
                if len(values) > 0 and len(columns) == 9:  # 치환 모드용 컬럼 개수 확인
                    if len(values) <= 8:  # 체크박스 열이 없는 경우
                        # 맨 앞에 체크박스 열 추가
                        new_values = ["✓"] + list(values)
                        # 체크박스 상태 추적을 위한 item_id 저장
                        item_id = self.tree.insert("", "end", values=new_values, tags=tags)
                        self.replace_ui.row_checks[item_id] = True
                    else:
                        # 이미 체크박스 열이 있으면 그대로 사용
                        self.tree.insert("", "end", values=values, tags=tags)
    
    def clear_result(self):
        """트리뷰 초기화 및 관련 변수 재설정"""
        if hasattr(self, "tree"):
            # 트리뷰 항목 모두 삭제
            self.tree.delete(*self.tree.get_children())
            # 필터링 관련 변수도 초기화
            self.search_ui._detached_items = []
            # 검색 모드인 경우 체크박스 상태도 초기화
            if self.current_mode == "replace":
                self.replace_ui.row_checks = {}
    
    def clear_history(self):
        """검색 히스토리 초기화"""
        self.search_history.clear()
        self.refresh_history_listbox()
        save_search_history(self.search_history, "string")  # 실제 저장
    
    def refresh_history_listbox(self):
        """히스토리 리스트박스 새로고침"""
        self.history_listbox.delete(0, tk.END)
        for keyword in self.search_history:
            self.history_listbox.insert(tk.END, keyword)
    
    def research_from_history(self, event):
        """히스토리에서 검색어 선택 시 실행"""
        selection = event.widget.curselection()
        if selection:
            keyword = event.widget.get(selection[0])
            self.search_ui.search_text.set(keyword)
            # 자동 검색 실행 (선택 사항)
            self.search_ui.start_search()
    
    def open_search_options(self):
        """검색 옵션 팝업 열기"""
        SearchOptionsPopup(self.top, self.folder_path, self.files_manager, self.cache)
    
    def _update_progress(self, message, success=True):
        """진행 상태 업데이트 (진행 창이 있을 때만)"""
        if self.progress_manager.is_active():
            self.progress_manager.update_progress(message, success)
    
    # 메시지 박스 도우미 함수들
    def show_info(self, title, message, parent=None):
        """정보 메시지 표시"""
        return show_message(parent or self.top, "info", title, message, log_level="info")

    def show_warning(self, title, message, parent=None):
        """경고 메시지 표시"""
        return show_message(parent or self.top, "warning", title, message, log_level="warning")

    def show_error(self, title, message, parent=None):
        """오류 메시지 표시"""
        return show_message(parent or self.top, "error", title, message, log_level="error")

    def show_confirm(self, title, message, parent=None):
        """확인/취소 메시지 표시, 사용자 선택 결과 반환"""
        return show_message(parent or self.top, "yesno", title, message, log_level="info")