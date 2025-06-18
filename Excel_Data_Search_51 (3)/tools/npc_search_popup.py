# npc_search_popup.py 파일

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, PanedWindow
import pandas as pd
import sqlite3
import threading
import re  # 정규식 처리용

from utils.cache_utils import load_cached_data, hash_paths, update_excel_cache
from utils.excel_utils import ExcelFileManager
from utils.config_utils import load_search_history, save_search_history
from utils.common_utils import logger, PathUtils, FileUtils
from ui.common_components import show_message
from utils.type_mappings import get_table_name_for_type, get_description_for_type, resolve_type_info

class NPCListPopup:
    """전체 NPC 목록을 표시하는 팝업 클래스"""
    def __init__(self, master, folder, db_folder, typecode_mapping, excel_cache):
        self.folder = folder
        self.db_folder = db_folder
        self.typecode_mapping = typecode_mapping
        self.cache = excel_cache
        self.top = Toplevel(master)
        self.top.title("👤 전체 NPC 목록")
        self.top.geometry("1400x700")
        self._detached_items = []  # 분리된 항목을 저장할 리스트 추가
        self.current_category = 30  # 기본 카테고리
        self.current_page = 0       # 현재 페이지
        self.items_per_page = 1000  # 페이지당 항목 수
        self.total_pages = 0        # 총 페이지 수 (로드 후 계산)
        self._is_window_closed = False  # 창 닫힘 상태 추적
        self.top.protocol("WM_DELETE_WINDOW", self._on_window_close)
        
        self._build_ui()
        self._load_all_npcs()


    def _on_window_close(self):
        """창 닫힘 이벤트 처리"""
        self._is_window_closed = True
        self.top.destroy()
        
    def _is_window_valid(self):
        """창이 유효한지 확인"""
        try:
            return not self._is_window_closed and self.top.winfo_exists()
        except tk.TclError:
            return False
        
    def _safe_ui_update(self, callback):
        """안전한 UI 업데이트 - 창이 유효할 때만 실행"""
        if self._is_window_valid():
            try:
                self.top.after(0, callback)
            except tk.TclError:
                # 창이 이미 닫힌 경우 무시
                pass

    def _build_ui(self):
        """UI를 구성합니다."""
        # 상단 프레임
        top_frame = tk.Frame(self.top)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(top_frame, text="전체 NPC 목록", font=("Helvetica", 12, "bold")).pack(side="left")
    
        # 카테고리 필터 프레임 추가
        category_frame = tk.Frame(self.top)
        category_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(category_frame, text="카테고리 필터:").pack(side="left")
        
        # 카테고리 필터 변수
        self.category_var = tk.StringVar(value="NPC")  # 기본값을 NPC로 설정
        
        # 카테고리 라디오 버튼
        categories = [
            ("전체", "ALL"),
            ("영웅", "HERO"),
            ("몬스터", "MONSTER"),
            ("NPC", "NPC"),
            ("기타", "OTHER")
        ]
        for text, value in categories:
            rb = tk.Radiobutton(category_frame, text=text, value=value, 
                            variable=self.category_var, command=self._apply_category_filter)
            rb.pack(side="left", padx=5)

        # 카테고리 탭 추가
        category_frame = tk.Frame(self.top)
        category_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(category_frame, text="NPC 카테고리:").pack(side="left")
        
        # 카테고리 탭 버튼
        self.cat_30_btn = tk.Button(category_frame, text="일반 NPC (30)", 
                                    command=lambda: self._load_category(30))
        self.cat_30_btn.pack(side="left", padx=5)
        
        self.cat_31_btn = tk.Button(category_frame, text="특수 NPC (31)", 
                                    command=lambda: self._load_category(31))
        self.cat_31_btn.pack(side="left", padx=5)
        
        self.cat_32_btn = tk.Button(category_frame, text="기타 NPC (32)", 
                                    command=lambda: self._load_category(32))
        self.cat_32_btn.pack(side="left", padx=5)
        
        # 초기 로드 상태 표시
        self.page_info = tk.Label(category_frame, text="")
        self.page_info.pack(side="right", padx=10)
        
        # 페이지 이동 버튼
        self.next_page_btn = tk.Button(category_frame, text="다음 페이지 →", 
                                    command=self._load_next_page)
        self.next_page_btn.pack(side="right", padx=5)
        
        self.prev_page_btn = tk.Button(category_frame, text="← 이전 페이지", 
                                    command=self._load_prev_page, state=tk.DISABLED)
        self.prev_page_btn.pack(side="right", padx=5)

        # EventID 검색 추가
        search_frame = tk.Frame(self.top)
        search_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(search_frame, text="EventID:").pack(side="left")
        self.event_id_entry = tk.Entry(search_frame)
        self.event_id_entry.pack(side="left", padx=5)

        search_event_btn = tk.Button(search_frame, text="EventID 검색", 
                                    command=self._search_by_event_id)
        search_event_btn.pack(side="left", padx=5)

        # 이름 검색 부분 추가
        tk.Label(search_frame, text="이름:").pack(side="left", padx=(15, 0))
        self.name_entry = tk.Entry(search_frame)
        self.name_entry.pack(side="left", padx=5)

        search_name_btn = tk.Button(search_frame, text="이름 검색", 
                                command=self._search_by_name)
        search_name_btn.pack(side="left", padx=5)

        # 전체 목록으로 돌아가기 버튼 추가
        reset_search_btn = tk.Button(search_frame, text="전체 목록 보기", 
                                command=self._reset_search)
        reset_search_btn.pack(side="left", padx=5)

        # 필터 프레임
        filter_frame = tk.Frame(top_frame)
        filter_frame.pack(side="right")
        
        # "# 제외" 체크박스 추가
        self.exclude_hash_var = tk.BooleanVar(value=False)
        exclude_hash_check = tk.Checkbutton(filter_frame, text="'#'로 시작하는 이름 제외", 
                                            variable=self.exclude_hash_var,
                                            command=self._apply_name_filter)
        exclude_hash_check.pack(side="left", padx=10)
        
        # 기존 필터 요소들
        tk.Label(filter_frame, text="필터:").pack(side="left")
        self.filter_var = tk.StringVar()
        filter_entry = tk.Entry(filter_frame, textvariable=self.filter_var, width=15)
        filter_entry.pack(side="left", padx=5)
        
        filter_btn = tk.Button(filter_frame, text="적용", command=self._apply_filter)
        filter_btn.pack(side="left", padx=2)
        
        clear_btn = tk.Button(filter_frame, text="초기화", command=self._clear_filter)
        clear_btn.pack(side="left", padx=2)
        
        # PanedWindow로 좌/우 분할
        self.paned = PanedWindow(self.top, orient=tk.HORIZONTAL)
        self.paned.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 좌측: NPC 목록
        left_frame = tk.Frame(self.paned)
        self.paned.add(left_frame, width=600)
        
        # 버튼 프레임 추가
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill="x", pady=5)
        
        tk.Label(left_frame, text="NPC 목록", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
        # 숨김 버튼 추가
        self.hide_btn = tk.Button(btn_frame, text="선택한 NPC 숨기기", command=self._add_hash_to_selected)
        self.hide_btn.pack(side="left", padx=5)

        # 숨김 해제 버튼 추가
        self.unhide_btn = tk.Button(btn_frame, text="선택 NPC 숨김 해제", command=self._remove_hash_from_selected)
        self.unhide_btn.pack(side="left", padx=5)
        
        # 목록 프레임 (스크롤바 포함)
        list_frame = tk.Frame(left_frame)
        list_frame.pack(fill="both", expand=True)
        
        # 수직 스크롤바
        list_scroll_y = tk.Scrollbar(list_frame)
        list_scroll_y.pack(side="right", fill="y")
        
        # 트리뷰로 구현 - NPC 정보 출력용 컬럼 설정
        columns = ["Status", "ID", "Name", "Type", "EventID"]
        self.npc_tree = ttk.Treeview(list_frame, columns=columns, show="headings", 
                                    yscrollcommand=list_scroll_y.set)
        
        # 컬럼 설정
        self.npc_tree.heading("Status", text="상태")
        self.npc_tree.column("Status", width=80, anchor="center")
        self.npc_tree.heading("ID", text="ID")
        self.npc_tree.column("ID", width=80, anchor="w")
        self.npc_tree.heading("Name", text="이름")
        self.npc_tree.column("Name", width=150, anchor="w")
        self.npc_tree.heading("Type", text="타입")
        self.npc_tree.column("Type", width=100, anchor="w")
        self.npc_tree.heading("EventID", text="EventID")
        self.npc_tree.column("EventID", width=80, anchor="center")
        
        self.npc_tree.pack(fill="both", expand=True)
        list_scroll_y.config(command=self.npc_tree.yview)
        
        # 우측: 상세 정보
        right_frame = tk.Frame(self.paned)
        self.paned.add(right_frame, width=800)
        
        # 탭 컨트롤 추가 - 기본 정보, 이벤트 방향, 스폰 정보 등을 탭으로 구분
        self.tab_control = ttk.Notebook(right_frame)
        self.tab_control.pack(fill="both", expand=True)
        
        # 탭1: 기본 정보
        self.basic_tab = tk.Frame(self.tab_control)
        self.tab_control.add(self.basic_tab, text="기본 정보")
        
        # 기본 정보 프레임
        basic_frame = tk.Frame(self.basic_tab)
        basic_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 기본 정보 표시 (레이블-값 형태로 구성)
        self.info_frame = tk.Frame(basic_frame)
        self.info_frame.pack(fill="x", padx=5, pady=5)
        
        # NPC 기본 정보 레이블들
        self.info_labels = {}
        info_fields = [
            ("ID", "id_value"), 
            ("이름", "name_value"),
            ("별명", "nickname_value"),
            ("타입", "type_value"),
            ("크기X", "sizex_value"),
            ("크기Y", "sizey_value"),
            ("모델ID", "modelid_value"),
            ("EventID", "eventid_value")
        ]
        
        for i, (label_text, value_key) in enumerate(info_fields):
            # 레이블과 값 쌍 생성
            row_frame = tk.Frame(self.info_frame)
            row_frame.pack(fill="x", pady=2)
            
            tk.Label(row_frame, text=f"{label_text}:", width=10, anchor="e").pack(side="left", padx=5)
            value_label = tk.Label(row_frame, text="", width=30, anchor="w", bg="#f0f0f0", relief="sunken", padx=5)
            value_label.pack(side="left", padx=5, fill="x", expand=True)
            
            self.info_labels[value_key] = value_label
        
        # 탭2: 이벤트 정보
        self.event_tab = tk.Frame(self.tab_control)
        self.tab_control.add(self.event_tab, text="이벤트 정보")

        # 상하 분할 프레임 생성
        event_paned = tk.PanedWindow(self.event_tab, orient=tk.VERTICAL)
        event_paned.pack(fill="both", expand=True, padx=5, pady=5)

        # 상단: 기존 이벤트 정보
        event_top_frame = tk.Frame(event_paned)
        event_paned.add(event_top_frame, height=350)

        tk.Label(event_top_frame, text="이벤트 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=2)

        # 수직 스크롤바
        event_scroll_y = tk.Scrollbar(event_top_frame)
        event_scroll_y.pack(side="right", fill="y")

        # 이벤트 정보 트리뷰
        columns = ["GroupID", "DialogueGroupID", "RequireType", "RequireOption", 
                "HideType", "HideOption", "EndAction", "EndOption1", "EndOption2", "EndOption3"]
        self.event_tree = ttk.Treeview(event_top_frame, columns=columns, show="headings", 
                                    yscrollcommand=event_scroll_y.set)

        # 컬럼 설정
        for col in columns:
            self.event_tree.heading(col, text=col)
            self.event_tree.column(col, width=80, anchor="center")

        self.event_tree.pack(fill="both", expand=True)
        event_scroll_y.config(command=self.event_tree.yview)

        # 이벤트 트리뷰 선택 이벤트 연결
        self.event_tree.bind("<<TreeviewSelect>>", self._on_event_condition_select)

        # 하단: Condition 정보
        event_bottom_frame = tk.Frame(event_paned)
        event_paned.add(event_bottom_frame, height=200)

        tk.Label(event_bottom_frame, text="Condition/Quest 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=2)

        # 수직 스크롤바
        condition_event_scroll_y = tk.Scrollbar(event_bottom_frame)
        condition_event_scroll_y.pack(side="right", fill="y")

        # Condition 정보 트리뷰
        condition_columns = ["Type", "TemplateID", "ConditionType", "Condition1", "Condition2", "QuestName"]
        self.event_condition_tree = ttk.Treeview(event_bottom_frame, columns=condition_columns, 
                                                show="headings", yscrollcommand=condition_event_scroll_y.set)

        # 컬럼 설정
        # 컬럼 설정 (이벤트 Condition 트리뷰)
        for col in condition_columns:
            self.event_condition_tree.heading(col, text=col)
            if col == "QuestName":
                self.event_condition_tree.column(col, width=150, anchor="w")
            elif col == "TemplateID":
                self.event_condition_tree.column(col, width=120, anchor="w")  # 넓게 설정
            else:
                self.event_condition_tree.column(col, width=80, anchor="center")

        self.event_condition_tree.pack(fill="both", expand=True)
        condition_event_scroll_y.config(command=self.event_condition_tree.yview)
        
        # 탭3: 스폰 정보
        self.spawn_tab = tk.Frame(self.tab_control)
        self.tab_control.add(self.spawn_tab, text="스폰 정보")

        # 상하 분할 프레임 생성
        spawn_paned = tk.PanedWindow(self.spawn_tab, orient=tk.VERTICAL)
        spawn_paned.pack(fill="both", expand=True, padx=5, pady=5)

        # 상단: 기존 스폰 정보
        spawn_top_frame = tk.Frame(spawn_paned)
        spawn_paned.add(spawn_top_frame, height=350)

        tk.Label(spawn_top_frame, text="스폰 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=2)

        # 수직 스크롤바
        spawn_scroll_y = tk.Scrollbar(spawn_top_frame)
        spawn_scroll_y.pack(side="right", fill="y")

        # 스폰 정보 트리뷰
        columns = ["MapID", "MapSpawnGroupID", "HelperName", "ShowConditionTID", "HideConditionTID"]
        self.spawn_tree = ttk.Treeview(spawn_top_frame, columns=columns, show="headings", 
                                    yscrollcommand=spawn_scroll_y.set)

        # 컬럼 설정
        for col in columns:
            self.spawn_tree.heading(col, text=col)
            self.spawn_tree.column(col, width=100, anchor="center")

        self.spawn_tree.pack(fill="both", expand=True)
        spawn_scroll_y.config(command=self.spawn_tree.yview)

        # 스폰 트리뷰 선택 이벤트 연결
        self.spawn_tree.bind("<<TreeviewSelect>>", self._on_spawn_condition_select)

        # 하단: Condition 정보
        spawn_bottom_frame = tk.Frame(spawn_paned)
        spawn_paned.add(spawn_bottom_frame, height=200)

        tk.Label(spawn_bottom_frame, text="Condition 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=2)

        # 수직 스크롤바
        condition_spawn_scroll_y = tk.Scrollbar(spawn_bottom_frame)
        condition_spawn_scroll_y.pack(side="right", fill="y")

        # Condition 정보 트리뷰
        condition_columns = ["TemplateID", "ConditionType", "Condition1", "Condition2"]
        self.spawn_condition_tree = ttk.Treeview(spawn_bottom_frame, columns=condition_columns, 
                                                show="headings", yscrollcommand=condition_spawn_scroll_y.set)

        # 컬럼 설정
        for col in condition_columns:
            self.spawn_condition_tree.heading(col, text=col)
            self.spawn_condition_tree.column(col, width=100, anchor="center")

        self.spawn_condition_tree.pack(fill="both", expand=True)
        condition_spawn_scroll_y.config(command=self.spawn_condition_tree.yview)

        # 상태 표시
        self.status_label = tk.Label(self.top, text="")
        self.status_label.pack(anchor="w", padx=10, pady=5)
        
        # 이벤트 연결
        self.npc_tree.bind("<<TreeviewSelect>>", self._on_npc_select)


        # 탭4: 퀘스트 연결 정보 (새로 추가)
        self.quest_tab = tk.Frame(self.tab_control)
        self.tab_control.add(self.quest_tab, text="퀘스트 연결 정보")

        # 퀘스트 연결 정보 트리뷰
        quest_frame = tk.Frame(self.quest_tab)
        quest_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # 수직 스크롤바
        quest_scroll_y = tk.Scrollbar(quest_frame)
        quest_scroll_y.pack(side="right", fill="y")

        # 퀘스트 정보 트리뷰
        columns = ["TemplateID", "GroupID", "QuestType", "QuestName", "RewardGroupID"]
        self.quest_tree = ttk.Treeview(quest_frame, columns=columns, show="headings", 
                                    yscrollcommand=quest_scroll_y.set)

        # 컬럼 설정
        for col in columns:
            self.quest_tree.heading(col, text=col)
            if col == "QuestName":
                self.quest_tree.column(col, width=200, anchor="w")  # 이름은 더 넓게
            else:
                self.quest_tree.column(col, width=100, anchor="center")

        self.quest_tree.pack(fill="both", expand=True)
        quest_scroll_y.config(command=self.quest_tree.yview)


        # 이벤트 탭에 수동 갱신 버튼 추가
        manual_refresh_frame = tk.Frame(event_top_frame)
        manual_refresh_frame.pack(fill="x", pady=2)
        manual_refresh_btn = tk.Button(manual_refresh_frame, text="조건 정보 갱신", 
                                    command=lambda: self._on_event_condition_select(None))
        manual_refresh_btn.pack(side="right", padx=5)



    def _load_category(self, category_type):
        """특정 카테고리의 NPC를 로드합니다."""
        self.current_category = category_type
        self.current_page = 0  # 페이지 초기화
        
        # 기존 데이터 초기화
        if self._is_window_valid():
            self.npc_tree.delete(*self.npc_tree.get_children())
        
        # 상태 표시
        self._safe_ui_update(lambda: self.status_label.config(text=f"🔍 CategoryType {category_type} NPC 로딩 중..."))
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=self._load_category_thread, 
                        args=(category_type,), daemon=True).start()

    
    def _load_category_thread(self, category_type):
        """백그라운드 스레드에서 특정 카테고리의 NPC를 로드합니다."""
        try:
            all_npcs = []
            
            # 현재 페이지의 데이터만 로드
            offset = self.current_page * self.items_per_page
            items_loaded = self._load_from_db(all_npcs, category_type, 
                                            self.items_per_page, offset)
            
            # 총 항목 수 확인 (총 페이지 계산용)
            conn = sqlite3.connect(os.path.join(self.db_folder, "HeroTemplate.db"))
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM HeroTemplate WHERE CategoryType = ?", 
                        (category_type,))
            total_items = cursor.fetchone()[0]
            conn.close()
            
            # 총 페이지 수 계산
            self.total_pages = (total_items + self.items_per_page - 1) // self.items_per_page
            
            # 트리뷰에 데이터 추가
            for idx, npc in enumerate(all_npcs):
                self._safe_ui_update(lambda idx=idx, values=npc: 
                            self.npc_tree.insert("", "end", iid=f"npc_{idx}", values=values))
            
            # 페이지 정보 업데이트
            page_info = f"페이지 {self.current_page + 1}/{self.total_pages} (총 {total_items}개)"
            self._safe_ui_update(lambda: self.page_info.config(text=page_info))
            
            # 페이지 버튼 상태 업데이트
            self._safe_ui_update(lambda: self.prev_page_btn.config(
                state=tk.NORMAL if self.current_page > 0 else tk.DISABLED))
            self._safe_ui_update(lambda: self.next_page_btn.config(
                state=tk.NORMAL if self.current_page < self.total_pages - 1 else tk.DISABLED))
            
            # 카테고리 버튼 강조 표시
            self._safe_ui_update(lambda: self._highlight_category_button(category_type))
            
            # 로딩 완료 메시지
            self._safe_ui_update(lambda: self.status_label.config(
                text=f"✅ CategoryType {category_type} NPC 로딩 완료: {len(all_npcs)}개"))

            # 창이 닫혔는지 확인하는 로직 추가
            if not self._is_window_valid():
                return
            
        except Exception as e:
            error_msg = f"❌ 오류 발생: {str(e)}"
            logger.error(f"NPC 로드 오류: {e}")
            print(f"NPC 로드 오류: {e}")
            self._safe_ui_update(lambda: self.status_label.config(text=error_msg))

    def _load_next_page(self):
        """다음 페이지를 로드합니다."""
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._load_category(self.current_category)

    def _load_prev_page(self):
        """이전 페이지를 로드합니다."""
        if self.current_page > 0:
            self.current_page -= 1
            self._load_category(self.current_category)

    def _highlight_category_button(self, category_type):
        """현재 선택된 카테고리 버튼을 강조 표시합니다."""
        # 모든 버튼 기본 스타일로 초기화
        self.cat_30_btn.config(bg="SystemButtonFace", relief=tk.RAISED)
        self.cat_31_btn.config(bg="SystemButtonFace", relief=tk.RAISED)
        self.cat_32_btn.config(bg="SystemButtonFace", relief=tk.RAISED)
        
        # 선택된 버튼 강조
        if category_type == 30:
            self.cat_30_btn.config(bg="#e0e0ff", relief=tk.SUNKEN)
        elif category_type == 31:
            self.cat_31_btn.config(bg="#e0e0ff", relief=tk.SUNKEN)
        elif category_type == 32:
            self.cat_32_btn.config(bg="#e0e0ff", relief=tk.SUNKEN)
    
    def _reset_search(self):
        """검색 결과를 초기화하고 전체 NPC 목록을 다시 로드합니다."""
        # 검색어 초기화
        self.event_id_entry.delete(0, tk.END)
        
        # 필터 초기화
        self.filter_var.set("")
        self.exclude_hash_var.set(False)
        
        # 전체 목록 다시 로드
        self._safe_ui_update(lambda: self.status_label.config(text="🔄 전체 NPC 목록 다시 로드 중..."))
        threading.Thread(target=self._load_all_npcs, daemon=True).start()
    
    def _load_all_npcs(self):
        """모든 NPC 정보를 불러옵니다."""
        self.status_label.config(text="🔍 전체 NPC 목록 로딩 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=self._load_npcs_thread, daemon=True).start()

    def _load_npcs_thread(self):
        """백그라운드 스레드에서 NPC 정보를 DB에서 로드합니다."""
        if not self._is_window_valid():
            return
        self._load_category_thread(30)  # CategoryType 30만 로드
        try:
            self._safe_ui_update(lambda: self.status_label.config(text="🔍 전체 NPC 목록 로딩 중..."))
            print("전체 NPC 목록 로드 시작")
            
            all_npcs = []
            
            # DB 파일에서만 데이터 로드 (Excel 로드 코드 제거)
            try:
                self._load_from_db(all_npcs)
                print(f"DB에서 NPC 정보 로드 완료: {len(all_npcs)}개")
            except Exception as db_e:
                logger.error(f"DB 로드 오류: {db_e}")
                print(f"DB 로드 오류: {db_e}")
                
                self._safe_ui_update(lambda: self.status_label.config(
                    text=f"❌ DB 로드 오류: {str(db_e)}"))
                return
            
            
            # ID 기준으로 중복 제거
            unique_npcs = {}
            for npc in all_npcs:
                if len(npc) >= 5 and str(npc[1]) not in unique_npcs:  # ID 유효성 검사 추가
                    unique_npcs[str(npc[1])] = npc
            
            print(f"중복 제거 후 NPC 개수: {len(unique_npcs)}개")
            
            # 트리뷰 업데이트
            self._safe_ui_update(lambda: self.npc_tree.delete(*self.npc_tree.get_children()))

            # ID 기준 정렬
            sorted_npcs = sorted(unique_npcs.values(), key=lambda x: int(str(x[1]).split('.')[0]) if str(x[1]).split('.')[0].isdigit() else 999999)
            
            # 디버깅: 첫 몇개 아이템 출력
            if sorted_npcs:
                print("첫 5개 아이템 샘플:")
                for i, npc in enumerate(sorted_npcs[:5]):
                    print(f"  {i+1}. {npc}")
            
            # 트리뷰에 추가
            for idx, npc in enumerate(sorted_npcs):
                self._safe_ui_update(lambda idx=idx, values=npc: 
                            self.npc_tree.insert("", "end", iid=f"npc_{idx}", values=values))
            
            # 로딩 완료 메시지
            self._safe_ui_update(lambda: self.status_label.config(
                text=f"✅ 전체 NPC 목록 로딩 완료: {len(sorted_npcs)}개"))
            
            # 기본 필터(NPC) 적용
            self.top.after(100, self._apply_category_filter)

            if not self._is_window_valid():
                return

        except Exception as e:
            error_msg = f"❌ 오류 발생: {str(e)}"
            logger.error(f"NPC 로드 오류: {e}")
            print(f"NPC 로드 오류: {e}")
            self._safe_ui_update(lambda: self.status_label.config(text=error_msg))



    def _load_from_db(self, all_npcs, category_type=30, limit=1000, offset=0):
        """지정된 CategoryType의 NPC 정보를 페이지 단위로 로드합니다."""
        db_path = os.path.join(self.db_folder, "HeroTemplate.db")
        
        if not os.path.exists(db_path):
            raise FileNotFoundError(f"HeroTemplate.db 파일을 찾을 수 없음: {db_path}")
        
        # DB 연결
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 테이블 컬럼 정보 가져오기
        cursor.execute("PRAGMA table_info(HeroTemplate)")
        columns_info = cursor.fetchall()
        column_names = [col[1] for col in columns_info]
        
        # CategoryType 컬럼 찾기
        category_type_col = None
        for col in column_names:
            if col == 'CategoryType':
                category_type_col = col
                break
        
        # 필요한 컬럼만 가져오기 (전체 *가 아닌)
        needed_columns = "UniqueID, BaseHeroID, Name, NickName, CategoryType, EventGroupID, SizeX, SizeY, ModelID"
        
        # 페이지네이션과 필터링을 적용한 쿼리
        if category_type_col:
            query = f"SELECT {needed_columns} FROM HeroTemplate WHERE {category_type_col} = ? LIMIT ? OFFSET ?"
            print(f"SQL 쿼리 실행: CategoryType={category_type}, LIMIT={limit}, OFFSET={offset}")
            cursor.execute(query, (category_type, limit, offset))
        else:
            cursor.execute(f"SELECT {needed_columns} FROM HeroTemplate LIMIT ? OFFSET ?", (limit, offset))
        
        rows = cursor.fetchall()
        print(f"DB에서 가져온 NPC 항목 수: {len(rows)}개")
        
        # 각 NPC 정보 처리
        for row in rows:
            # 필요한 필드만 인덱스로 처리 (컬럼명 순서에 의존하지 않음)
            unique_id = row[0]
            name = row[2] if row[2] else "이름 없음"
            category_type_name = "NPC"  # 이미 필터링됨
            event_group_id = row[5] if len(row) > 5 else ""
            nickname = row[3] if len(row) > 3 else ""
            size_x = row[6] if len(row) > 6 else ""
            size_y = row[7] if len(row) > 7 else ""
            model_id = row[8] if len(row) > 8 else ""
            
            # 상태 결정
            status = "사용중"
            if name.startswith('#'):
                status = "비활성화"
            
            # NPC 정보 저장
            all_npcs.append((
                status,
                unique_id,
                name,
                category_type_name,
                event_group_id,
                nickname,
                size_x,
                size_y,
                model_id
            ))
        
        conn.close()
        
        # 총 항목 수 반환 (필요 시)
        return len(rows)


    def _on_spawn_condition_select(self, event):
        """스폰 정보 선택 시 연결된 Condition 정보를 로드합니다."""
        selected_item = self.spawn_tree.focus()
        if not selected_item:
            return
        
        values = self.spawn_tree.item(selected_item, "values")
        if not values or len(values) < 5:
            return
        
        # ShowConditionTID와 HideConditionTID 가져오기
        show_condition_tid = values[3]
        hide_condition_tid = values[4]
        
        # Condition 트리뷰 초기화
        self.spawn_condition_tree.delete(*self.spawn_condition_tree.get_children())
        
        # ShowConditionTID 조회
        if show_condition_tid and show_condition_tid != "-":
            self._load_condition_info(show_condition_tid, "Show")
        
        # HideConditionTID 조회
        if hide_condition_tid and hide_condition_tid != "-":
            self._load_condition_info(hide_condition_tid, "Hide")


    def _load_condition_info(self, condition_id, condition_type=""):
        """ConditionTemplate.db에서 Condition 정보를 로드합니다."""
        # 값이 0인 경우 조회하지 않음
        if not condition_id or condition_id == "-" or condition_id == "0" or condition_id == 0:
            print(f"ConditionID가 0이거나 비어있어 조회하지 않음: {condition_id}")
            return
        
        try:
            print(f"Condition 조회 시작: {condition_type}, ID={condition_id}")
            
            # 정수형으로 변환 시도
            if isinstance(condition_id, str) and condition_id.isdigit():
                condition_id = int(condition_id)
            
            # ConditionTemplate.db 파일 경로
            db_path = os.path.join(self.db_folder, "ConditionTemplate.db")
            
            if not os.path.exists(db_path):
                print(f"ConditionTemplate.db 파일을 찾을 수 없음: {db_path}")
                return
            
            # DB 연결
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # TemplateID로 검색
            query = "SELECT TemplateID, ConditionType, Condition1, Condition2 FROM ConditionTemplate WHERE TemplateID = ?"
            print(f"Condition 조회 쿼리: {query}, 값: {condition_id}")
            cursor.execute(query, (condition_id,))
            row = cursor.fetchone()
            
            if row:
                print(f"ConditionID {condition_id} 데이터 발견: {row}")
                
                # ConditionType이 0인 경우, ConditionGroupTemplate 추가 조회
                if row[1] == 0:
                    print(f"ConditionType이 0이므로 ConditionGroupTemplate 조회 시도")
                    
                    # 우선 기본 정보는 트리뷰에 추가
                    self._add_condition_to_tree(condition_type, row)
                    
                    # ConditionGroupTemplate.db 파일 경로
                    group_db_path = os.path.join(self.db_folder, "ConditionGroupTemplate.db")
                    
                    if not os.path.exists(group_db_path):
                        print(f"ConditionGroupTemplate.db 파일을 찾을 수 없음: {group_db_path}")
                        conn.close()
                        return
                    
                    # ConditionGroupTemplate 연결
                    group_conn = sqlite3.connect(group_db_path)
                    group_cursor = group_conn.cursor()
                    
                    # OpenConditionTID 컬럼 조회
                    group_query = "SELECT OpenConditionTID, CheckConditionTID1, CheckConditionTID2, CheckConditionTID3, CheckConditionTID4, CheckConditionTID5 FROM ConditionGroupTemplate WHERE OpenConditionTID = ?"
                    print(f"ConditionGroup 조회 쿼리: {group_query}, 값: {condition_id}")
                    group_cursor.execute(group_query, (condition_id,))
                    group_row = group_cursor.fetchone()
                    
                    if group_row:
                        print(f"ConditionGroup 데이터 발견: {group_row}")
                        
                        # CheckConditionTID 1~5 처리 (0이 아닌 값만)
                        for i in range(1, 6):
                            check_id = group_row[i]
                            if check_id and check_id != 0:
                                print(f"CheckConditionTID{i} = {check_id} 조회")
                                # 재귀적으로 조회하지만 타입은 원래 타입에 "Group-"을 추가
                                self._load_condition_info(check_id, f"{condition_type}-Group{i}")
                    else:
                        print(f"ConditionGroupTemplate에서 OpenConditionTID {condition_id}를 찾을 수 없음")
                    
                    group_conn.close()
                else:
                    # 일반 Condition인 경우 트리뷰에 추가
                    self._add_condition_to_tree(condition_type, row)
            else:
                print(f"ConditionID {condition_id}에 해당하는 정보가 없음")
            
            conn.close()
            
        except Exception as e:
            print(f"Condition 정보 로드 오류: {e}")
            import traceback
            traceback.print_exc()

    def _add_condition_to_tree(self, condition_type, row):
        """Condition 정보를 적절한 트리뷰에 추가합니다."""
        # 스폰 트리 또는 이벤트 트리에 추가
        if condition_type in ["Show", "Hide"] or "Group" in condition_type and condition_type.split("-")[0] in ["Show", "Hide"]:
            # 스폰 Condition 트리뷰에 추가
            values = []
            for i in range(4):  # TemplateID, ConditionType, Condition1, Condition2
                values.append(row[i] if i < len(row) and row[i] is not None else "-")
            
            # 트리뷰에 추가
            item_id = self.spawn_condition_tree.insert("", "end", values=values, 
                                                    tags=(condition_type.lower().split("-")[0],))
            
            # 조건 타입에 따라 배경색 설정
            if "Show" in condition_type:
                self.spawn_condition_tree.tag_configure("show", background="#e0ffe0")  # 연한 녹색
            elif "Hide" in condition_type:
                self.spawn_condition_tree.tag_configure("hide", background="#ffe0e0")  # 연한 빨간색
            
            print(f"스폰 Condition 트리에 항목 추가됨: {values}")
        else:
            # 이벤트 Condition 트리뷰에 추가
            values = [condition_type]
            for i in range(4):  # TemplateID, ConditionType, Condition1, Condition2
                values.append(row[i] if i < len(row) and row[i] is not None else "-")
            
            values.append("-")  # QuestName 빈칸
            
            # 트리뷰에 추가
            item_id = self.event_condition_tree.insert("", "end", values=values, 
                                                    tags=(condition_type.lower().split("-")[0],))
            
            # 조건 타입에 따라 배경색 설정
            if "Require" in condition_type:
                self.event_condition_tree.tag_configure("require", background="#e0e0ff")  # 연한 파란색
            elif "Hide" in condition_type:
                self.event_condition_tree.tag_configure("hide", background="#ffe0e0")  # 연한 빨간색
            
            print(f"이벤트 Condition 트리에 항목 추가됨: {values}")


    def _load_quest_info(self, dialogue_group_ids):
        """
        DialogueGroupID로 연결된 퀘스트 정보를 로드합니다.
        
        Args:
            dialogue_group_ids (list): 이벤트의 DialogueGroupID 리스트
        """
        # 트리뷰 초기화
        self.quest_tree.delete(*self.quest_tree.get_children())
        
        if not dialogue_group_ids or all(not id for id in dialogue_group_ids):
            print("연결된 DialogueGroupID가 없어 퀘스트 정보를 로드할 수 없습니다.")
            return
        
        try:
            # MissionTemplate.db 파일 경로
            mission_db_path = os.path.join(self.db_folder, "QuestMissionTemplate.db")
            
            if not os.path.exists(mission_db_path):
                print(f"QuestMissionTemplate.db 파일을 찾을 수 없음: {mission_db_path}")
                return
            
            # QuestTemplate.db 파일 경로
            quest_db_path = os.path.join(self.db_folder, "QuestTemplate.db")
            
            if not os.path.exists(quest_db_path):
                print(f"QuestTemplate.db 파일을 찾을 수 없음: {quest_db_path}")
                return
            
            # 연결된 퀘스트 TemplateID 수집
            quest_template_ids = set()
            
            # 여러 DialogueGroupID를 처리
            for dialogue_group_id in dialogue_group_ids:
                if not dialogue_group_id:
                    continue
                    
                # QuestMissionTemplate.db에 연결
                mission_conn = sqlite3.connect(mission_db_path)
                mission_cursor = mission_conn.cursor()
                
                # MissionType=404(대화)이고 MissionID2가 DialogueGroupID인 항목 검색
                try:
                    # 정수형으로 변환 시도
                    if isinstance(dialogue_group_id, str) and dialogue_group_id.isdigit():
                        dialogue_group_id = int(dialogue_group_id)
                    
                    query = "SELECT TemplateID FROM QuestMissionTemplate WHERE MissionType = 404 AND MissionID2 = ?"
                    print(f"퀘스트 미션 조회 쿼리: {query}, DialogueGroupID: {dialogue_group_id}")
                    mission_cursor.execute(query, (dialogue_group_id,))
                    mission_rows = mission_cursor.fetchall()
                    
                    if mission_rows:
                        print(f"DialogueGroupID {dialogue_group_id}와 연결된 퀘스트 미션: {len(mission_rows)}개")
                        for row in mission_rows:
                            if row[0]:  # TemplateID가 있는 경우
                                quest_template_ids.add(row[0])
                    else:
                        print(f"DialogueGroupID {dialogue_group_id}와 연결된 퀘스트 미션 없음")
                except Exception as e:
                    print(f"퀘스트 미션 조회 오류: {e}")
                
                mission_conn.close()
            
            if not quest_template_ids:
                print("연결된 퀘스트가 없습니다.")
                return
            
            print(f"총 {len(quest_template_ids)}개의 연결된 퀘스트 발견: {quest_template_ids}")
            
            # QuestTemplate에서 퀘스트 정보 조회
            quest_conn = sqlite3.connect(quest_db_path)
            quest_cursor = quest_conn.cursor()
            
            for template_id in quest_template_ids:
                try:
                    # QuestTemplate 조회
                    query = "SELECT TemplateID, GroupID, QuestType, QuestName, RewardGroupID FROM QuestTemplate WHERE TemplateID = ?"
                    quest_cursor.execute(query, (template_id,))
                    quest_row = quest_cursor.fetchone()
                    
                    if quest_row:
                        # 트리뷰에 추가
                        values = []
                        for i in range(5):  # TemplateID, GroupID, QuestType, QuestName, RewardGroupID
                            values.append(quest_row[i] if i < len(quest_row) and quest_row[i] is not None else "-")
                        
                        self.quest_tree.insert("", "end", values=tuple(values))
                        print(f"퀘스트 추가: TemplateID={values[0]}, QuestName={values[3]}")
                    else:
                        print(f"TemplateID {template_id}에 해당하는 퀘스트 정보가 없음")
                except Exception as e:
                    print(f"퀘스트 정보 조회 오류 (TemplateID={template_id}): {e}")
            
            quest_conn.close()
            
        except Exception as e:
            print(f"퀘스트 연결 정보 로드 오류: {e}")
            import traceback
            traceback.print_exc()

    def _apply_category_filter(self):
        """카테고리 필터를 적용합니다."""
        if not self._is_window_valid():
            return
        
        category = self.category_var.get()
        
        # 모든 항목을 원래 상태로 복원
        self._restore_tree()
        
        # 전체 카테고리인 경우 모든 항목 표시
        if category == "ALL":
            # 다른 필터 재적용 (필요한 경우)
            if self.filter_var.get():
                self._apply_filter()
            if self.exclude_hash_var.get():
                self._apply_name_filter()
            self.status_label.config(text="✅ 전체 카테고리 표시")
            return
            
        # 필터링할 CategoryType 값 결정
        category_types = []
        if category == "HERO":
            category_types = ["영웅"]
        elif category == "MONSTER":
            category_types = ["몬스터"]
        elif category == "NPC":
            category_types = ["NPC"]
        elif category == "OTHER":
            category_types = ["기타"]
        
        # 해당 카테고리가 아닌 항목 숨기기
        to_detach = []
        for item in self.npc_tree.get_children():
            values = self.npc_tree.item(item, "values")
            if len(values) >= 4:  # 타입 정보가 있는지 확인
                item_type = values[3]
                if item_type not in category_types:
                    to_detach.append(item)
                    self.npc_tree.detach(item)
                    self._detached_items.append(item)  # 분리한 항목 저장
        
        # 다른 필터도 적용 (필요한 경우)
        if self.filter_var.get():
            self._apply_filter()
        if self.exclude_hash_var.get():
            self._apply_name_filter()
        
        # 상태 메시지 업데이트
        visible_count = len(self.npc_tree.get_children())
        total_count = visible_count + len(to_detach)
        self.status_label.config(text=f"✅ {category} 카테고리 필터 적용: {visible_count}/{total_count}개 항목 표시")
        
    
    def _on_npc_select(self, event):
        """NPC 선택 시 상세 정보를 표시합니다."""
        selected_item = self.npc_tree.focus()
        if not selected_item:
            return
        
        values = self.npc_tree.item(selected_item, "values")
        if not values:
            return
        
        # 상태, ID, 이름, 타입, EventID 순서
        status = values[0]
        npc_id = values[1]
        name = values[2]
        npc_type = values[3]
        event_id = values[4]
        
        # 기본 정보 표시
        self._display_basic_info(npc_id, selected_item)
        
        # 이벤트 방향 정보 로드
        self._load_event_direction_info(event_id)
        
        # 스폰 정보 로드
        self._load_map_spawn_info(npc_id)
    
        # Condition 정보 초기화
        self.event_condition_tree.delete(*self.event_condition_tree.get_children())
        self.spawn_condition_tree.delete(*self.spawn_condition_tree.get_children())


    def _display_basic_info(self, npc_id, selected_item):
        """선택한 NPC의 기본 정보를 표시합니다."""
        values = self.npc_tree.item(selected_item, "values")
        
        # 상태, ID, 이름, 타입, EventID, NickName, SizeX, SizeY, ModelID 순서
        npc_info = {}
        
        # 기본 정보 (목록에서 바로 가져옴)
        npc_info["id_value"] = values[1]
        npc_info["name_value"] = values[2]
        npc_info["type_value"] = values[3]
        npc_info["eventid_value"] = values[4]
        
        # 상세 정보는 DB에서 조회
        try:
            # HeroTemplate.db 파일 경로
            db_path = os.path.join(self.db_folder, "HeroTemplate.db")
            
            if os.path.exists(db_path):
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # NPC ID로 검색
                try:
                    # 정수형으로 변환 시도
                    numeric_id = int(float(npc_id))
                    cursor.execute("SELECT * FROM HeroTemplate WHERE UniqueID = ?", (numeric_id,))
                except (ValueError, TypeError):
                    # 문자열로 검색
                    cursor.execute("SELECT * FROM HeroTemplate WHERE UniqueID = ?", (npc_id,))
                
                row = cursor.fetchone()
                
                if row:
                    # 테이블 컬럼 정보 가져오기
                    cursor.execute("PRAGMA table_info(HeroTemplate)")
                    columns_info = cursor.fetchall()
                    column_names = [col[1] for col in columns_info]
                    
                    # 필요한 컬럼 인덱스 찾기
                    nickname_idx = None
                    size_x_idx = None
                    size_y_idx = None
                    model_id_idx = None
                    
                    for i, col in enumerate(column_names):
                        col_lower = col.lower()
                        if 'nickname' in col_lower:
                            nickname_idx = i
                        elif 'sizex' in col_lower or 'size_x' in col_lower:
                            size_x_idx = i
                        elif 'sizey' in col_lower or 'size_y' in col_lower:
                            size_y_idx = i
                        elif 'modelid' in col_lower or 'model_id' in col_lower:
                            model_id_idx = i
                    
                    # 상세 정보 설정
                    npc_info["nickname_value"] = str(row[nickname_idx]) if nickname_idx is not None and nickname_idx < len(row) and row[nickname_idx] is not None else "-"
                    npc_info["sizex_value"] = str(row[size_x_idx]) if size_x_idx is not None and size_x_idx < len(row) and row[size_x_idx] is not None else "-"
                    npc_info["sizey_value"] = str(row[size_y_idx]) if size_y_idx is not None and size_y_idx < len(row) and row[size_y_idx] is not None else "-"
                    npc_info["modelid_value"] = str(row[model_id_idx]) if model_id_idx is not None and model_id_idx < len(row) and row[model_id_idx] is not None else "-"
                else:
                    # DB에서 찾지 못한 경우 목록에서 가져온 값 사용
                    npc_info["nickname_value"] = values[5] if len(values) > 5 else "-"
                    npc_info["sizex_value"] = values[6] if len(values) > 6 else "-"
                    npc_info["sizey_value"] = values[7] if len(values) > 7 else "-"
                    npc_info["modelid_value"] = values[8] if len(values) > 8 else "-"
                
                conn.close()
        except Exception as e:
            logger.error(f"DB에서 NPC 상세 정보 조회 오류: {e}")
            # 오류 발생 시 목록에서 가져온 값 사용
            npc_info["nickname_value"] = values[5] if len(values) > 5 else "-"
            npc_info["sizex_value"] = values[6] if len(values) > 6 else "-"
            npc_info["sizey_value"] = values[7] if len(values) > 7 else "-"
            npc_info["modelid_value"] = values[8] if len(values) > 8 else "-"
        
        # 레이블에 값 설정
        for key, value in npc_info.items():
            if key in self.info_labels:
                self.info_labels[key].config(text=value)
                

    def _load_event_direction_info(self, event_group_id):
        """EventDirection.db에서 이벤트 방향 정보를 로드합니다."""
        # 트리뷰 초기화
        self.event_tree.delete(*self.event_tree.get_children())
        
        if not event_group_id or event_group_id == "-":
            return
        
        try:
            # 정수형으로 변환 시도
            if isinstance(event_group_id, str) and event_group_id.isdigit():
                event_group_id = int(event_group_id)
            
            # EventDirection.db 파일 경로
            db_path = os.path.join(self.db_folder, "EventDirection.db")
            
            if not os.path.exists(db_path):
                print(f"EventDirection.db 파일을 찾을 수 없음: {db_path}")
                return
            
            # DB 연결
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # GroupID로 바로 검색
            query = "SELECT * FROM EventDirection WHERE GroupID = ?"
            print(f"이벤트 방향 조회 쿼리: {query}, 값: {event_group_id}")
            cursor.execute(query, (event_group_id,))
            rows = cursor.fetchall()
            
            if not rows:
                print(f"GroupID {event_group_id}에 대한 이벤트 방향 정보가 없음")
                conn.close()
                return
            
            print(f"GroupID {event_group_id}로 {len(rows)}개 항목 발견")
            
            # 테이블 컬럼 정보 가져오기
            cursor.execute("PRAGMA table_info(EventDirection)")
            columns_info = cursor.fetchall()
            column_names = [col[1] for col in columns_info]
            
            # DialogueGroupID 컬럼 인덱스 찾기
            dialogue_group_id_idx = -1
            try:
                dialogue_group_id_idx = column_names.index("DialogueGroupID")
            except ValueError:
                print("DialogueGroupID 컬럼을 찾을 수 없음")
            
            # DialogueGroupID 값 수집 (퀘스트 연결 정보 검색용)
            dialogue_group_ids = []
            
            # 필요한 컬럼 인덱스 찾기
            indices = {}
            for col in ["GroupID", "DialogueGroupID", "RequireType", "RequireOption", 
                    "HideType", "HideOption", "EndAction", "EndOption1", "EndOption2", "EndOption3"]:
                try:
                    indices[col] = column_names.index(col)
                except ValueError:
                    indices[col] = -1  # 컬럼이 없는 경우
            
            # 결과 처리
            for row in rows:
                display_values = []
                
                # DialogueGroupID 값 수집
                if dialogue_group_id_idx >= 0 and dialogue_group_id_idx < len(row):
                    dialogue_id = row[dialogue_group_id_idx]
                    if dialogue_id:  # null이 아닌 경우
                        dialogue_group_ids.append(dialogue_id)
                
                # 각 필요한 컬럼 값 추출
                for col in ["GroupID", "DialogueGroupID", "RequireType", "RequireOption", 
                        "HideType", "HideOption", "EndAction", "EndOption1", "EndOption2", "EndOption3"]:
                    idx = indices[col]
                    if idx >= 0 and idx < len(row):
                        display_values.append(row[idx] if row[idx] is not None else "-")
                    else:
                        display_values.append("-")  # 컬럼이 없는 경우
                
                # 트리뷰에 추가
                self.event_tree.insert("", "end", values=tuple(display_values))
            
            conn.close()
            
            # 수집된 DialogueGroupID를 이용해 퀘스트 연결 정보 로드
            if dialogue_group_ids:
                print(f"DialogueGroupID 수집됨: {dialogue_group_ids}")
                self._load_quest_info(dialogue_group_ids)
            else:
                print("수집된 DialogueGroupID가 없어 퀘스트 연결 정보를 로드할 수 없습니다.")

            if self.event_tree.get_children():
                first_item = self.event_tree.get_children()[0]
                self.event_tree.selection_set(first_item)
                self.event_tree.focus(first_item)
                print("이벤트 트리 첫 번째 항목 자동 선택됨")
                
                # 중요: 수동으로 이벤트 핸들러 직접 호출
                self._on_event_condition_select(None)            

        except Exception as e:
            print(f"이벤트 방향 정보 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            

    def _load_map_spawn_info(self, npc_id):
        """MapSpawn.db에서 스폰 정보를 로드합니다."""
        # 트리뷰 초기화
        self.spawn_tree.delete(*self.spawn_tree.get_children())
        
        if not npc_id or npc_id == "-":
            return
        
        try:
            # MapSpawn.db 파일 경로
            db_path = os.path.join(self.db_folder, "MapSpawn.db")
            
            if not os.path.exists(db_path):
                print(f"MapSpawn.db 파일을 찾을 수 없음: {db_path}")
                return
            
            # DB 연결
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 테이블 컬럼 정보 가져오기
            cursor.execute("PRAGMA table_info(MapSpawn)")
            columns_info = cursor.fetchall()
            column_names = [col[1] for col in columns_info]
            
            # 필요한 컬럼 찾기
            hero_id_col = None
            map_id_col = None
            map_spawn_group_id_col = None
            helper_name_col = None
            show_condition_tid_col = None
            hide_condition_tid_col = None
            
            for col in column_names:
                col_lower = col.lower()
                if 'heroid' in col_lower:
                    hero_id_col = col
                elif 'mapid' in col_lower and not 'mapspawngroupid' in col_lower:
                    map_id_col = col
                elif 'mapspawngroupid' in col_lower:
                    map_spawn_group_id_col = col
                elif 'helpername' in col_lower:
                    helper_name_col = col
                elif 'showcondition' in col_lower:
                    show_condition_tid_col = col
                elif 'hidecondition' in col_lower:
                    hide_condition_tid_col = col
            
            if not hero_id_col:
                print("HeroID 컬럼을 찾을 수 없음")
                conn.close()
                return
            
            # 쿼리 구성
            query = f"SELECT * FROM MapSpawn WHERE {hero_id_col} = ?"
            cursor.execute(query, (npc_id,))
            rows = cursor.fetchall()
            
            if not rows:
                print(f"HeroID {npc_id}에 대한 스폰 정보가 없음")
                conn.close()
                return
            
            # 결과 처리
            for row in rows:
                # 각 컬럼 인덱스 찾기
                map_id_idx = column_names.index(map_id_col) if map_id_col in column_names else -1
                map_spawn_group_id_idx = column_names.index(map_spawn_group_id_col) if map_spawn_group_id_col in column_names else -1
                helper_name_idx = column_names.index(helper_name_col) if helper_name_col in column_names else -1
                show_condition_tid_idx = column_names.index(show_condition_tid_col) if show_condition_tid_col in column_names else -1
                hide_condition_tid_idx = column_names.index(hide_condition_tid_col) if hide_condition_tid_col in column_names else -1
                
                # 값 가져오기
                map_id = row[map_id_idx] if map_id_idx >= 0 else "-"
                map_spawn_group_id = row[map_spawn_group_id_idx] if map_spawn_group_id_idx >= 0 else "-"
                helper_name = row[helper_name_idx] if helper_name_idx >= 0 else "-"
                show_condition_tid = row[show_condition_tid_idx] if show_condition_tid_idx >= 0 else "-"
                hide_condition_tid = row[hide_condition_tid_idx] if hide_condition_tid_idx >= 0 else "-"
                
                # 트리뷰에 추가
                values = (
                    map_id, map_spawn_group_id, helper_name, show_condition_tid, hide_condition_tid
                )
                self.spawn_tree.insert("", "end", values=values)
            
            conn.close()

            # 항목이 있으면 첫 번째 항목을 자동으로 선택하여 Condition 정보 로드 트리거
            if self.spawn_tree.get_children():
                first_item = self.spawn_tree.get_children()[0]
                self.spawn_tree.selection_set(first_item)
                self.spawn_tree.focus(first_item)
                print("스폰 트리 첫 번째 항목 자동 선택됨")
                
                # 수동으로 이벤트 핸들러 직접 호출
                self._on_spawn_condition_select(None)

        except Exception as e:
            print(f"스폰 정보 로드 오류: {e}")

    def _on_event_condition_select(self, event):
        """이벤트 정보 선택 시 연결된 Condition 및 퀘스트 정보를 로드합니다."""
        selected_item = self.event_tree.focus()
        if not selected_item:
            return
        
        values = self.event_tree.item(selected_item, "values")
        if not values or len(values) < 6:
            return
        
        # RequireType, RequireOption, HideType, HideOption 가져오기
        require_type = values[2]
        require_option = values[3]
        hide_type = values[4]
        hide_option = values[5]
        
        # Condition 트리뷰 초기화
        self.event_condition_tree.delete(*self.event_condition_tree.get_children())
        
        # RequireType 처리
        if require_type and require_type != "-":
            try:
                require_type_int = int(require_type)
                
                # ConditionTemplate 검색 (Type = 50)
                if require_type_int in [50, 51] and require_option and require_option != "-":
                    self._load_condition_info(require_option, "Require")
                
                # QuestTemplate 검색 (Type = 40, 41, 42)
                elif require_type_int in [40, 41, 42] and require_option and require_option != "-":
                    self._load_quest_condition_info(require_option, f"Require-{require_type_int}")
            except ValueError:
                print(f"RequireType '{require_type}' 변환 오류")
        
        # HideType 처리
        if hide_type and hide_type != "-":
            try:
                hide_type_int = int(hide_type)
                
                # ConditionTemplate 검색 (Type = 50)
                if hide_type_int in [50, 51] and hide_option and hide_option != "-":
                    self._load_condition_info(hide_option, "Hide")
                
                # QuestTemplate 검색 (Type = 40, 41, 42)
                elif hide_type_int in [40, 41, 42] and hide_option and hide_option != "-":
                    self._load_quest_condition_info(hide_option, f"Hide-{hide_type_int}")
            except ValueError:
                print(f"HideType '{hide_type}' 변환 오류")

    def _load_quest_condition_info(self, quest_id, condition_type=""):
        """QuestTemplate.db에서 퀘스트 정보를 로드합니다."""
        if not quest_id or quest_id == "-":
            return
        
        try:
            # 정수형으로 변환 시도
            if isinstance(quest_id, str) and quest_id.isdigit():
                quest_id = int(quest_id)
            
            # QuestTemplate.db 파일 경로
            db_path = os.path.join(self.db_folder, "QuestTemplate.db")
            
            if not os.path.exists(db_path):
                print(f"QuestTemplate.db 파일을 찾을 수 없음: {db_path}")
                return
            
            # DB 연결
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # TemplateID로 검색
            query = "SELECT TemplateID, QuestName FROM QuestTemplate WHERE TemplateID = ?"
            print(f"퀘스트 조회 쿼리: {query}, 값: {quest_id}")
            cursor.execute(query, (quest_id,))
            row = cursor.fetchone()
            
            if row:
                print(f"QuestID {quest_id} 데이터 발견: {row}")
                
                # 이벤트 Condition 트리뷰에 추가
                values = [condition_type, quest_id, "-", "-", "-", row[1] if row[1] else "-"]
                
                item_id = self.event_condition_tree.insert("", "end", values=values, 
                                                        tags=(condition_type.lower(),))
                
                # 조건 타입에 따라 배경색 설정
                if "Require" in condition_type:
                    self.event_condition_tree.tag_configure(condition_type.lower(), background="#e0ffe0")  # 연한 녹색
                else:
                    self.event_condition_tree.tag_configure(condition_type.lower(), background="#ffe0ff")  # 연한 보라색
            else:
                print(f"QuestID {quest_id}에 해당하는 정보가 없음")
            
            conn.close()
            
        except Exception as e:
            print(f"퀘스트 정보 로드 오류: {e}")
            import traceback
            traceback.print_exc()


    def _add_hash_to_selected(self):
        """선택한 NPC의 엑셀 A열에 #을 추가합니다."""
        selected_item = self.npc_tree.focus()
        if not selected_item:
            messagebox.showwarning("선택 오류", "NPC를 선택해주세요.")
            return
        
        values = self.npc_tree.item(selected_item, "values")
        if not values or len(values) < 2:
            messagebox.showwarning("데이터 오류", "선택한 NPC의 데이터가 유효하지 않습니다.")
            return
        
        npc_id = values[1]  # NPC ID (UniqueID)
        print(f"[숨기기] NPC ID {npc_id} 엑셀 파일 검색 시작")
        
        # NPC 관련 엑셀 파일 찾기 (HeroTemplate 파일)
        excel_files_found = []
        
        for file, info in self.cache.items():
            if 'hero' in file.lower() or 'herotemplate' in file.lower():
                print(f"[숨기기] 가능성 있는 파일 발견: {file}")
                excel_files_found.append((file, info["path"]))
        
        if not excel_files_found:
            print(f"[숨기기] NPC 관련 엑셀 파일을 찾을 수 없음")
            messagebox.showwarning("파일 없음", "NPC 관련 엑셀 파일을 찾을 수 없습니다.")
            return
        
        # 변수 추가: 처리 결과 추적
        total_files_processed = 0
        total_matches_found = 0
        
        # 각 파일에서 시트 검색
        for file, path in excel_files_found:
            print(f"[숨기기] 파일 검색: {file}")
            file_info = self.cache.get(file, {})
            
            for sheet, meta in file_info.get("sheets", {}).items():
                print(f"[숨기기] 시트 검사: {sheet}")
                try:
                    header = meta["header_row"]
                    
                    # 엑셀 파일 읽기
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # UniqueID 컬럼 찾기
                    unique_id_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'uniqueid' in col_lower or 'unique_id' in col_lower:
                            unique_id_col = col
                            print(f"[숨기기] ID 컬럼 발견: {col}")
                            break
                    
                    if not unique_id_col:
                        print(f"[숨기기] ID 컬럼 없음, 다음 시트 확인")
                        continue
                    
                    # ID가 모두 숫자인지 확인
                    is_numeric = True
                    try:
                        df[unique_id_col] = pd.to_numeric(df[unique_id_col], errors='coerce')
                        df[unique_id_col] = df[unique_id_col].fillna(0).astype(int)
                        is_numeric = True
                    except:
                        is_numeric = False
                    
                    # 검색 방법 선택
                    if is_numeric:
                        print(f"[숨기기] 숫자형 ID 검색")
                        try:
                            numeric_npc_id = int(npc_id)
                            matched = df[df[unique_id_col] == numeric_npc_id]
                        except:
                            print(f"[숨기기] 숫자 변환 실패, 문자열 검색으로 전환")
                            df[unique_id_col] = df[unique_id_col].astype(str)
                            matched = df[df[unique_id_col] == npc_id]
                    else:
                        print(f"[숨기기] 문자열 ID 검색")
                        df[unique_id_col] = df[unique_id_col].astype(str)
                        matched = df[df[unique_id_col] == npc_id]
                    
                    print(f"[숨기기] 검색 결과: {len(matched)}행")
                    
                    if not matched.empty:
                        total_matches_found += len(matched)
                        print(f"[숨기기] {file}/{sheet}에서 NPC ID {npc_id} 발견")
                        # A열에 # 추가 - 여기서 header_row 값 전달
                        from utils.excel_utils import ExcelFileManager
                        
                        result = ExcelFileManager.add_hash_to_a_column(path, sheet, npc_id, header_row=header, id_column=unique_id_col)
                        # 수정 성공 시 업데이트 부분
                        if result:
                            show_message(self.top, "info", "성공", f"NPC ID {npc_id}의 A열에 #이 추가되었습니다.")
                            
                            # 개별 항목만 업데이트 (전체 리스트 리프레시 없음)
                            self._refresh_selected_item(npc_id, is_hidden=True)
                            
                            # 목록에서도 # 표시와 상태 업데이트 (트리뷰 전체를 새로고침하지 않음)
                            current_name = values[2]
                            if not current_name.startswith('#'):
                                new_name = f"#{current_name}"
                                self.npc_tree.item(selected_item, values=("비활성화", values[1], new_name, values[3], values[4]))
                            else:
                                # 이미 #이 있으면 상태만 업데이트
                                self.npc_tree.item(selected_item, values=("비활성화", values[1], current_name, values[3], values[4]))
                                
                except Exception as e:
                    print(f"[파일 검색 오류] {file} / {sheet}: {e}")
        
        # 모든 파일 검색 후 결과 처리
        if total_matches_found > 0:
            show_message(self.top, "info", "성공", f"NPC ID {npc_id}의 A열에 # 추가 처리가 완료되었습니다. ({total_matches_found}개 항목, {total_files_processed}개 파일)")
        else:
            show_message(self.top, "warning", "항목 없음", f"NPC ID {npc_id}에 해당하는 데이터를 엑셀 파일에서 찾을 수 없습니다.")

    def _remove_hash_from_selected(self):
        """선택한 NPC의 엑셀 A열에서 #을 제거합니다."""
        selected_item = self.npc_tree.focus()
        if not selected_item:
            messagebox.showwarning("선택 오류", "NPC를 선택해주세요.")
            return
        
        values = self.npc_tree.item(selected_item, "values")
        if not values or len(values) < 2:
            messagebox.showwarning("데이터 오류", "선택한 NPC의 데이터가 유효하지 않습니다.")
            return
        
        npc_id = values[1]  # NPC ID (UniqueID)
        print(f"[숨김 해제] NPC ID {npc_id} 엑셀 파일 검색 시작")
        
        # NPC 관련 엑셀 파일 찾기 (HeroTemplate 파일)
        excel_files_found = []
        
        for file, info in self.cache.items():
            if 'hero' in file.lower() or 'herotemplate' in file.lower():
                print(f"[숨김 해제] 가능성 있는 파일 발견: {file}")
                excel_files_found.append((file, info["path"]))
        
        if not excel_files_found:
            print(f"[숨김 해제] NPC 관련 엑셀 파일을 찾을 수 없음")
            messagebox.showwarning("파일 없음", "NPC 관련 엑셀 파일을 찾을 수 없습니다.")
            return
        
        # 변수 추가: 처리 결과 추적
        total_files_processed = 0
        total_matches_found = 0
        
        # 각 파일에서 시트 검색
        for file, path in excel_files_found:
            print(f"[숨김 해제] 파일 검색: {file}")
            file_info = self.cache.get(file, {})
            
            for sheet, meta in file_info.get("sheets", {}).items():
                print(f"[숨김 해제] 시트 검사: {sheet}")
                try:
                    header = meta["header_row"]
                    
                    # 엑셀 파일 읽기
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # UniqueID 컬럼 찾기
                    unique_id_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'uniqueid' in col_lower or 'unique_id' in col_lower:
                            unique_id_col = col
                            print(f"[숨김 해제] ID 컬럼 발견: {col}")
                            break
                    
                    if not unique_id_col:
                        print(f"[숨김 해제] ID 컬럼 없음, 다음 시트 확인")
                        continue
                    
                    # ID가 모두 숫자인지 확인
                    is_numeric = True
                    try:
                        df[unique_id_col] = pd.to_numeric(df[unique_id_col], errors='coerce')
                        df[unique_id_col] = df[unique_id_col].fillna(0).astype(int)
                        is_numeric = True
                    except:
                        is_numeric = False
                    
                    # 검색 방법 선택
                    if is_numeric:
                        print(f"[숨김 해제] 숫자형 ID 검색")
                        try:
                            numeric_npc_id = int(npc_id)
                            matched = df[df[unique_id_col] == numeric_npc_id]
                        except:
                            print(f"[숨김 해제] 숫자 변환 실패, 문자열 검색으로 전환")
                            df[unique_id_col] = df[unique_id_col].astype(str)
                            matched = df[df[unique_id_col] == npc_id]
                    else:
                        print(f"[숨김 해제] 문자열 ID 검색")
                        df[unique_id_col] = df[unique_id_col].astype(str)
                        matched = df[df[unique_id_col] == npc_id]
                    
                    print(f"[숨김 해제] 검색 결과: {len(matched)}행")
                    
                    if not matched.empty:
                        total_matches_found += len(matched)
                        print(f"[숨김 해제] {file}/{sheet}에서 NPC ID {npc_id} 발견")
                        # A열에서 # 제거
                        from utils.excel_utils import ExcelFileManager
                        
                        result = ExcelFileManager.remove_hash_from_a_column(path, sheet, npc_id, header_row=header, id_column=unique_id_col)
                        # 수정: 이름과 상태 모두 업데이트
                        if result:
                            show_message(self.top, "info", "성공", f"NPC ID {npc_id}의 A열에서 #이 제거되었습니다.")
                            
                            # 목록에서도 # 제거와 상태 업데이트
                            current_name = values[2]
                            if current_name.startswith('#'):
                                new_name = current_name[1:]
                                self.npc_tree.item(selected_item, values=("사용중", values[1], new_name, values[3], values[4]))
                            else:
                                # 이미 #이 없으면 상태만 업데이트
                                self.npc_tree.item(selected_item, values=("사용중", values[1], current_name, values[3], values[4]))
                        
                except Exception as e:
                    print(f"[파일 검색 오류] {file} / {sheet}: {e}")
        
        # 모든 파일 검색 후 결과 처리
        if total_matches_found > 0:
            show_message(self.top, "info", "성공", f"NPC ID {npc_id}의 A열에서 # 제거 처리가 완료되었습니다. ({total_matches_found}개 항목, {total_files_processed}개 파일)")
        else:
            show_message(self.top, "warning", "항목 없음", f"NPC ID {npc_id}에 해당하는 데이터를 엑셀 파일에서 찾을 수 없습니다.")

    def _refresh_selected_item(self, npc_id, is_hidden=False):
        """선택한 항목만 리프레시합니다."""
        # 모든 항목 확인
        for item in self.npc_tree.get_children():
            values = self.npc_tree.item(item, "values")
            if values and values[1] == npc_id:
                # 현재 값 가져오기
                current_name = values[2] if len(values) > 2 else ""
                current_type = values[3] if len(values) > 3 else ""
                event_id = values[4] if len(values) > 4 else ""
                
                # 숨김 상태에 따라 이름 및 상태 업데이트
                if is_hidden:
                    # 이름에 #이 없으면 추가
                    if not current_name.startswith('#'):
                        current_name = f"#{current_name}"
                    status = "비활성화"
                else:
                    # 이름에서 #이 있으면 제거
                    if current_name.startswith('#'):
                        current_name = current_name[1:]
                    status = "사용중"
                
                # 선택 항목 업데이트
                self.npc_tree.item(item, values=(status, npc_id, current_name, current_type, event_id))
                
                # 상세 정보도 새로고침 (선택한 항목 유지)
                self._on_npc_select(None)
                break

    def _search_by_event_id(self):
        """EventID를 포함하는 NPC를 검색합니다."""
        event_id = self.event_id_entry.get().strip()
        if not event_id:
            messagebox.showwarning("입력 오류", "EventID를 입력하세요.")
            return
        
        self.status_label.config(text=f"🔍 EventID {event_id} 검색 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=lambda: self._search_event_id_thread(event_id), daemon=True).start()

    def _search_by_name(self):
        """이름에 특정 문자열을 포함하는 NPC를 검색합니다."""
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "검색할 이름을 입력하세요.")
            return
        
        self.status_label.config(text=f"🔍 이름 '{name}' 검색 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=lambda: self._search_name_thread(name), daemon=True).start()

    def _search_name_thread(self, name):
        """백그라운드 스레드에서 이름을 검색합니다."""
        try:
            found_npcs = []
            
            # DB 연결
            db_path = os.path.join(self.db_folder, "HeroTemplate.db")
            
            if not os.path.exists(db_path):
                raise FileNotFoundError(f"HeroTemplate.db 파일을 찾을 수 없음: {db_path}")
            
            conn = sqlite3.connect(db_path)
            conn.create_function("REGEXP", 2, lambda x, y: 1 if re.search(x, y) else 0)
            cursor = conn.cursor()
            
            # 필요한 컬럼만 가져오기
            columns = "UniqueID, Name, NickName, CategoryType, EventGroupID"
            
            # 이름 검색 쿼리 작성 (부분 일치)
            query = f"SELECT {columns} FROM HeroTemplate WHERE Name LIKE ? OR NickName LIKE ?"
            search_param = f"%{name}%"
            print(f"이름 검색 쿼리: {query}, 값: {search_param}")
            cursor.execute(query, (search_param, search_param))
            
            rows = cursor.fetchall()
            print(f"이름 '{name}'로 찾은 NPC 수: {len(rows)}개")
            
            # 결과 처리
            for row in rows:
                unique_id = row[0]
                item_name = row[1] if row[1] else (row[2] if row[2] else "이름 없음")
                
                # 상태 결정 (기본값)
                status = "사용중"
                if item_name.startswith('#'):
                    status = "비활성화"
                
                # CategoryType 변환
                category_type = row[3] if len(row) > 3 else 0
                try:
                    category_type = int(category_type)
                except (ValueError, TypeError):
                    category_type = 0
                    
                if category_type in [10, 11]:
                    category_type_name = "영웅"
                elif category_type in [20, 21]:
                    category_type_name = "몬스터"
                elif category_type in [30, 31, 32]:
                    category_type_name = "NPC"
                elif category_type == 99:
                    category_type_name = "기타"
                else:
                    category_type_name = f"타입-{category_type}"
                
                # EventID 값
                event_group_id = row[4] if len(row) > 4 else ""
                
                # NPC 정보 저장
                found_npcs.append((
                    status,      # 상태
                    unique_id,   # UniqueID
                    item_name,   # 이름
                    category_type_name,   # 타입
                    event_group_id     # EventID
                ))
            
            conn.close()
            
            # 결과 표시
            self.top.after(0, lambda: self._display_search_results(found_npcs, f"이름: {name}"))
            
        except Exception as e:
            error_msg = f"❌ 검색 오류: {str(e)}"
            print(f"이름 검색 오류: {e}")
            self.top.after(0, lambda: self.status_label.config(text=error_msg))

    def _search_event_id_thread(self, event_id):
        """백그라운드 스레드에서 EventID를 검색합니다."""
        try:
            found_npcs = []
            
            # DB 연결
            db_path = os.path.join(self.db_folder, "HeroTemplate.db")
            
            if not os.path.exists(db_path):
                raise FileNotFoundError(f"HeroTemplate.db 파일을 찾을 수 없음: {db_path}")
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 필요한 컬럼만 가져오기
            columns = "UniqueID, Name, CategoryType, EventGroupID"
            
            # EventGroupID 검색 쿼리 작성
            try:
                # 숫자형 EventID 검색 시도
                numeric_event_id = int(event_id)
                query = f"SELECT {columns} FROM HeroTemplate WHERE EventGroupID = ?"
                print(f"EventID {event_id} 검색 쿼리 실행")
                cursor.execute(query, (numeric_event_id,))
            except ValueError:
                # 문자열 검색 (필요한 경우)
                query = f"SELECT {columns} FROM HeroTemplate WHERE EventGroupID = ?"
                cursor.execute(query, (event_id,))
            
            rows = cursor.fetchall()
            print(f"EventID {event_id}로 찾은 NPC 수: {len(rows)}개")
            
            # 결과 처리
            for row in rows:
                unique_id = row[0]
                name = row[1] if row[1] else "이름 없음"
                
                # 상태 결정 (기본값)
                status = "사용중"
                if name.startswith('#'):
                    status = "비활성화"
                
                # CategoryType 변환
                category_type = row[2]
                if category_type in [10, 11]:
                    category_type_name = "영웅"
                elif category_type in [20, 21]:
                    category_type_name = "몬스터"
                elif category_type in [30, 31, 32]:
                    category_type_name = "NPC"
                elif category_type == 99:
                    category_type_name = "기타"
                else:
                    category_type_name = f"타입-{category_type}"
                
                # EventID 값
                event_group_id = row[3]
                
                # NPC 정보 저장
                found_npcs.append((
                    status,      # 상태
                    unique_id,   # UniqueID
                    name,        # 이름
                    category_type_name,   # 타입
                    event_group_id     # EventID
                ))
            
            conn.close()
            
            # 결과 표시
            self.top.after(0, lambda: self._display_search_results(found_npcs, event_id))
            
        except Exception as e:
            error_msg = f"❌ 검색 오류: {str(e)}"
            print(f"EventID 검색 오류: {e}")
            self.top.after(0, lambda: self.status_label.config(text=error_msg))


    def _display_search_results(self, found_npcs, event_id):
        """검색 결과를 트리뷰에 표시합니다."""
        # 트리뷰 초기화
        self.npc_tree.delete(*self.npc_tree.get_children())
        
        # 결과가 없는 경우
        if not found_npcs:
            self.status_label.config(text=f"⚠️ EventID {event_id}를 포함하는 NPC를 찾을 수 없습니다.")
            return
        
        # 결과 정렬 및 표시
        sorted_npcs = sorted(found_npcs, key=lambda x: x[1])
        for idx, npc in enumerate(sorted_npcs):
            self.npc_tree.insert("", "end", iid=f"npc_search_{idx}", values=npc)
        
        # 상태 업데이트
        self.status_label.config(text=f"✅ EventID {event_id} 검색 완료: {len(found_npcs)}개 NPC 발견")

    def _apply_name_filter(self):
        """# 필터 체크박스에 따라 필터링을 적용합니다."""
        # 체크 상태 확인
        exclude_hash = self.exclude_hash_var.get()
        
        # 모든 항목을 원래 상태로 복원
        self._restore_tree()
        
        if exclude_hash:
            # 필터 적용 - '#'으로 시작하는 이름 제외
            to_detach = []
            for item in self.npc_tree.get_children():
                values = self.npc_tree.item(item, "values")
                if len(values) >= 3 and isinstance(values[2], str) and values[2].startswith('#'):
                    to_detach.append(item)
                    self.npc_tree.detach(item)
                    self._detached_items.append(item)  # 분리한 항목 저장
            
            self.status_label.config(text=f"✅ '#'로 시작하는 이름 제외 필터 적용 ({len(to_detach)}개 항목 숨김)")
        else:
            # 필터 초기화
            self.status_label.config(text="✅ 모든 항목 표시")
    
    def _restore_tree(self):
        """트리뷰에서 숨겨진 항목을 모두 복원합니다."""
        # 저장된 분리 항목들을 다시 붙이기
        for item in self._detached_items:
            self.npc_tree.reattach(item, "", "end")
        # 분리 항목 리스트 초기화
        self._detached_items.clear()

    def _apply_filter(self):
        """트리뷰에 필터를 적용합니다."""
        keyword = self.filter_var.get().strip().lower()
        if not keyword:
            return
            
        # 항목 숨기기 전 모든 항목을 복원
        self._restore_tree()
        
        # 항목 숨기기
        hidden_count = 0
        for item in self.npc_tree.get_children():
            values = self.npc_tree.item(item, "values")
            # 어떤 컬럼이든 키워드를 포함하는지 확인
            if not any(keyword in str(v).lower() for v in values):
                self.npc_tree.detach(item)
                self._detached_items.append(item)
                hidden_count += 1
        
        # 이후 # 필터 재적용
        if self.exclude_hash_var.get():
            self._apply_hash_filter_only()
        
        # 상태 업데이트
        if hidden_count > 0:
            self.status_label.config(
                text=f"🔍 필터 적용: '{keyword}' - {hidden_count}개 항목 숨김")
        else:
            self.status_label.config(
                text=f"🔍 필터 적용: '{keyword}' - 일치하는 항목 없음")

    def _apply_hash_filter_only(self):
        """# 필터만 별도로 적용합니다 (다른 필터 적용 후)"""
        if not self.exclude_hash_var.get():
            return 0
            
        hash_hidden = 0
        for item in self.npc_tree.get_children():
            values = self.npc_tree.item(item, "values")
            if len(values) >= 3 and isinstance(values[2], str) and values[2].startswith('#'):
                self.npc_tree.detach(item)
                self._detached_items.append(item)
                hash_hidden += 1
        
        return hash_hidden

    def _clear_filter(self):
        """트리뷰 필터를 초기화합니다."""
        # 필터 텍스트 초기화
        self.filter_var.set("")
        
        # 체크박스도 초기화
        self.exclude_hash_var.set(False)
        
        # 카테고리 필터 초기화 (전체로 설정)
        self.category_var.set("ALL")
        
        # 숨겨진 항목 모두 복원
        self._restore_tree()
                
        # 상태 업데이트
        self.status_label.config(text=f"✅ 필터 초기화 완료")