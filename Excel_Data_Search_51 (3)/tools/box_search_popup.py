# box_search_popup.py 파일

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, PanedWindow
import pandas as pd
import sqlite3
import threading
import re  # 추가 - 정규식 처리용

from utils.cache_utils import load_cached_data, hash_paths, update_excel_cache
from utils.excel_utils import ExcelFileManager
from utils.config_utils import load_search_history, save_search_history
from utils.common_utils import logger, PathUtils, FileUtils
from ui.common_components import show_message
from utils.type_mappings import get_table_name_for_type, get_description_for_type, resolve_type_info

class BoxListPopup:
    """전체 Box 목록을 표시하는 팝업 클래스"""
    def __init__(self, master, folder, db_folder, typecode_mapping, excel_cache):
        self.folder = folder
        self.db_folder = db_folder
        self.typecode_mapping = typecode_mapping
        self.cache = excel_cache
        self.top = Toplevel(master)
        self.top.title("📦 전체 Box 목록")
        self.top.geometry("1400x700")
        self._detached_items = []  # 분리된 항목을 저장할 리스트 추가
        
        self._build_ui()
        self._load_all_boxes()
        
    # def _build_ui(self):
    #     """UI를 구성합니다."""
    #     # 상단 프레임
    #     top_frame = tk.Frame(self.top)
    #     top_frame.pack(fill="x", padx=10, pady=5)
        
    #     tk.Label(top_frame, text="전체 Box 목록", font=("Helvetica", 12, "bold")).pack(side="left")

    #     # 보상 ID 검색 추가
    #     search_frame = tk.Frame(self.top)
    #     search_frame.pack(fill="x", padx=10, pady=5)
        

    #     tk.Label(search_frame, text="보상 ID:").pack(side="left")
    #     self.reward_id_entry = tk.Entry(search_frame)
    #     self.reward_id_entry.pack(side="left", padx=5)

    #     search_reward_btn = tk.Button(search_frame, text="보상 ID 검색", 
    #                                 command=self._search_by_reward_id)
    #     search_reward_btn.pack(side="left", padx=5)

    #     # 전체 목록으로 돌아가기 버튼 추가
    #     reset_search_btn = tk.Button(search_frame, text="전체 목록 보기", 
    #                             command=self._reset_search)
    #     reset_search_btn.pack(side="left", padx=5)
    
    #     # 필터 프레임
    #     filter_frame = tk.Frame(top_frame)
    #     filter_frame.pack(side="right")
        
    #     # "# 제외" 체크박스 추가
    #     self.exclude_hash_var = tk.BooleanVar(value=False)
    #     exclude_hash_check = tk.Checkbutton(filter_frame, text="'#'로 시작하는 이름 제외", 
    #                                         variable=self.exclude_hash_var,
    #                                         command=self._apply_name_filter)
    #     exclude_hash_check.pack(side="left", padx=10)
        
    #     # 기존 필터 요소들
    #     tk.Label(filter_frame, text="필터:").pack(side="left")
    #     self.filter_var = tk.StringVar()
    #     filter_entry = tk.Entry(filter_frame, textvariable=self.filter_var, width=15)
    #     filter_entry.pack(side="left", padx=5)
        
    #     filter_btn = tk.Button(filter_frame, text="적용", command=self._apply_filter)
    #     filter_btn.pack(side="left", padx=2)
        
    #     clear_btn = tk.Button(filter_frame, text="초기화", command=self._clear_filter)
    #     clear_btn.pack(side="left", padx=2)
        
    #     # PanedWindow로 좌/우 분할
    #     self.paned = PanedWindow(self.top, orient=tk.HORIZONTAL)
    #     self.paned.pack(fill="both", expand=True, padx=10, pady=5)
        
    #     # 좌측: Box 목록
    #     left_frame = tk.Frame(self.paned)
    #     self.paned.add(left_frame, width=400)
    
    #     # 버튼 프레임 추가
    #     btn_frame = tk.Frame(left_frame)
    #     btn_frame.pack(fill="x", pady=5)
        
    #     tk.Label(left_frame, text="Box 목록", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
    
    #     # 숨김 버튼 추가
    #     self.hide_btn = tk.Button(btn_frame, text="선택한 Box 숨기기", command=self._add_hash_to_selected)
    #     self.hide_btn.pack(side="left", padx=5)

    #     # 숨김 해제 버튼 추가
    #     self.unhide_btn = tk.Button(btn_frame, text="선택 Box 숨김 해제", command=self._remove_hash_from_selected)
    #     self.unhide_btn.pack(side="left", padx=5)
        
    #     # 목록 프레임 (스크롤바 포함)
    #     list_frame = tk.Frame(left_frame)
    #     list_frame.pack(fill="both", expand=True)
        
    #     # 수직 스크롤바
    #     list_scroll_y = tk.Scrollbar(list_frame)
    #     list_scroll_y.pack(side="right", fill="y")
        
    #     # 트리뷰로 구현
    #     columns = ["ItemID", "BoxName", "BoxType", "Status"]  # Status 컬럼 추가
    #     self.box_tree = ttk.Treeview(list_frame, columns=columns, show="headings", 
    #                                 yscrollcommand=list_scroll_y.set)
        
    #     # 컬럼 설정
    #     self.box_tree.heading("ItemID", text="ItemID")
    #     self.box_tree.column("ItemID", width=100, anchor="w")
    #     self.box_tree.heading("BoxName", text="Box 이름")
    #     self.box_tree.column("BoxName", width=200, anchor="w")
    #     self.box_tree.heading("BoxType", text="Box 타입")
    #     self.box_tree.column("BoxType", width=100, anchor="w")
    #     # 상태 컬럼 추가
    #     self.box_tree.heading("Status", text="상태")
    #     self.box_tree.column("Status", width=80, anchor="center")
        
    #     self.box_tree.pack(fill="both", expand=True)
    #     list_scroll_y.config(command=self.box_tree.yview)
        
    #     # 우측: 상세 정보
    #     right_frame = tk.Frame(self.paned)
    #     self.paned.add(right_frame, width=800)
        
    #     tk.Label(right_frame, text="Box 상세 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
    #     # 상세 정보 프레임
    #     detail_frame = tk.Frame(right_frame)
    #     detail_frame.pack(fill="both", expand=True)     
                
    #     # 상세 정보 버튼 프레임 추가
    #     detail_btn_frame = tk.Frame(right_frame)
    #     detail_btn_frame.pack(fill="x", pady=5)   

    #     # 선택한 보상 ID 전체 숨기기 버튼 추가
    #     self.hide_reward_btn = tk.Button(detail_btn_frame, text="선택 보상 ID 모두 숨기기", 
    #                                 command=self._add_hash_to_all_matching_rewards)
    #     self.hide_reward_btn.pack(side="left", padx=5)
        
    #     # 선택한 보상 ID 전체 숨김 해제 버튼 추가
    #     self.unhide_reward_btn = tk.Button(detail_btn_frame, text="선택 보상 ID 모두 숨김 해제", 
    #                                 command=self._remove_hash_from_all_matching_rewards)
    #     self.unhide_reward_btn.pack(side="left", padx=5)

    #     # 수직 스크롤바
    #     detail_scroll_y = tk.Scrollbar(detail_frame)
    #     detail_scroll_y.pack(side="right", fill="y")
        
    #     # 트리뷰로 구현
    #     columns = ["RewardType", "RewardID", "RewardName", "RewardCount", "RewardProbability"]
    #     self.detail_tree = ttk.Treeview(detail_frame, columns=columns, show="headings", 
    #                                    yscrollcommand=detail_scroll_y.set)
        
    #     # 컬럼 설정
    #     self.detail_tree.heading("RewardType", text="보상 타입")
    #     self.detail_tree.column("RewardType", width=80, anchor="center")
    #     self.detail_tree.heading("RewardID", text="보상 ID")
    #     self.detail_tree.column("RewardID", width=80, anchor="center")
    #     self.detail_tree.heading("RewardName", text="보상 이름")
    #     self.detail_tree.column("RewardName", width=200, anchor="w")
    #     self.detail_tree.heading("RewardCount", text="보상 개수")
    #     self.detail_tree.column("RewardCount", width=80, anchor="center")
    #     self.detail_tree.heading("RewardProbability", text="확률")
    #     self.detail_tree.column("RewardProbability", width=80, anchor="center")
        
    #     self.detail_tree.pack(fill="both", expand=True)
    #     detail_scroll_y.config(command=self.detail_tree.yview)
        
    #     # 상태 표시
    #     self.status_label = tk.Label(self.top, text="")
    #     self.status_label.pack(anchor="w", padx=10, pady=5)
        
    #     # 이벤트 연결
    #     self.box_tree.bind("<<TreeviewSelect>>", self._on_box_select)

    def _build_ui(self):
        """UI를 구성합니다."""
        # 상단 프레임
        top_frame = tk.Frame(self.top)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(top_frame, text="전체 Box 목록", font=("Helvetica", 12, "bold")).pack(side="left")

        # 보상 ID 검색 추가
        search_frame = tk.Frame(self.top)
        search_frame.pack(fill="x", padx=10, pady=5)
        

        tk.Label(search_frame, text="보상 ID:").pack(side="left")
        self.reward_id_entry = tk.Entry(search_frame)
        self.reward_id_entry.pack(side="left", padx=5)

        search_reward_btn = tk.Button(search_frame, text="보상 ID 검색", 
                                    command=self._search_by_reward_id)
        search_reward_btn.pack(side="left", padx=5)

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
        
        # 좌측: Box 목록
        left_frame = tk.Frame(self.paned)
        self.paned.add(left_frame, width=600)
        
        # 버튼 프레임 추가
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill="x", pady=5)
        
        tk.Label(left_frame, text="Box 목록", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
        # 숨김 버튼 추가
        self.hide_btn = tk.Button(btn_frame, text="선택한 Box 숨기기", command=self._add_hash_to_selected)
        self.hide_btn.pack(side="left", padx=5)

        # 숨김 해제 버튼 추가
        self.unhide_btn = tk.Button(btn_frame, text="선택 Box 숨김 해제", command=self._remove_hash_from_selected)
        self.unhide_btn.pack(side="left", padx=5)
        
        # 목록 프레임 (스크롤바 포함)
        list_frame = tk.Frame(left_frame)
        list_frame.pack(fill="both", expand=True)
        
        # 수직 스크롤바
        list_scroll_y = tk.Scrollbar(list_frame)
        list_scroll_y.pack(side="right", fill="y")
        
        # 트리뷰로 구현 - 상태 컬럼 추가
        columns = ["ItemID", "BoxName", "BoxType", "Status"]
        self.box_tree = ttk.Treeview(list_frame, columns=columns, show="headings", 
                                    yscrollcommand=list_scroll_y.set)
        
        # 컬럼 설정
        self.box_tree.heading("ItemID", text="ItemID")
        self.box_tree.column("ItemID", width=100, anchor="w")
        self.box_tree.heading("BoxName", text="Box 이름")
        self.box_tree.column("BoxName", width=200, anchor="w")
        self.box_tree.heading("BoxType", text="Box 타입")
        self.box_tree.column("BoxType", width=100, anchor="w")
        self.box_tree.heading("Status", text="상태")
        self.box_tree.column("Status", width=80, anchor="center")
        
        self.box_tree.pack(fill="both", expand=True)
        list_scroll_y.config(command=self.box_tree.yview)
        
        # 우측: 상세 정보
        right_frame = tk.Frame(self.paned)
        self.paned.add(right_frame, width=800)
        
        tk.Label(right_frame, text="Box 상세 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
        # 상세 정보 버튼 프레임 추가
        detail_btn_frame = tk.Frame(right_frame)
        detail_btn_frame.pack(fill="x", pady=5)
        
        # 선택한 보상 ID 전체 숨기기 버튼 추가
        self.hide_reward_btn = tk.Button(detail_btn_frame, text="선택 보상 ID 모두 숨기기", 
                                    command=self._add_hash_to_all_matching_rewards)
        self.hide_reward_btn.pack(side="left", padx=5)
        
        # 선택한 보상 ID 전체 숨김 해제 버튼 추가
        self.unhide_reward_btn = tk.Button(detail_btn_frame, text="선택 보상 ID 모두 숨김 해제", 
                                    command=self._remove_hash_from_all_matching_rewards)
        self.unhide_reward_btn.pack(side="left", padx=5)
        
        # 상세 정보 프레임
        detail_frame = tk.Frame(right_frame)
        detail_frame.pack(fill="both", expand=True)
        
        # 수직 스크롤바
        detail_scroll_y = tk.Scrollbar(detail_frame)
        detail_scroll_y.pack(side="right", fill="y")
        
        # 트리뷰로 구현
        columns = ["RewardType", "RewardID", "RewardName", "RewardCount", "RewardProbability"]
        self.detail_tree = ttk.Treeview(detail_frame, columns=columns, show="headings", 
                                    yscrollcommand=detail_scroll_y.set)
        
        # 컬럼 설정
        self.detail_tree.heading("RewardType", text="보상 타입")
        self.detail_tree.column("RewardType", width=80, anchor="center")
        self.detail_tree.heading("RewardID", text="보상 ID")
        self.detail_tree.column("RewardID", width=80, anchor="center")
        self.detail_tree.heading("RewardName", text="보상 이름")
        self.detail_tree.column("RewardName", width=200, anchor="w")
        self.detail_tree.heading("RewardCount", text="보상 개수")
        self.detail_tree.column("RewardCount", width=80, anchor="center")
        self.detail_tree.heading("RewardProbability", text="확률")
        self.detail_tree.column("RewardProbability", width=80, anchor="center")
        
        self.detail_tree.pack(fill="both", expand=True)
        detail_scroll_y.config(command=self.detail_tree.yview)
        
        # 상태 표시
        self.status_label = tk.Label(self.top, text="")
        self.status_label.pack(anchor="w", padx=10, pady=5)
        
        # 이벤트 연결
        self.box_tree.bind("<<TreeviewSelect>>", self._on_box_select)

    def _reset_search(self):
        """검색 결과를 초기화하고 전체 Box 목록을 다시 로드합니다."""
        # 검색어 초기화
        self.reward_id_entry.delete(0, tk.END)
        
        # 필터 초기화
        self.filter_var.set("")
        self.exclude_hash_var.set(False)
        
        # 전체 목록 다시 로드
        self.status_label.config(text="🔄 전체 Box 목록 다시 로드 중...")
        threading.Thread(target=self._load_all_boxes, daemon=True).start()


    def _remove_hash_from_selected(self):
        """선택한 Box의 엑셀 A열에서 #을 제거합니다."""
        selected_item = self.box_tree.focus()
        if not selected_item:
            messagebox.showwarning("선택 오류", "Box를 선택해주세요.")
            return
        
        values = self.box_tree.item(selected_item, "values")
        if not values or len(values) < 3:
            messagebox.showwarning("데이터 오류", "선택한 Box의 데이터가 유효하지 않습니다.")
            return
        
        item_id = values[0]  # ItemID
        print(f"[숨김 해제] ItemID {item_id} 엑셀 파일 검색 시작")
        
        # Box 관련 엑셀 파일 찾기 (BoxTemplate 또는 Box 관련 파일)
        excel_files_found = []
        
        for file, info in self.cache.items():
            # ItemTemplate 파일 제외
            if 'itemtemplate' in file.lower():
                print(f"[숨기기] ItemTemplate 파일 제외: {file}")
                continue
                
            if 'box' in file.lower() or 'boxtemplate' in file.lower():
                print(f"[숨기기] 가능성 있는 파일 발견: {file}")
                excel_files_found.append((file, info["path"]))
        
        if not excel_files_found:
            print(f"[숨김 해제] Box 관련 엑셀 파일을 찾을 수 없음")
            messagebox.showwarning("파일 없음", "Box 관련 엑셀 파일을 찾을 수 없습니다.")
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
                    
                    # ItemID/ItemTID 컬럼 찾기
                    item_id_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'itemid' in col_lower or 'item_id' in col_lower or 'itemtid' in col_lower or 'templateid' in col_lower:
                            item_id_col = col
                            print(f"[숨김 해제] ID 컬럼 발견: {col}")
                            break
                    
                    if not item_id_col:
                        print(f"[숨김 해제] ID 컬럼 없음, 다음 시트 확인")
                        continue
                    
                    # ItemID가 모두 숫자인지 확인
                    is_numeric = True
                    try:
                        df[item_id_col] = pd.to_numeric(df[item_id_col], errors='coerce')
                        df[item_id_col] = df[item_id_col].fillna(0).astype(int)
                        is_numeric = True
                    except:
                        is_numeric = False
                    
                    # 검색 방법 선택
                    if is_numeric:
                        print(f"[숨김 해제] 숫자형 ID 검색")
                        try:
                            numeric_item_id = int(item_id)
                            matched = df[df[item_id_col] == numeric_item_id]
                        except:
                            print(f"[숨김 해제] 숫자 변환 실패, 문자열 검색으로 전환")
                            df[item_id_col] = df[item_id_col].astype(str)
                            matched = df[df[item_id_col] == item_id]
                    else:
                        print(f"[숨김 해제] 문자열 ID 검색")
                        df[item_id_col] = df[item_id_col].astype(str)
                        matched = df[df[item_id_col] == item_id]
                    
                    print(f"[숨김 해제] 검색 결과: {len(matched)}행")
                    
                    if not matched.empty:
                        total_matches_found += len(matched)
                        print(f"[숨김 해제] {file}/{sheet}에서 ItemID {item_id} 발견")
                        # A열에서 # 제거
                        from utils.excel_utils import ExcelFileManager
                        
                        result = ExcelFileManager.remove_hash_from_a_column(path, sheet, item_id, header_row=header)
                        # 수정: 이름과 상태 모두 업데이트
                        if result:
                            show_message(self.top, "info", "성공", f"ItemID {item_id}의 A열에서 #이 제거되었습니다.")
                            
                            # 목록에서도 # 제거와 상태 업데이트
                            current_name = values[1]
                            if current_name.startswith('#'):
                                new_name = current_name[1:]
                                self.box_tree.item(selected_item, values=(values[0], new_name, values[2], "사용중"))
                        
                        # 중요: 여기서 return 문 제거 (모든 파일 처리를 위해)
                except Exception as e:
                    print(f"[파일 검색 오류] {file} / {sheet}: {e}")
        
        # 모든 파일 검색 후 결과 처리
        if total_matches_found > 0:
            show_message(self.top, "info", "성공", f"ItemID {item_id}의 A열에서 # 제거 처리가 완료되었습니다. ({total_matches_found}개 항목, {total_files_processed}개 파일)")
            
            # UI 업데이트 - 선택한 항목의 이름 앞에서 # 제거
            current_name = values[1]
            if current_name.startswith('#'):
                new_name = current_name[1:]
                self.box_tree.item(selected_item, values=(values[0], new_name, values[2]))
        else:
            show_message(self.top, "warning", "항목 없음", f"ItemID {item_id}에 해당하는 데이터를 엑셀 파일에서 찾을 수 없습니다.")


    def _add_hash_to_selected(self):
        """선택한 Box의 엑셀 A열에 #을 추가합니다."""
        selected_item = self.box_tree.focus()
        if not selected_item:
            messagebox.showwarning("선택 오류", "Box를 선택해주세요.")
            return
        
        values = self.box_tree.item(selected_item, "values")
        if not values or len(values) < 3:
            messagebox.showwarning("데이터 오류", "선택한 Box의 데이터가 유효하지 않습니다.")
            return
        
        item_id = values[0]  # ItemID
        print(f"[숨기기] ItemID {item_id} 엑셀 파일 검색 시작")
        
        # Box 관련 엑셀 파일 찾기 (BoxTemplate 또는 Box 관련 파일)
        excel_files_found = []
        
        for file, info in self.cache.items():
            # ItemTemplate 파일 제외
            if 'itemtemplate' in file.lower():
                print(f"[숨기기] ItemTemplate 파일 제외: {file}")
                continue
                
            if 'box' in file.lower() or 'boxtemplate' in file.lower():
                print(f"[숨기기] 가능성 있는 파일 발견: {file}")
                excel_files_found.append((file, info["path"]))
        
        if not excel_files_found:
            print(f"[숨기기] Box 관련 엑셀 파일을 찾을 수 없음")
            messagebox.showwarning("파일 없음", "Box 관련 엑셀 파일을 찾을 수 없습니다.")
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
                    
                    # ItemID/ItemTID 컬럼 찾기
                    item_id_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'itemid' in col_lower or 'item_id' in col_lower or 'itemtid' in col_lower or 'templateid' in col_lower:
                            item_id_col = col
                            print(f"[숨기기] ID 컬럼 발견: {col}")
                            break
                    
                    if not item_id_col:
                        print(f"[숨기기] ID 컬럼 없음, 다음 시트 확인")
                        continue
                    
                    # ItemID가 모두 숫자인지 확인
                    is_numeric = True
                    try:
                        df[item_id_col] = pd.to_numeric(df[item_id_col], errors='coerce')
                        df[item_id_col] = df[item_id_col].fillna(0).astype(int)
                        is_numeric = True
                    except:
                        is_numeric = False
                    
                    # 검색 방법 선택
                    if is_numeric:
                        print(f"[숨기기] 숫자형 ID 검색")
                        try:
                            numeric_item_id = int(item_id)
                            matched = df[df[item_id_col] == numeric_item_id]
                        except:
                            print(f"[숨기기] 숫자 변환 실패, 문자열 검색으로 전환")
                            df[item_id_col] = df[item_id_col].astype(str)
                            matched = df[df[item_id_col] == item_id]
                    else:
                        print(f"[숨기기] 문자열 ID 검색")
                        df[item_id_col] = df[item_id_col].astype(str)
                        matched = df[df[item_id_col] == item_id]
                    
                    print(f"[숨기기] 검색 결과: {len(matched)}행")
                    
                    if not matched.empty:
                        total_matches_found += len(matched)
                        print(f"[숨기기] {file}/{sheet}에서 ItemID {item_id} 발견")
                        # A열에 # 추가 - 여기서 header_row 값 전달
                        from utils.excel_utils import ExcelFileManager
                        
                        result = ExcelFileManager.add_hash_to_a_column(path, sheet, item_id, header_row=header)
                        # 수정 성공 시 업데이트 부분
                        if result:
                            show_message(self.top, "info", "성공", f"ItemID {item_id}의 A열에 #이 추가되었습니다.")
                            
                            # 개별 항목만 업데이트 (전체 리스트 리프레시 없음)
                            self._refresh_selected_item(item_id, is_hidden=True)
                            
                            # 목록에서도 # 표시와 상태 업데이트 (트리뷰 전체를 새로고침하지 않음)
                            current_name = values[1]
                            if not current_name.startswith('#'):
                                new_name = f"#{current_name}"
                                self.box_tree.item(selected_item, values=(values[0], new_name, values[2], "숨겨짐"))
                            else:
                                # 이미 #이 있으면 상태만 업데이트
                                self.box_tree.item(selected_item, values=(values[0], current_name, values[2], "숨겨짐"))
                                
                except Exception as e:
                    print(f"[파일 검색 오류] {file} / {sheet}: {e}")
        
        # 모든 파일 검색 후 결과 처리
        if total_matches_found > 0:
            show_message(self.top, "info", "성공", f"ItemID {item_id}의 A열에 # 추가 처리가 완료되었습니다. ({total_matches_found}개 항목, {total_files_processed}개 파일)")
            
            # UI 업데이트 - 선택한 항목의 이름 앞에 # 추가
            current_name = values[1]
            if not current_name.startswith('#'):
                new_name = f"#{current_name}"
                self.box_tree.item(selected_item, values=(values[0], new_name, values[2]))
        else:
            show_message(self.top, "warning", "항목 없음", f"ItemID {item_id}에 해당하는 데이터를 엑셀 파일에서 찾을 수 없습니다.")


    def _apply_name_filter(self):
        """# 필터 체크박스에 따라 필터링을 적용합니다."""
        # 체크 상태 확인
        exclude_hash = self.exclude_hash_var.get()
        
        # 모든 항목을 원래 상태로 복원
        self._restore_tree()
        
        if exclude_hash:
            # 필터 적용 - '#'으로 시작하는 이름 제외
            to_detach = []
            for item in self.box_tree.get_children():
                values = self.box_tree.item(item, "values")
                if len(values) >= 2 and isinstance(values[1], str) and values[1].startswith('#'):
                    to_detach.append(item)
                    self.box_tree.detach(item)
                    self._detached_items.append(item)  # 분리한 항목 저장
            
            self.status_label.config(text=f"✅ '#'로 시작하는 이름 제외 필터 적용 ({len(to_detach)}개 항목 숨김)")
        else:
            # 필터 초기화
            self.status_label.config(text="✅ 모든 항목 표시")
    
    def _restore_tree(self):
        """트리뷰에서 숨겨진 항목을 모두 복원합니다."""
        # 저장된 분리 항목들을 다시 붙이기
        for item in self._detached_items:
            self.box_tree.reattach(item, "", "end")
        # 분리 항목 리스트 초기화
        self._detached_items.clear()

    def _load_all_boxes(self):
        """모든 Box 정보를 불러옵니다."""
        self.status_label.config(text="🔍 전체 Box 목록 로딩 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=self._load_boxes_thread, daemon=True).start()


    def _load_boxes_thread(self):
        """백그라운드 스레드에서 Box 정보를 로드합니다."""
        try:
            self.top.after(0, lambda: self.status_label.config(text="🔍 전체 Box 목록 로딩 중..."))
            print("전체 Box 목록 로드 시작")
            
            all_boxes = []
            
            # 간소화된 방식으로 엑셀 파일에서 데이터 로드
            try:
                self._load_from_excel_simplified(all_boxes)
                print(f"엑셀에서 Box 정보 로드 완료: {len(all_boxes)}개")
            except Exception as excel_e:
                logger.error(f"엑셀 로드 오류: {excel_e}")
                print(f"엑셀 로드 오류: {excel_e}")
                
                # 여기서는 DB 로드 시도를 제거하고 오류 표시만 함
                self.top.after(0, lambda: self.status_label.config(
                    text=f"❌ 엑셀 로드 오류: {str(excel_e)}"))
                return
            
            # ID 기준으로 중복 제거
            unique_boxes = {}
            for box in all_boxes:
                if len(box) >= 3 and box[0] not in unique_boxes:  # ID 유효성 검사 추가
                    unique_boxes[box[0]] = box
            
            print(f"중복 제거 후 Box 개수: {len(unique_boxes)}개")
            
            # 트리뷰 업데이트
            self.top.after(0, lambda: self.box_tree.delete(*self.box_tree.get_children()))

            # ID 기준 정렬
            sorted_boxes = sorted(unique_boxes.values(), key=lambda x: x[0])
            
            # 디버깅: 첫 몇개 아이템 출력
            if sorted_boxes:
                print("첫 5개 아이템 샘플:")
                for i, box in enumerate(sorted_boxes[:5]):
                    print(f"  {i+1}. {box}")
            
            # 트리뷰에 추가 - lambda 안에서 올바르게 작동하도록 수정
            for idx, box in enumerate(sorted_boxes):
                # 상태 정보가 있는지 확인
                box_with_status = box
                if len(box) < 4:
                    box_with_status = box + ("사용중",)
                    
                # box[0]은 ID, box[1]은 이름, box[2]는 타입, box[3]은 상태
                # 트리뷰에 아이템 추가 (인덱스를 사용하여 항목 구분)
                box_idx = idx  # 고유한 인덱스
                box_values = box_with_status  # 표시할 값들
                self.top.after(0, lambda idx=box_idx, values=box_values: 
                            self.box_tree.insert("", "end", iid=f"box_{idx}", values=values))
            
            # 로딩 완료 메시지
            self.top.after(0, lambda: self.status_label.config(
                text=f"✅ 전체 Box 목록 로딩 완료: {len(sorted_boxes)}개"))
            
        except Exception as e:
            error_msg = f"❌ 오류 발생: {str(e)}"
            logger.error(f"Box 로드 오류: {e}")
            print(f"Box 로드 오류: {e}")
            self.top.after(0, lambda: self.status_label.config(text=error_msg))

    def _load_from_excel_simplified(self, all_boxes):
        """Excel 파일에서 Box 정보를 간소화된 방식으로 로드합니다."""
        box_files_count = 0
        box_template_items = {}  # ItemTID를 키로 하는 딕셔너리 (이름 매핑용)
        
        # 먼저 BoxTemplate 파일만 찾기
        box_template_files = []
        
        for file, info in self.cache.items():
            file_lower = file.lower()
            
            # BoxTemplate 파일만 처리
            if 'boxtemplate' in file_lower:
                box_template_files.append((file, info))
                box_files_count += 1
                continue
                
            # BoxTemplate은 아니지만 Box가 이름에 있는 파일 확인
            if 'box' in file_lower and not any(skip in file_lower for skip in ['jukebox', 'model']):
                # 시트 이름도 확인
                for sheet in info.get("sheets", {}):
                    if 'box' in sheet.lower():
                        box_template_files.append((file, info))
                        box_files_count += 1
                        break
        
        print(f"Box 정보 로드할 파일 {len(box_template_files)}개 찾음")
        
        # BoxTemplate 파일이 없으면 ItemTemplate에서 보조적으로 검색
        if len(box_template_files) == 0:
            for file, info in self.cache.items():
                if 'itemtemplate' in file.lower() and 'box' in file.lower():
                    box_template_files.append((file, info))
                    box_files_count += 1
        
        # 파일 처리
        for file, info in box_template_files:
            path = info["path"]
            print(f"Box 파일 처리: {file}")
            
            for sheet, meta in info.get("sheets", {}).items():
                # BoxTemplate 관련 시트만 처리
                if 'box' not in sheet.lower() and 'boxtemplate' not in sheet.lower():
                    continue
                    
                print(f"  시트 처리: {sheet}")
                try:
                    # 헤더 행 정보 가져오기
                    header = meta.get("header_row", 0)
                    
                    # 엑셀 파일 읽기
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # 필요한 컬럼 찾기
                    item_id_col = None
                    name_col = None
                    box_type_col = None
                    
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'itemid' in col_str or 'item_id' in col_str or 'itemtid' in col_str or 'templateid' in col_str:
                            item_id_col = col
                        elif 'name' in col_str or 'desc' in col_str:
                            name_col = col
                        elif 'boxtype' in col_str or 'box_type' in col_str:
                            box_type_col = col
                    
                    if not item_id_col:
                        print(f"  ItemID 컬럼을 찾을 수 없음: {sheet}")
                        continue
                    
                    # A열 상태 확인 시도
                    try:
                        # A열 가져오기 (첫 번째 열)
                        a_col = None
                        for col in df.columns:
                            if col == 0 or col == '#' or str(col) == '#':
                                a_col = col
                                break
                        
                        has_a_col = a_col is not None
                    except:
                        has_a_col = False
                    
                    # 모든 Box 항목 추가
                    for idx, row in df.iterrows():
                        try:
                            if pd.notna(row[item_id_col]):
                                # ItemID 정리 (소수점 제거)
                                item_tid = str(row[item_id_col]).split('.')[0]
                                
                                # ID가 숫자로 시작하는지 확인 (추가 필터링)
                                if not item_tid[0].isdigit():
                                    continue
                                
                                # 이름 가져오기 (일단 빈 값으로)
                                item_name = ""
                                if name_col and pd.notna(row[name_col]):
                                    item_name = str(row[name_col])
                                
                                # 타입 가져오기
                                box_type = ""
                                if box_type_col and pd.notna(row[box_type_col]):
                                    box_type = str(row[box_type_col])
                                else:
                                    # BoxTemplate 파일이면 기본값 설정
                                    if 'boxtemplate' in file.lower():
                                        box_type = "Box"
                                
                                # 상태 결정
                                status = "사용중"  # 기본값
                                
                                # A열 확인 (있는 경우)
                                if has_a_col and pd.notna(row[a_col]):
                                    a_value = str(row[a_col])
                                    if a_value.startswith('#'):
                                        status = "숨겨짐"
                                
                                # Box 정보 저장
                                box_template_items[item_tid] = (item_tid, item_name, box_type, status)
                        except Exception as row_error:
                            logger.warning(f"행 {idx} 처리 오류: {row_error}")
                            continue
                            
                except Exception as sheet_error:
                    logger.warning(f"시트 {sheet} 처리 오류: {sheet_error}")
                    continue
        
        # ItemTemplate에서 이름 정보 가져오기
        if box_template_items:
            print(f"ItemTemplate에서 {len(box_template_items)}개 Box 이름 조회 중...")
            self._get_box_names_from_item_template(box_template_items)
        
        # 결과 반환
        all_boxes.extend(box_template_items.values())
        
        if box_files_count == 0:
            logger.warning("Box 관련 엑셀 파일을 찾을 수 없음")

    def _get_box_names_from_item_template(self, box_items):
        """ItemTemplate에서 Box 이름을 가져옵니다."""
        item_template_files = []
        
        # ItemTemplate 파일 찾기
        for file, info in self.cache.items():
            file_lower = file.lower()
            if 'itemtemplate' in file_lower:
                item_template_files.append((file, info))
        
        if not item_template_files:
            print("ItemTemplate 파일을 찾을 수 없음")
            return
        
        print(f"ItemTemplate 파일 {len(item_template_files)}개 발견")
        
        # ItemTemplate 파일에서 이름 검색
        for file, info in item_template_files:
            path = info["path"]
            print(f"ItemTemplate 파일 처리: {file}")
            
            for sheet, meta in info.get("sheets", {}).items():
                try:
                    header = meta.get("header_row", 0)
                    
                    # 컬럼이 많은 파일에서는 필요한 컬럼만 로드
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # 필요한 컬럼 찾기
                    item_id_col = None
                    name_col = None
                    
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'templateid' in col_str or 'itemid' in col_str:
                            item_id_col = col
                        elif 'name' in col_str:
                            name_col = col
                    
                    if not (item_id_col and name_col):
                        continue
                    
                    # A열 확인 (상태 결정용)
                    # A열 확인 (상태 결정용)
                    a_col = None
                    for col in df.columns:
                        if col == 0 or col == '#' or str(col) == '#':
                            a_col = col
                            break
                    has_a_col = a_col is not None
                    
                    # 추가 디버깅 로그
                    if has_a_col:
                        print(f"A열 발견: {a_col}")
                        print(f"첫 5개 행 A열 값 샘플:")
                        for i in range(min(5, len(df))):
                            print(f"  행 {i+1}: '{df.iloc[i][a_col]}'")
                    
                    # A열 값에 따른 상태 결정 로직 개선
                    for idx, row in df.iterrows():
                        if pd.notna(row[item_id_col]) and pd.notna(row[name_col]):
                            item_id = str(row[item_id_col]).split('.')[0]
                            item_name = str(row[name_col])
                            
                            # A열 확인하여 상태 결정
                            is_hidden = False
                            if has_a_col and pd.notna(row[a_col]):
                                a_value = str(row[a_col])
                                if a_value.startswith('#'):
                                    is_hidden = True
                                    # 상태가 "숨겨짐"으로 설정
                                    if item_id in box_items:
                                        # 이름에 #이 없으면 추가
                                        if not item_name.startswith('#'):
                                            item_name = f"#{item_name}"
                                        
                                        # 기존 타입 정보 유지하면서 업데이트
                                        _, _, box_type, _ = box_items[item_id]
                                        box_items[item_id] = (item_id, item_name, box_type, "숨겨짐")
                                        print(f"  숨겨진 Box 발견: {item_id} -> {item_name}")
                    # Box ID별 이름 매핑
                    for idx, row in df.iterrows():
                        if pd.notna(row[item_id_col]) and pd.notna(row[name_col]):
                            item_id = str(row[item_id_col]).split('.')[0]
                            item_name = str(row[name_col])
                            
                            # A열 확인하여 상태 결정
                            is_hidden = False
                            if has_a_col and pd.notna(row[a_col]):
                                a_value = str(row[a_col])
                                if a_value.startswith('#'):
                                    is_hidden = True
                                    # 이름에 #이 없으면 추가
                                    if not item_name.startswith('#'):
                                        item_name = f"#{item_name}"
                            
                            # 이미 수집된 Box ID와 일치하면 이름 업데이트
                            if item_id in box_items:
                                _, _, box_type, status = box_items[item_id]
                                
                                # 상태 확인 - A열에 #이 있으면 숨겨짐으로 설정
                                if is_hidden:
                                    status = "숨겨짐"
                                
                                # 업데이트된 정보 저장
                                box_items[item_id] = (item_id, item_name, box_type, status)
                                print(f"  Box 이름 업데이트: {item_id} -> {item_name}")
                
                except Exception as e:
                    print(f"ItemTemplate 시트 처리 오류: {sheet} - {e}")

    def _refresh_selected_item(self, item_id, is_hidden=False):
        """선택한 항목만 리프레시합니다."""
        # 모든 항목 확인
        for item in self.box_tree.get_children():
            values = self.box_tree.item(item, "values")
            if values and values[0] == item_id:
                # 현재 값 가져오기
                current_name = values[1] if len(values) > 1 else ""
                current_type = values[2] if len(values) > 2 else ""
                
                # 숨김 상태에 따라 이름 및 상태 업데이트
                if is_hidden:
                    # 이름에 #이 없으면 추가
                    if not current_name.startswith('#'):
                        current_name = f"#{current_name}"
                    status = "숨겨짐"
                else:
                    # 이름에서 #이 있으면 제거
                    if current_name.startswith('#'):
                        current_name = current_name[1:]
                    status = "사용중"
                
                # 선택 항목 업데이트
                self.box_tree.item(item, values=(item_id, current_name, current_type, status))
                
                # 상세 정보도 새로고침 (선택한 항목 유지)
                self._on_box_select(None)
                break
            
    def _on_box_select(self, event):
        """Box 선택 시 상세 정보를 표시합니다."""
        selected_item = self.box_tree.focus()
        if not selected_item:
            return
        
        values = self.box_tree.item(selected_item, "values")
        if not values:
            return
        
        item_id = values[0]
        
        # 상세 정보 로드
        self.status_label.config(text=f"🔍 ItemID {item_id} 상세 정보 로딩 중...")
        
        # 백그라운드에서 실행
        threading.Thread(target=lambda: self._load_box_detail(item_id), daemon=True).start()

    def _load_box_detail(self, item_id):
        """선택한 Box의 상세 정보를 로드합니다."""
        try:
            # 상세 정보 트리뷰 초기화
            self.top.after(0, lambda: self.detail_tree.delete(*self.detail_tree.get_children()))
            
            # 1. DB에서 검색
            db_path = os.path.join(self.db_folder, "BoxTemplate.db")
            if os.path.exists(db_path):
                found = self._load_detail_from_db(db_path, item_id)
                if found:
                    return
            
            # 2. BoxTemplate Excel 파일에서 검색
            box_files = []
            for file, info in self.cache.items():
                file_lower = file.lower()
                
                # BoxTemplate 파일 먼저 찾기
                if 'boxtemplate' in file_lower:
                    box_files.append((file, info))
                    continue
                    
                # 또는 Box가 포함된 파일도 확인
                if 'box' in file_lower and not 'jukebox' in file_lower:
                    box_files.append((file, info))
            
            print(f"상세 정보 검색할 Box 파일: {len(box_files)}개")
            
            for file, info in box_files:
                found = self._load_detail_from_excel_file(file, info, item_id)
                if found:
                    return
                    
            # 3. 위에서 찾지 못했으면 실패 메시지
            self.top.after(0, lambda: self.status_label.config(
                text=f"⚠️ ItemID={item_id}에 해당하는 상세 정보를 찾을 수 없습니다."))
            
        except Exception as e:
            error_msg = f"❌ 상세 정보 로드 오류: {str(e)}"
            self.top.after(0, lambda: self.status_label.config(text=error_msg))

    def _load_detail_from_excel_file(self, file, info, item_id):
        """특정 엑셀 파일에서 Box 상세 정보를 로드합니다."""
        path = info["path"]
        
        # 박스 상세 정보는 BoxTemplate 관련 시트에만 있음
        for sheet, meta in info.get("sheets", {}).items():
            if 'box' not in sheet.lower() and 'boxtemplate' not in sheet.lower():
                continue
                
            try:
                header = meta["header_row"]
                
                # 엑셀 파일 읽기
                df = pd.read_excel(path, sheet_name=sheet, header=header)
                
                # 필요한 컬럼 찾기
                item_id_col = None
                reward_type_col = None
                reward_id_col = None
                reward_count_col = None
                reward_prob_col = None
                
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'itemid' in col_lower or 'item_id' in col_lower or 'itemtid' in col_lower:
                        item_id_col = col
                    elif 'rewardtype' in col_lower or 'reward_type' in col_lower:
                        reward_type_col = col
                    elif 'rewardid' in col_lower or 'reward_id' in col_lower:
                        reward_id_col = col
                    elif 'count' in col_lower:
                        reward_count_col = col
                    elif 'prob' in col_lower or 'rate' in col_lower or 'chance' in col_lower:
                        reward_prob_col = col
                
                if not (item_id_col and reward_id_col):
                    continue
                    
                # ID 검색
                try:
                    # 문자열 및 숫자 비교를 위해 두 가지 방식 모두 시도
                    id_match = df[df[item_id_col].astype(str) == str(item_id)]
                    
                    if id_match.empty:
                        # 숫자 변환 시도
                        df[item_id_col] = pd.to_numeric(df[item_id_col], errors='coerce')
                        id_match = df[df[item_id_col] == float(item_id)]
                except:
                    # 오류 시 문자열 비교만 시도
                    df[item_id_col] = df[item_id_col].astype(str)
                    id_match = df[df[item_id_col] == str(item_id)]
                
                if id_match.empty:
                    continue
                    
                # 결과 발견
                print(f"상세 정보 발견: {file}/{sheet}, {len(id_match)}행")
                
                # 결과 추가
                for _, row in id_match.iterrows():
                    reward_type = row[reward_type_col] if reward_type_col else ""
                    reward_id = row[reward_id_col] if reward_id_col else ""
                    reward_count = row[reward_count_col] if reward_count_col else ""
                    reward_prob = row[reward_prob_col] if reward_prob_col else ""
                    
                    # 보상 이름 조회
                    reward_name = self._resolve_reward_info(reward_type, reward_id)
                    
                    values = (
                        reward_type, reward_id, reward_name,
                        reward_count, reward_prob
                    )
                    
                    self.top.after(0, lambda v=values: self.detail_tree.insert("", "end", values=v))
                
                self.top.after(0, lambda: self.status_label.config(
                    text=f"✅ ItemID={item_id} 상세 정보 로드 완료 ({len(id_match)}개 항목)"))
                    
                return True
                
            except Exception as e:
                print(f"[Excel 상세 정보 로드 오류] {file} / {sheet}: {e}")
        
        return False


    def _load_from_db(self, db_path, all_boxes):
        """DB에서 모든 Box 정보를 로드합니다."""
        try:
            logger.debug(f"[BoxList] DB 연결: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # 컬럼 확인
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(BoxTemplate)")
            columns = [row[1] for row in cursor.fetchall()]
            logger.debug(f"[BoxList] DB 컬럼 목록: {columns}")
            
            # 필요한 컬럼 찾기
            item_id_col = None
            name_col = None
            box_type_col = None
            
            for col in columns:
                col_lower = col.lower()
                if 'itemtid' in col_lower or 'item_id' in col_lower or 'itemid' in col_lower:
                    item_id_col = col
                elif 'name' in col_lower or 'desc' in col_lower:
                    name_col = col
                elif 'boxtype' in col_lower or 'box_type' in col_lower:
                    box_type_col = col
            
            logger.debug(f"[BoxList] 매핑된 컬럼: ItemID={item_id_col}, Name={name_col}, BoxType={box_type_col}")
            if not item_id_col:
                logger.error(f"[BoxList] BoxTemplate DB에 ItemTID/ItemID 컬럼이 없습니다.")
                self.top.after(0, lambda: self.status_label.config(
                    text="❌ BoxTemplate DB에 ItemTID/ItemID 컬럼이 없습니다."))
                conn.close()
                return
            
            # 모든 Box 조회 (ItemTID만 필요)
            query = f"SELECT DISTINCT {item_id_col} FROM BoxTemplate"
            logger.debug(f"[BoxList] 쿼리 실행: {query}")
            df = pd.read_sql_query(query, conn)
            conn.close()
            
            logger.debug(f"[BoxList] DB에서 Box 정보 로드 완료: {len(df)}행")
            
            # 결과 목록에 추가 - 각 ItemTID에 대해 ItemTemplate에서 정보 조회
            for idx, row in df.iterrows():
                item_tid = str(row.iloc[0])
                # ItemTemplate에서 정보 조회
                item_info = self.get_item_info(item_tid)
                all_boxes.append((
                    item_tid,                   # ItemID
                    item_info["name"],          # Box 이름 (ItemTemplate에서)
                    item_info["type"]           # Box 타입 (ItemTemplate에서)
                ))
                    
        except Exception as e:
            logger.error(f"[BoxList] DB 로드 오류: {str(e)}")


    def _load_from_excel(self, all_boxes):
        """Excel 파일에서 모든 Box 정보를 로드합니다."""
        for file, info in self.cache.items():
            if not ('box' in file.lower() or 'boxtemplate' in file.lower()):
                continue
                
            for sheet, meta in info.get("sheets", {}).items():
                try:
                    path, header = info["path"], meta["header_row"]
                    
                    # 엑셀 파일을 직접 열어서 A열도 함께 확인
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(path, read_only=True)
                        ws = wb[sheet]
                    except Exception as e:
                        print(f"[Excel A열 확인 실패] {path}/{sheet}: {e}")
                        # openpyxl 로드 실패 시 pandas로 시도
                        df = pd.read_excel(path, sheet_name=sheet, header=header)
                        has_openpyxl = False
                    else:
                        df = pd.read_excel(path, sheet_name=sheet, header=header)
                        has_openpyxl = True
                    
                    # 필요한 컬럼 찾기
                    item_id_col = None
                    name_col = None
                    box_type_col = None
                    
                    for col in df.columns:
                        col_lower = col.lower() if isinstance(col, str) else str(col).lower()
                        if 'itemid' in col_lower or 'item_id' in col_lower or 'itemtid' in col_lower or 'templateid' in col_lower:
                            item_id_col = col
                        elif 'name' in col_lower or 'desc' in col_lower:
                            name_col = col
                        elif 'boxtype' in col_lower or 'box_type' in col_lower or 'type' in col_lower:
                            box_type_col = col
                    
                    if not item_id_col:
                        continue
                        
                    # 중복 없이 모든 Box 추가
                    for idx, row in df.iterrows():
                        if pd.notna(row[item_id_col]):
                            item_tid = str(row[item_id_col])
                            item_name = str(row[name_col]) if name_col and pd.notna(row[name_col]) else ""
                            box_type = str(row[box_type_col]) if box_type_col and pd.notna(row[box_type_col]) else ""
                            
                            # A열 확인하여 상태 결정
                            status = "사용중"  # 기본값
                            if has_openpyxl:
                                try:
                                    # 실제 엑셀 행 번호 계산 (헤더 행 + 인덱스 + 1)
                                    actual_row = header + idx + 1
                                    a_cell = ws.cell(row=actual_row, column=1)
                                    if a_cell.value and isinstance(a_cell.value, str) and a_cell.value.startswith('#'):
                                        status = "숨겨짐"
                                        # 이름 앞에 #이 없으면 #을 추가 (UI 표시용)
                                        if not item_name.startswith('#'):
                                            item_name = f"#{item_name}"
                                except Exception as cell_err:
                                    print(f"[A열 확인 오류] 행 {actual_row}: {cell_err}")
                            else:
                                # openpyxl 실패 시 이름으로 상태 추정
                                if item_name.startswith('#'):
                                    status = "숨겨짐"
                            
                            all_boxes.append((
                                item_tid,    # ItemID
                                item_name,   # Box 이름
                                box_type,    # Box 타입
                                status       # 상태 (새로 추가)
                            ))
                            
                except Exception as e:
                    print(f"[Excel 로드 오류] {file} / {sheet}: {e}")
    
    def _on_box_select(self, event):
        """Box 선택 시 상세 정보를 표시합니다."""
        selected_item = self.box_tree.focus()
        if not selected_item:
            return
        
        values = self.box_tree.item(selected_item, "values")
        if not values:
            return
        
        item_id = values[0]
        
        # 상세 정보 로드
        self.status_label.config(text=f"🔍 ItemID {item_id} 상세 정보 로딩 중...")
        
        # 백그라운드에서 실행
        threading.Thread(target=lambda: self._load_box_detail(item_id), daemon=True).start()
        
    def _load_box_detail(self, item_id):
        """선택한 Box의 상세 정보를 로드합니다."""
        try:
            # 상세 정보 트리뷰 초기화
            self.top.after(0, lambda: self.detail_tree.delete(*self.detail_tree.get_children()))
            
            # 1. DB에서 검색
            db_path = os.path.join(self.db_folder, "BoxTemplate.db")
            if os.path.exists(db_path):
                found = self._load_detail_from_db(db_path, item_id)
                if found:
                    return
            
            # 2. Excel 파일에서 검색
            self._load_detail_from_excel(item_id)
            
        except Exception as e:
            error_msg = f"❌ 상세 정보 로드 오류: {str(e)}"
            self.top.after(0, lambda: self.status_label.config(text=error_msg))


    def _load_detail_from_db(self, db_path, item_id):
        """DB에서 Box 상세 정보를 로드합니다."""
        try:
            logger.debug(f"[BoxList] DB 연결: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # 컬럼 확인
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(BoxTemplate)")
            columns = [row[1] for row in cursor.fetchall()]
            logger.debug(f"[BoxList] DB 컬럼 목록: {columns}")
            
            # 필요한 컬럼 찾기
            item_id_col = None
            reward_type_col = None
            reward_id_col = None
            reward_count_col = None
            reward_prob_col = None
            
            for col in columns:
                col_lower = col.lower()
                if 'itemtid' in col_lower or 'item_id' in col_lower or 'itemid' in col_lower:
                    item_id_col = col
                elif 'rewardtype' in col_lower or 'reward_type' in col_lower:
                    reward_type_col = col
                elif 'rewardid' in col_lower or 'reward_id' in col_lower:
                    reward_id_col = col
                elif 'min' in col_lower and 'count' in col_lower:
                    reward_count_col = col
                elif 'prob' in col_lower or 'rate' in col_lower or 'chance' in col_lower:
                    reward_prob_col = col
            
            logger.debug(f"[BoxList] 매핑된 컬럼: ItemID={item_id_col}, RewardType={reward_type_col}, RewardID={reward_id_col}")
            
            # 필수 컬럼은 ItemID/ItemTID와 RewardID만 체크
            if not item_id_col:
                logger.error("[BoxList] ItemTID/ItemID 컬럼을 찾을 수 없습니다.")
                raise Exception("ItemTID/ItemID 컬럼을 찾을 수 없습니다.")
                
            if not reward_id_col:
                logger.error("[BoxList] RewardID 컬럼을 찾을 수 없습니다.")
                raise Exception("RewardID 컬럼을 찾을 수 없습니다.")
            
            # ItemID로 상세 정보 조회 (RewardType이 없어도 조회는 가능)
            query = f"SELECT {item_id_col}, {reward_id_col}"
            
            if reward_type_col:
                query += f", {reward_type_col}"
            else:
                query += ", 0"  # 기본값
                
            if reward_count_col:
                query += f", {reward_count_col}"
            else:
                query += ", 0"
                    
            if reward_prob_col:
                query += f", {reward_prob_col}"
            else:
                query += ", 0"
                    
            query += f" FROM BoxTemplate WHERE {item_id_col} = ?"
            
            logger.debug(f"[BoxList] 쿼리 실행: {query}, ItemTID={item_id}")
            df = pd.read_sql_query(query, conn, params=(item_id,))
            conn.close()
            
            if df.empty:
                logger.debug(f"[BoxList] 해당 ItemTID에 대한 상세 정보가 없습니다: {item_id}")
                self.top.after(0, lambda: self.status_label.config(
                    text=f"⚠️ ItemTID={item_id}에 해당하는 상세 정보가 없습니다."))
                return False
            
            logger.debug(f"[BoxList] DB에서 상세 정보 로드 성공: {len(df)}행")
            
            # RewardName 조회 및 트리뷰에 추가
            for idx, row in df.iterrows():
                # itemid_col 위치와 reward_id_col 위치를 확인하여 접근
                col_index = {
                    'item_id': 0,
                    'reward_id': 1,
                    'reward_type': 2 if reward_type_col else None,
                    'reward_count': 3 if reward_count_col else None,
                    'reward_prob': 4 if reward_prob_col else None
                }
                
                item_id_val = row.iloc[col_index['item_id']]
                reward_id_val = row.iloc[col_index['reward_id']]
                reward_type_val = row.iloc[col_index['reward_type']] if col_index['reward_type'] is not None else 0
                reward_count_val = row.iloc[col_index['reward_count']] if col_index['reward_count'] is not None else 0
                reward_prob_val = row.iloc[col_index['reward_prob']] if col_index['reward_prob'] is not None else 0
                
                # RewardType에 따른 RewardID 이름 조회
                logger.debug(f"[BoxList] RewardInfo 조회: Type={reward_type_val}, ID={reward_id_val}")
                reward_name = self._resolve_reward_info(reward_type_val, reward_id_val)
                
                values = (
                    reward_type_val, reward_id_val, reward_name,
                    reward_count_val, reward_prob_val
                )
                
                self.top.after(0, lambda v=values: self.detail_tree.insert("", "end", values=v))
            
            self.top.after(0, lambda: self.status_label.config(
                text=f"✅ ItemTID={item_id} 상세 정보 로드 완료 ({len(df)}개 항목)"))
            return True
                
        except Exception as e:
            logger.error(f"[BoxList] DB 상세 정보 로드 오류: {str(e)}")
            return False
                    
     
    def _load_detail_from_excel(self, item_id):
        """Excel 파일에서 Box 상세 정보를 로드합니다."""
        found = False
        
        for file, info in self.cache.items():
            if found:
                break
                
            for sheet, meta in info.get("sheets", {}).items():
                if not ('box' in sheet.lower() or 'boxtemplate' in sheet.lower()):
                    continue
                    
                try:
                    path, header = info["path"], meta["header_row"]
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # 필요한 컬럼 찾기
                    item_id_col = None
                    reward_type_col = None
                    reward_id_col = None
                    reward_count_col = None
                    reward_prob_col = None
                    
                    for col in df.columns:
                        col_lower = col.lower()
                        if 'itemid' in col_lower or 'item_id' in col_lower:
                            item_id_col = col
                        elif 'rewardtype' in col_lower or 'reward_type' in col_lower:
                            reward_type_col = col
                        elif 'rewardid' in col_lower or 'reward_id' in col_lower:
                            reward_id_col = col
                        elif 'count' in col_lower:
                            reward_count_col = col
                        elif 'prob' in col_lower or 'rate' in col_lower or 'chance' in col_lower:
                            reward_prob_col = col
                    
                    if not (item_id_col and reward_type_col and reward_id_col):
                        continue
                        
                    # ItemID 컬럼 데이터 변환
                    df[item_id_col] = df[item_id_col].astype(str)
                    
                    # 해당 ItemID 조회
                    matched = df[df[item_id_col] == item_id]
                    
                    if not matched.empty:
                        found = True
                        
                        # 결과 트리뷰에 추가
                        for _, row in matched.iterrows():
                            reward_type = row[reward_type_col]
                            reward_id = row[reward_id_col]
                            reward_count = row[reward_count_col] if reward_count_col else 0
                            reward_prob = row[reward_prob_col] if reward_prob_col else 0
                            
                            # RewardType에 따른 RewardID 이름 조회
                            reward_name = self._resolve_reward_info(reward_type, reward_id)
                            
                            values = (
                                reward_type, reward_id, reward_name,
                                reward_count, reward_prob
                            )
                            
                            self.top.after(0, lambda v=values: self.detail_tree.insert("", "end", values=v))
                        
                        self.top.after(0, lambda: self.status_label.config(
                            text=f"✅ ItemID={item_id} 상세 정보 로드 완료 ({len(matched)}개 항목)"))
                        break
                        
                except Exception as e:
                    print(f"[Excel 상세 정보 로드 오류] {file} / {sheet}: {str(e)}")
        
        if not found:
            self.top.after(0, lambda: self.status_label.config(
                text=f"⚠️ ItemID={item_id}에 해당하는 상세 정보를 찾을 수 없습니다."))

    def _apply_filter(self):
        """트리뷰에 필터를 적용합니다."""
        keyword = self.filter_var.get().strip().lower()
        if not keyword:
            return
            
        # 항목 숨기기 전 모든 항목을 복원
        self._restore_tree()
        
        # 항목 숨기기
        hidden_count = 0
        for item in self.box_tree.get_children():
            values = self.box_tree.item(item, "values")
            # 어떤 컬럼이든 키워드를 포함하는지 확인
            if not any(keyword in str(v).lower() for v in values):
                self.box_tree.detach(item)
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
        for item in self.box_tree.get_children():
            values = self.box_tree.item(item, "values")
            if len(values) >= 2 and isinstance(values[1], str) and values[1].startswith('#'):
                self.box_tree.detach(item)
                hash_hidden += 1
        
        return hash_hidden

    def _clear_filter(self):
        """트리뷰 필터를 초기화합니다."""
        # 필터 텍스트 초기화
        self.filter_var.set("")
        
        # 체크박스도 초기화
        self.exclude_hash_var.set(False)
        
        # 숨겨진 항목 모두 복원
        self._restore_tree()
                
        # 상태 업데이트
        self.status_label.config(text=f"✅ 필터 초기화 완료")


    def _resolve_reward_info(self, reward_type, reward_id):
        """RewardType에 따라 RewardID의 이름을 조회합니다."""
        # 바로 fallback 함수 호출
        return self._resolve_reward_info_fallback(reward_type, reward_id)
        
            
    # RewardType이 매핑에 있는지 확인
    def _resolve_reward_info_fallback(self, reward_type, reward_id):
        """타입코드 매핑 없이 직접 RewardType에 따라 이름을 조회합니다."""
        try:
            print(f"[RewardInfo] 대체 조회 시작: Type={reward_type}, ID={reward_id}")
            
            # RewardType에 따른 타겟 테이블과 컬럼 직접 매핑
            type_mapping = {
                10: {"table": "HeroTemplate", "column": "BaseHeroID"},
                11: {"table": "HeroTemplate", "column": "BaseHeroID"},
                20: {"table": "ItemTemplate", "column": "TemplateID"},
                21: {"table": "ItemTemplate", "column": "TemplateID"},
                30: {"table": "GoodsMaxValue", "column": "GoodsType"},
                40: {"table": "BoxTemplate", "column": "ItemTID"},
                41: {"table": "BoxTemplate", "column": "ItemTID"},
                50: {"table": "TicketMaxValue", "column": "TicketType"},
                70: {"table": "WisdomBookTemplate", "column": "TemplateID"},
                80: {"table": "CostumeTemplate", "column": "TemplateID"}
                # 여기에 더 많은 RewardType 매핑 추가
            }
            # RewardType이 매핑에 있는지 확인
            if isinstance(reward_type, str):
                try:
                    reward_type = float(reward_type)
                except:
                    print(f"[RewardInfo] RewardType 변환 오류: {reward_type}")
                    return "타입 변환 오류"
                    
            reward_type_int = int(float(reward_type))
            if reward_type_int not in type_mapping:
                print(f"[RewardInfo] 매핑 없음: Type={reward_type_int}")
                return f"RewardType {reward_type_int} 매핑 없음"
                
            # 대상 테이블과 컬럼 가져오기
            target = type_mapping[reward_type_int]
            table, column = target["table"], target["column"]
            print(f"[RewardInfo] 매핑 정보: 테이블={table}, 컬럼={column}")
            
            # DB 파일 확인
            db_path = os.path.join(self.db_folder, f"{table}.db")
            print(f"[RewardInfo] DB 경로: {db_path}, 존재={os.path.exists(db_path)}")
            
            if not os.path.exists(db_path):
                return f"{table} DB 없음"
                
            try:
                # 숫자형으로 변환
                if isinstance(reward_id, str) and reward_id.isdigit():
                    numeric_reward_id = int(reward_id)
                else:
                    numeric_reward_id = int(float(reward_id))
            except:
                print(f"[RewardInfo] RewardID 숫자 변환 실패: {reward_id}")
                # 숫자 변환 실패 시 원본 값 사용
                numeric_reward_id = reward_id
            
            # 두 가지 방식으로 검색 시도 (정확한 숫자 일치와 문자열 일치)
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 1. 컬럼 목록 확인
            cursor.execute(f"PRAGMA table_info({table})")
            columns_info = cursor.fetchall()
            column_names = [col[1] for col in columns_info]
            print(f"[RewardInfo] 테이블 컬럼: {column_names}")
            
            if column not in column_names:
                print(f"[RewardInfo] 경고: {column} 컬럼이 테이블에 없습니다. 대체 컬럼 검색 중...")
                # ID 관련 컬럼 자동 탐지
                for col_name in column_names:
                    if 'id' in col_name.lower() or 'template' in col_name.lower():
                        column = col_name
                        print(f"[RewardInfo] 대체 컬럼 발견: {column}")
                        break
            
            # 2. 정확한 숫자 일치로 검색
            try:
                query = f"SELECT * FROM {table} WHERE {column} = ?"
                print(f"[RewardInfo] 쿼리1: {query}, 값={numeric_reward_id}")
                cursor.execute(query, (numeric_reward_id,))
                rows = cursor.fetchall()
                
                if not rows:
                    # 3. 문자열 일치로 검색
                    query = f"SELECT * FROM {table} WHERE {column} = ?"
                    str_reward_id = str(reward_id)
                    print(f"[RewardInfo] 쿼리2: {query}, 값={str_reward_id}")
                    cursor.execute(query, (str_reward_id,))
                    rows = cursor.fetchall()
                    
                    if not rows:
                        # 4. LIKE 검색 (부분 일치)
                        query = f"SELECT * FROM {table} WHERE {column} LIKE ?"
                        like_pattern = f"%{str_reward_id}%"
                        print(f"[RewardInfo] 쿼리3: {query}, 값={like_pattern}")
                        cursor.execute(query, (like_pattern,))
                        rows = cursor.fetchall()
            except Exception as qe:
                print(f"[RewardInfo] 쿼리 실행 오류: {qe}")
                # 오류 발생 시 다른 방식으로 쿼리 실행
                try:
                    df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
                    print(f"[RewardInfo] 전체 테이블 조회: {len(df)} 행")
                    
                    # 데이터프레임에서 필터링
                    df[column] = df[column].astype(str)
                    str_reward_id = str(reward_id)
                    filtered = df[df[column].str.contains(str_reward_id)]
                    
                    if not filtered.empty:
                        print(f"[RewardInfo] DataFrame 필터링 성공: {len(filtered)} 행")
                        # 행 데이터를 rows 형식으로 변환
                        rows = [tuple(row) for _, row in filtered.iterrows()]
                except Exception as dfe:
                    print(f"[RewardInfo] DataFrame 처리 오류: {dfe}")
                    rows = []
            
            # 결과가 있으면 이름 가져오기
            if rows:
                print(f"[RewardInfo] 결과 행 수: {len(rows)}")
                # 컬럼 인덱스와 함께 데이터 변환
                cols_dict = {column_names[i]: i for i in range(len(column_names))}
                df = pd.DataFrame(rows, columns=column_names)
                
                # 이름 컬럼 찾기
                name_cols = ["Name", "DisplayName", "Title", "Description"]
                for col in name_cols:
                    if col in df.columns and not pd.isna(df.iloc[0][col]):
                        result = str(df.iloc[0][col])
                        print(f"[RewardInfo] 이름 찾음: {result} (컬럼: {col})")
                        conn.close()
                        return result
                
                # 이름 컬럼이 없으면 ID 반환
                first_row_id = df.iloc[0][column]
                conn.close()
                return f"{table}: {first_row_id}"
            else:
                print(f"[RewardInfo] 결과 없음: {table} 테이블에서 {reward_id} 찾을 수 없음")
                conn.close()
                return f"{table}에 없음"
                
        except Exception as e:
            import traceback
            print(f"[RewardInfo] 전체 처리 오류: {e}")
            print(traceback.format_exc())
            return "오류 발생"           


    def get_item_info(self, item_tid):
        """ItemTID를 이용해 ItemTemplate에서 아이템 정보를 가져옵니다."""
        try:
            print(f"[ItemInfo] {item_tid} 정보 조회 시작")
            
            # ItemTemplate DB 경로
            db_path = os.path.join(self.db_folder, "ItemTemplate.db")
            print(f"[ItemInfo] DB 경로: {db_path}, 존재={os.path.exists(db_path)}")
            
            if not os.path.exists(db_path):
                return {"name": "DB 없음", "type": "알 수 없음"}
                
            # DB 연결
            conn = sqlite3.connect(db_path)
            
            # 필요한 컬럼 직접 지정
            id_col = "TemplateID"  # 정확한 ID 컬럼명
            name_col = "Name"      # 정확한 이름 컬럼명
            type_col = "ItemType"  # 정확한 타입 컬럼명
            
            # 테이블 구조 확인
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(ItemTemplate)")
            column_info = cursor.fetchall()
            columns = [row[1] for row in column_info]
            print(f"[ItemInfo] 사용할 컬럼: ID={id_col}, 이름={name_col}, 타입={type_col}")
            
            # 필요한 컬럼이 존재하는지 확인
            if id_col not in columns or name_col not in columns or type_col not in columns:
                missing = []
                if id_col not in columns: missing.append(id_col)
                if name_col not in columns: missing.append(name_col)
                if type_col not in columns: missing.append(type_col)
                print(f"[ItemInfo] 필요한 컬럼이 없습니다: {missing}")
                return {"name": "컬럼 없음", "type": "알 수 없음"}
            
            # ItemTID를 숫자로 변환
            try:
                if isinstance(item_tid, str):
                    num_tid = int(item_tid)
                else:
                    num_tid = int(item_tid)
                print(f"[ItemInfo] 검색할 ID: {num_tid} (타입: {type(num_tid)})")
            except:
                print(f"[ItemInfo] ID 변환 실패: {item_tid}")
                return {"name": "ID 변환 오류", "type": "알 수 없음"}
            
            # 필요한 컬럼만 선택하여 조회
            query = f"SELECT {id_col}, {name_col}, {type_col} FROM ItemTemplate WHERE {id_col} = ?"
            print(f"[ItemInfo] 실행 쿼리: {query}, 파라미터: {num_tid}")
            
            cursor.execute(query, (num_tid,))
            result = cursor.fetchone()
            
            if not result:
                # ID가 문자열인 경우 시도
                cursor.execute(query, (str(num_tid),))
                result = cursor.fetchone()
                
            if not result:
                # ID 앞부분만 일치하는 경우 시도 (LIKE 검색)
                like_query = f"SELECT {id_col}, {name_col}, {type_col} FROM ItemTemplate WHERE {id_col} LIKE ?"
                cursor.execute(like_query, (f"{num_tid}%",))
                result = cursor.fetchone()
                
            if not result:
                # 전체 ID 목록 확인 (디버깅용)
                cursor.execute(f"SELECT {id_col} FROM ItemTemplate LIMIT 10")
                sample_ids = cursor.fetchall()
                print(f"[ItemInfo] 샘플 ID 목록: {sample_ids}")
                
                print(f"[ItemInfo] {num_tid}에 대한 결과 없음")
                conn.close()
                return {"name": "찾을 수 없음", "type": "알 수 없음"}
                
            # 결과 처리
            template_id = result[0]
            name = result[1]
            item_type = result[2]
            
            print(f"[ItemInfo] 조회 결과: ID={template_id}, 이름={name}, 타입={item_type}")
            
            # ItemType 처리
            type_display = "알 수 없음"
            try:
                if isinstance(item_type, (int, float)):
                    type_code = int(item_type)
                elif isinstance(item_type, str) and item_type.isdigit():
                    type_code = int(item_type)
                else:
                    type_code = 0
                    
                print(f"[ItemInfo] 변환된 타입 코드: {type_code}")
                    
                # 타입에 따른 표시 텍스트
                if type_code == 6000:
                    type_display = "랜덤박스"
                elif type_code == 6001:
                    type_display = "패키지"
                elif type_code == 6002:
                    type_display = "선택박스"
                else:
                    type_display = f"Type-{type_code}"
            except Exception as type_e:
                print(f"[ItemInfo] 타입 변환 오류: {type_e}")
                type_display = f"오류: {str(item_type)}"
                
            conn.close()
            return {"name": name, "type": type_display}
                
        except Exception as e:
            import traceback
            print(f"[ItemInfo] 조회 오류: {e}")
            print(traceback.format_exc())
            return {"name": f"오류: {str(e)[:30]}", "type": "알 수 없음"}

    def _search_by_reward_id(self):
        """보상 ID를 포함하는 Box를 검색합니다."""
        reward_id = self.reward_id_entry.get().strip()
        if not reward_id:
            messagebox.showwarning("입력 오류", "보상 ID를 입력하세요.")
            return
        
        self.status_label.config(text=f"🔍 보상 ID {reward_id} 검색 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=lambda: self._search_reward_id_thread(reward_id), daemon=True).start()

    def _search_reward_id_thread(self, reward_id):
        """백그라운드 스레드에서 보상 ID를 검색합니다."""
        try:
            found_boxes = []
            
            # 모든 Box 템플릿 파일 검색
            for file, info in self.cache.items():
                if 'box' in file.lower() or 'boxtemplate' in file.lower():
                    path = info["path"]
                    
                    for sheet, meta in info.get("sheets", {}).items():
                        try:
                            header = meta["header_row"]
                            
                            # 엑셀 파일 읽기
                            df = pd.read_excel(path, sheet_name=sheet, header=header)
                            
                            # RewardID 컬럼 찾기
                            reward_id_col = None
                            id_col = None
                            
                            for col in df.columns:
                                col_lower = str(col).lower()
                                if 'rewardid' in col_lower or 'reward_id' in col_lower:
                                    reward_id_col = col
                                elif 'itemid' in col_lower or 'item_id' in col_lower or 'templateid' in col_lower or 'itemtid' in col_lower:
                                    id_col = col
                            
                            if not (reward_id_col and id_col):
                                continue
                            
                            # 숫자형과 문자열 형 모두 검색
                            try:
                                # 숫자형 검색
                                numeric_reward_id = int(reward_id)
                                df[reward_id_col] = pd.to_numeric(df[reward_id_col], errors='coerce')
                                matched = df[df[reward_id_col] == numeric_reward_id]
                            except:
                                # 문자열 검색
                                df[reward_id_col] = df[reward_id_col].astype(str)
                                matched = df[df[reward_id_col] == str(reward_id)]
                            
                            # 결과 처리
                            if not matched.empty:
                                for _, row in matched.iterrows():
                                    box_id = str(row[id_col])
                                    # ItemID에서 소수점 제거
                                    box_id = self._format_item_id(box_id)
                                    
                                    # 이미 발견한 Box ID는 중복 추가하지 않음
                                    if not any(box[0] == box_id for box in found_boxes):
                                        # ItemTemplate에서 정보 조회
                                        item_info = self.get_item_info(box_id)
                                        found_boxes.append((
                                            box_id,                   # ItemID
                                            item_info["name"],        # Box 이름 (ItemTemplate에서)
                                            item_info["type"]         # Box 타입 (ItemTemplate에서)
                                        ))
                        
                        except Exception as e:
                            print(f"[검색 오류] {file} / {sheet}: {e}")
            
            # 결과 표시
            self.top.after(0, lambda: self._display_search_results(found_boxes, reward_id))
            
        except Exception as e:
            error_msg = f"❌ 검색 오류: {str(e)}"
            self.top.after(0, lambda: self.status_label.config(text=error_msg))


    def _format_item_id(self, item_id):
        """ItemID를 깔끔한 형태로 포맷팅합니다."""
        try:
            # 문자열로 변환
            str_id = str(item_id)
            
            # 소수점이 있고 소수점 이하가 0인 경우 제거
            if '.' in str_id:
                # 소수점으로 분리
                parts = str_id.split('.')
                if len(parts) == 2 and (parts[1] == '0' or parts[1] == '00'):
                    return parts[0]
            
            return str_id
        except:
            return str(item_id)


    def _display_search_results(self, found_boxes, reward_id):
        """검색 결과를 트리뷰에 표시합니다."""
        # 트리뷰 초기화
        self.box_tree.delete(*self.box_tree.get_children())
        
        # 결과가 없는 경우
        if not found_boxes:
            self.status_label.config(text=f"⚠️ 보상 ID {reward_id}를 포함하는 Box를 찾을 수 없습니다.")
            return
        
        # 결과 정렬 및 표시
        sorted_boxes = sorted(found_boxes, key=lambda x: x[0])
        for box in sorted_boxes:
            # ItemID 소수점 제거
            item_id = self._format_item_id(box[0])
            self.box_tree.insert("", "end", values=(item_id, box[1], box[2]))
        
        # 상태 업데이트
        self.status_label.config(text=f"✅ 보상 ID {reward_id} 검색 완료: {len(found_boxes)}개 Box 발견")

    def _add_hash_to_all_matching_rewards(self):
        """선택한 보상 항목의 ID와 일치하는 모든 데이터의 A열에 #을 추가합니다."""
        selected_item = self.detail_tree.focus()
        if not selected_item:
            show_message(self.top, "warning", "선택 오류", "보상 항목을 선택해주세요.")
            return
        
        values = self.detail_tree.item(selected_item, "values")
        if not values or len(values) < 2:
            show_message(self.top, "warning", "데이터 오류", "선택한 항목의 데이터가 유효하지 않습니다.")
            return
        
        # 보상 정보
        reward_type = values[0]
        reward_id = values[1]
        
        # 확인 대화상자
        if not show_message(self.top, "yesno", "확인", f"RewardID {reward_id}와 일치하는 모든 데이터를 숨기시겠습니까?"):
            return
        
        # 공통 유틸리티 함수 사용
        from utils.excel_utils import ExcelFileManager
        
        total_modified = ExcelFileManager.add_hash_to_reward_id(self.db_folder, self.cache, reward_id)
        
        if total_modified > 0:
            show_message(self.top, "info", "성공", f"RewardID {reward_id}와 일치하는 {total_modified}개 항목이 숨김 처리되었습니다.")
        else:
            show_message(self.top, "warning", "항목 없음", f"RewardID {reward_id}와 일치하는 항목을 찾을 수 없습니다.")

    def _remove_hash_from_all_matching_rewards(self):
        """선택한 보상 항목의 ID와 일치하는 모든 데이터의 A열에서 #을 제거합니다."""
        selected_item = self.detail_tree.focus()
        if not selected_item:
            show_message(self.top, "warning", "선택 오류", "보상 항목을 선택해주세요.")
            return
        
        values = self.detail_tree.item(selected_item, "values")
        if not values or len(values) < 2:
            show_message(self.top, "warning", "데이터 오류", "선택한 항목의 데이터가 유효하지 않습니다.")
            return
        
        # 보상 정보
        reward_type = values[0]
        reward_id = values[1]
        
        # 확인 대화상자
        if not show_message(self.top, "yesno", "확인", f"RewardID {reward_id}와 일치하는 모든 데이터의 숨김을 해제하시겠습니까?"):
            return
        
        # 공통 유틸리티 함수 사용
        from utils.excel_utils import ExcelFileManager
        
        total_modified = ExcelFileManager.remove_hash_from_reward_id(self.db_folder, self.cache, reward_id)
        
        if total_modified > 0:
            show_message(self.top, "info", "성공", f"RewardID {reward_id}와 일치하는 {total_modified}개 항목의 숨김이 해제되었습니다.")
        else:
            show_message(self.top, "warning", "항목 없음", f"RewardID {reward_id}와 일치하는 숨김 처리된 항목을 찾을 수 없습니다.")