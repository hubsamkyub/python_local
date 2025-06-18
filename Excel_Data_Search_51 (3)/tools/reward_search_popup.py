import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, PanedWindow
import pandas as pd
import sqlite3
import threading

from utils.cache_utils import load_cached_data, hash_paths, update_excel_cache
from utils.excel_utils import ExcelFileManager
from utils.config_utils import load_search_history, save_search_history
from utils.type_mappings import get_table_name_for_type, get_column_for_type, get_description_for_type
# 파일 상단에 추가
from db_relationships import get_deletion_impact, search_relationships

class RewardSearchPopup:
    def __init__(self, master, folder, db_folder, excel_cache=None):
        self.top = tk.Toplevel(master)
        self.top.title("🎁 Reward 검색기")
        self.top.geometry("1200x700")
        self.folder = folder
        self.db_folder = db_folder
        self.cache = excel_cache or update_excel_cache(folder, db_folder)
        self._reward_detached = []
        self.typecode_mapping = self.load_typecode_mapping()
        self.table_relationships = self.load_table_relationships()

        self.group_history = load_search_history("reward_group")
        self.reward_history = load_search_history("reward_id")

        # 상단 버튼 프레임 추가
        self._create_top_buttons()

        self.notebook = ttk.Notebook(self.top)
        self.notebook.pack(fill="both", expand=True)

        self._create_tab("RewardGroupID", "RewardGroupID", is_group=True)
        self._create_tab("RewardID", "RewardID", is_group=False)

    def _create_top_buttons(self):
        """상단 버튼들을 생성합니다."""
        button_frame = tk.Frame(self.top)
        button_frame.pack(fill="x", padx=10, pady=5)
        
        # 좌측 정렬용 프레임
        left_frame = tk.Frame(button_frame)
        left_frame.pack(side="left")
        
        # Reward목록 보기 버튼
        reward_list_btn = tk.Button(left_frame, text="📋 Reward목록 보기", 
                                   command=self.show_reward_group_list,
                                   font=("Helvetica", 10, "bold"),
                                   bg="#4CAF50", fg="white")
        reward_list_btn.pack(side="left", padx=5)

    def show_reward_group_list(self):
        """전체 RewardGroup 목록 팝업을 표시합니다."""
        RewardGroupListPopup(self.top, self.folder, self.db_folder, self.typecode_mapping, self.cache)

    def load_typecode_mapping(self):
        """typecode_mapping.json 파일을 로드합니다."""
        try:
            # 1. 폴더 내에서 파일 찾기
            mapping_path = os.path.join(self.folder, "typecode_mapping.json")
            
            # 2. 현재 디렉토리에서 파일 찾기
            if not os.path.exists(mapping_path):
                mapping_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "typecode_mapping.json")
            
            # 3. 상위 디렉토리에서 파일 찾기
            if not os.path.exists(mapping_path):
                mapping_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "typecode_mapping.json")
                
            # 4. documents 폴더에서 파일 찾기 (있을 경우)
            if not os.path.exists(mapping_path):
                documents_path = os.path.join(os.path.expanduser("~"), "Documents")
                mapping_path = os.path.join(documents_path, "typecode_mapping.json")
            
            # 파일이 존재하면 로드
            if os.path.exists(mapping_path):
                with open(mapping_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            
            # 파일을 찾지 못한 경우 파일 경로 출력 및 빈 리스트 반환
            print(f"타입코드 매핑 파일을 찾을 수 없습니다: {mapping_path}")
            return []
        except Exception as e:
            print(f"타입코드 매핑 로드 오류: {e}")
            return []
    
    def load_table_relationships(self):
        """table_relationships.json 파일을 로드합니다."""
        try:
            # 1. 폴더 내에서 파일 찾기
            rel_path = os.path.join(self.folder, "table_relationships.json")
            
            # 2. 현재 디렉토리에서 파일 찾기
            if not os.path.exists(rel_path):
                rel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "table_relationships.json")
            
            # 3. 상위 디렉토리에서 파일 찾기
            if not os.path.exists(rel_path):
                rel_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "table_relationships.json")
            
            # 4. documents 폴더에서 파일 찾기 (있을 경우)
            if not os.path.exists(rel_path):
                documents_path = os.path.join(os.path.expanduser("~"), "Documents")
                rel_path = os.path.join(documents_path, "table_relationships.json")
            
            # 파일이 존재하면 로드
            if os.path.exists(rel_path):
                with open(rel_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            
            # 파일을 찾지 못한 경우 파일 경로 출력 및 빈 객체 반환
            print(f"테이블 관계 파일을 찾을 수 없습니다: {rel_path}")
            return {}
        except Exception as e:
            print(f"테이블 관계 로드 오류: {e}")
            return {}

    # reward_search_popup.py 파일 수정

    def _create_tab(self, tab_name, column_name, is_group):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text=tab_name)

        # Entry + 버튼
        input_frame = tk.Frame(frame)
        input_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(input_frame, text=f"{column_name}:").pack(side="left")
        entry = tk.Entry(input_frame)
        entry.pack(side="left", padx=5)
        
        # 관계 정보 보기 버튼 추가
        if is_group:
            relation_btn = tk.Button(input_frame, text="관계 정보 보기", 
                                    command=lambda: self.show_relation_info(entry.get().strip(), is_group))
            relation_btn.pack(side="left", padx=5)
        
        search_btn = tk.Button(input_frame, text="검색", 
                            command=lambda: self.run_search(column_name, entry.get().strip(), result_tree, status_label, is_group))
        search_btn.pack(side="left", padx=5)

        # 히스토리 리스트박스
        history_listbox = tk.Listbox(frame, height=4)
        history_listbox.pack(fill="x", padx=10)
        history_listbox.bind("<<ListboxSelect>>", 
                        lambda e: self._on_select_history(e, entry, column_name, result_tree, status_label, is_group))

        delete_btn = tk.Button(frame, text=f"❌ {tab_name} 기록 삭제", 
                            command=lambda: self._delete_history(column_name, history_listbox))
        delete_btn.pack(padx=10, anchor="e")
        
        # 필터 프레임
        filter_frame = tk.Frame(frame)
        filter_frame.pack(fill="x", padx=10, pady=5)
        
        filter_label = tk.Label(filter_frame, text="결과 필터:")
        filter_label.pack(side="left")
        
        filter_var = tk.StringVar()
        filter_entry = tk.Entry(filter_frame, textvariable=filter_var, width=20)
        filter_entry.pack(side="left", padx=5)
        
        filter_btn = tk.Button(filter_frame, text="필터 적용", 
                            command=lambda: self._filter_tree(result_tree, filter_var))
        filter_btn.pack(side="left", padx=2)
        
        restore_btn = tk.Button(filter_frame, text="필터 초기화", 
                            command=lambda: self._restore_tree(result_tree))
        restore_btn.pack(side="left", padx=2)

        # 상세 정보 표시 프레임 추가
        detail_frame = tk.Frame(frame, relief=tk.RIDGE, bd=2)
        detail_frame.pack(fill="x", padx=10, pady=5)
        
        # 상세 정보 제목
        detail_title = tk.Label(detail_frame, text="선택한 행 상세 정보", font=("Helvetica", 10, "bold"))
        detail_title.pack(anchor="w", padx=5, pady=5)
        
        # 상세 정보 그리드를 위한 프레임
        grid_frame = tk.Frame(detail_frame)
        grid_frame.pack(fill="x", padx=5, pady=5)
        
        # TreeView 스크롤바
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 수직 스크롤바
        tree_scroll_y = tk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side="right", fill="y")
        
        # 수평 스크롤바
        tree_scroll_x = tk.Scrollbar(tree_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        # TreeView 구성 (간소화된 컬럼)
        columns = ["파일", "시트", "데이터"]
        result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                yscrollcommand=tree_scroll_y.set, 
                                xscrollcommand=tree_scroll_x.set)

        # 컬럼 설정
        result_tree.heading("파일", text="파일")
        result_tree.column("파일", anchor="w", width=150)
        result_tree.heading("시트", text="시트")
        result_tree.column("시트", anchor="w", width=100)
        result_tree.heading("데이터", text="데이터")
        result_tree.column("데이터", anchor="w", width=600)  # 데이터 컬럼을 넓게 설정

        result_tree.pack(fill="both", expand=True)
        
        # 스크롤바 연결
        tree_scroll_y.config(command=result_tree.yview)
        tree_scroll_x.config(command=result_tree.xview)

        # 스타일 설정
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)

        # 이벤트 바인딩
        result_tree.bind("<Double-1>", lambda e: self._open_excel(result_tree, column_name, is_group))
        result_tree.bind("<ButtonRelease-1>", lambda e: self._on_tree_select(e, result_tree, grid_frame, is_group))

        # 상태 표시
        status_label = tk.Label(frame, text="")
        status_label.pack(anchor="w", padx=10)

        # 저장
        if is_group:
            self.group_entry, self.group_result, self.group_status, self.group_history_listbox = entry, result_tree, status_label, history_listbox
            self.group_detail_frame = grid_frame
            self._refresh_history("RewardGroupID")
        else:
            self.reward_entry, self.reward_result, self.reward_status, self.reward_history_listbox = entry, result_tree, status_label, history_listbox
            self.reward_detail_frame = grid_frame
            self._refresh_history("RewardID")

    def _get_reward_group_columns(self):
        """RewardGroupTemplate 테이블의 컬럼을 가져옵니다."""
        try:
            db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
            if not os.path.exists(db_path):
                return ["파일", "시트", "GroupID", "RewardType", "RewardID", "RewardValue"]
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(RewardGroupTemplate)")
            columns = ["파일", "시트"] + [row[1] for row in cursor.fetchall()]
            conn.close()
            return columns
        except Exception as e:
            print(f"RewardGroup 컬럼 로드 오류: {e}")
            return ["파일", "시트", "GroupID", "RewardType", "RewardID", "RewardValue"]

    def _refresh_history(self, column_name):
        history = self.group_history if column_name == "RewardGroupID" else self.reward_history
        listbox = self.group_history_listbox if column_name == "RewardGroupID" else self.reward_history_listbox
        listbox.delete(0, tk.END)
        for item in history:
            listbox.insert(tk.END, item)

    def _update_history(self, column_name, keyword):
        history = self.group_history if column_name == "RewardGroupID" else self.reward_history
        if keyword and keyword not in history:
            history.insert(0, keyword)
            history[:] = history[:10]
            save_search_history(history, "reward_group" if column_name == "RewardGroupID" else "reward_id")
            self._refresh_history(column_name)

    def _delete_history(self, column_name, listbox):
        if column_name == "RewardGroupID":
            self.group_history.clear()
            save_search_history(self.group_history, "reward_group")
        else:
            self.reward_history.clear()
            save_search_history(self.reward_history, "reward_id")
        listbox.delete(0, tk.END)

    def _on_select_history(self, event, entry_widget, column_name, result_tree, status_label, is_group):
        idx = event.widget.curselection()
        if idx:
            keyword = event.widget.get(idx[0])
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, keyword)

    def run_search(self, column_name, keyword, tree, status_widget, is_group):
        threading.Thread(
            target=lambda: self._run_search_logic(column_name, keyword, tree, status_widget, is_group),
            daemon=True
        ).start()

    def _on_tree_select(self, event, tree, detail_frame, is_group):
        """트리 항목 선택 시 이벤트 처리"""
        item = tree.focus()
        if not item:
            return
        
        # 선택된 항목의 정보를 상태바에 표시
        values = tree.item(item, "values")
        if len(values) < 3:
            return
            
        file_path, sheet_name = values[0], values[1]
        
        # 이전 상세 정보 제거
        for widget in detail_frame.winfo_children():
            widget.destroy()
        
        try:
            # 전체 행 데이터 가져오기
            excel_path = os.path.join(self.folder, file_path)
            
            # 엑셀 파일 열기
            if os.path.exists(excel_path):
                sheet_info = self.cache.get(file_path, {}).get("sheets", {}).get(sheet_name, {})
                header_row = sheet_info.get("header_row", 0) if sheet_info else 0
                
                df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
                
                # 트리뷰 데이터와 일치하는 행 찾기
                search_column = "RewardGroupID" if is_group else "RewardID"
                search_value = None
                
                # 데이터 컬럼에서 ID 추출 시도
                data_text = values[2]
                if is_group:
                    # GroupID를 추출 (패턴: 'GroupID=숫자')
                    import re
                    match = re.search(r'GroupID=(\d+)', data_text)
                    if match:
                        search_value = int(match.group(1))
                else:
                    # RewardID를 추출 (패턴: 'ID=숫자')
                    import re
                    match = re.search(r'ID=(\d+)', data_text)
                    if match:
                        search_value = int(match.group(1))
                
                if search_value is not None:
                    df[search_column] = pd.to_numeric(df[search_column], errors='coerce').fillna(0).astype(int)
                    row = df[df[search_column] == search_value]
                    
                    if not row.empty:
                        # 첫 번째 일치하는 행 가져오기
                        row = row.iloc[0]
                        
                        # 그리드 형태로 데이터 표시 (이미지 2와 유사하게)
                        for i, col in enumerate(df.columns):
                            # 컬럼 이름
                            col_label = tk.Label(detail_frame, text=col, font=("Helvetica", 9, "bold"), 
                                            width=15, anchor="w", bg="#f0f0f0")
                            col_label.grid(row=i//2, column=(i%2)*2, sticky="w", padx=2, pady=2)
                            
                            # 값
                            val = row[col]
                            if isinstance(val, float) and val.is_integer():
                                val = int(val)
                            
                            val_label = tk.Label(detail_frame, text=str(val), width=25, anchor="w")
                            val_label.grid(row=i//2, column=(i%2)*2+1, sticky="w", padx=2, pady=2)
                            
                        # 상태 메시지 업데이트
                        status_widget = self.group_status if is_group else self.reward_status
                        status_widget.config(text=f"✅ 행 상세 정보 로드 완료")
                        return
                
                # 일치하는 행을 찾지 못한 경우
                not_found_label = tk.Label(detail_frame, text="일치하는 행 정보를 찾을 수 없습니다")
                not_found_label.pack(pady=10)
                
            else:
                error_label = tk.Label(detail_frame, text=f"파일을 찾을 수 없습니다: {file_path}")
                error_label.pack(pady=10)
                
        except Exception as e:
            # 에러 메시지 표시
            error_msg = str(e)
            error_label = tk.Label(detail_frame, text=f"오류 발생: {error_msg}", fg="red")
            error_label.pack(pady=10)
            
            # 상태 메시지 업데이트
            status_widget = self.group_status if is_group else self.reward_status
            status_widget.config(text=f"❌ 행 정보 로드 실패")

    def _filter_tree(self, tree, keyword_var):
        keyword = keyword_var.get().strip().lower()
        self._reward_detached.clear()

        for item in tree.get_children():
            values = tree.item(item, "values")
            if not any(keyword in str(v).lower() for v in values):
                self._reward_detached.append(item)
                tree.detach(item)

    def _restore_tree(self, tree):
        for item in self._reward_detached:
            tree.reattach(item, '', 'end')
        self._reward_detached.clear()

    def _run_search_logic(self, column_name, keyword, tree, status_widget, is_group):
        def update_status(msg):
            self.top.after(0, lambda: status_widget.config(text=msg))

        def clear_tree():
            self.top.after(0, lambda: tree.delete(*tree.get_children()))

        if not keyword.isdigit() or int(keyword) <= 0:
            self.top.after(0, lambda: messagebox.showwarning("입력 오류", "0보다 큰 정수를 입력하세요."))
            return

        self._update_history(column_name, keyword)
        target = int(keyword)
        flag = "has_reward_group_id" if is_group else "has_reward_id"
        col = "RewardGroupID" if is_group else "RewardID"
        count, total = 0, 0
        results = []

        update_status("🔍 검색 중...")
        clear_tree()

        for file, info in self.cache.items():
            for sheet, meta in info.get("sheets", {}).items():
                if sheet.startswith("#") or not meta.get(flag):
                    continue
                path, header = info["path"], meta["header_row"]
                try:
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    if col not in df.columns:
                        continue
                        
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
                    matched = df[df[col] == target]
                    
                    if not matched.empty:
                        for idx, row in matched.iterrows():
                            if is_group:
                                # GroupID 검색 결과 요약
                                data_summary = f"GroupID={row['RewardGroupID']}, Type={row.get('RewardType', '')}"
                                if 'RewardID' in row:
                                    data_summary += f", ID={row['RewardID']}"
                                if 'RewardValue' in row:
                                    data_summary += f", Value={row['RewardValue']}"
                            else:
                                # RewardID 검색 결과 요약
                                data_summary = f"Type={row.get('RewardType', '')}, ID={row[col]}"
                                if 'RewardValue' in row:
                                    data_summary += f", Value={row['RewardValue']}"
                                if 'Extra' in row:
                                    data_summary += f", Extra={row['Extra']}"
                            
                            # 파일, 시트, 데이터 요약 형식으로 결과 저장
                            results.append((os.path.basename(path), sheet, data_summary))
                            count += 1
                    total += 1
                except Exception as e:
                    print(f"[검색 오류] {path} / {sheet}: {e}")

        # 결과 정렬 및 표시
        for result in sorted(results):
            self.top.after(0, lambda r=result: tree.insert("", "end", values=r))

        update_status(f"✅ {col} 검색 완료 - {count}건 (대상 시트: {total})")
        

    def _open_excel(self, tree, column, is_group):
        item = tree.focus()
        if not item:
            return
        values = tree.item(item, "values")
        try:
            file, sheet = values[0], values[1]
            path = os.path.join(self.folder, file)
            
            # ExcelFileManager 심플 버전 사용
            success = ExcelFileManager.open_excel_file_simple(path, sheet)
            
            # 상태 메시지 업데이트
            status_widget = self.group_status if is_group else self.reward_status
            if success:
                status_widget.config(text=f"✅ {file} 파일을 열었습니다.")
            else:
                status_widget.config(text=f"❌ {file} 파일 열기 실패")
                
        except Exception as e:
            print(f"[엑셀 열기 오류] {file} / {sheet}: {e}")


    def show_relation_info(self, reward_id, is_group=True):
        """관계 정보 보기 팝업을 표시합니다."""
        if not reward_id.isdigit() or int(reward_id) <= 0:
            messagebox.showwarning("입력 오류", "유효한 ID를 입력하세요.")
            return

        if is_group:
            RewardGroupDetailPopup(
                self.top,
                int(reward_id),
                self.db_folder,
                self.typecode_mapping
            )
        else:
            RelationInfoPopup(
                self.top, 
                reward_id, 
                self.folder, 
                self.db_folder,
                self.typecode_mapping,
                is_group
            )


class RewardGroupListPopup:
    """전체 RewardGroup 목록을 표시하는 팝업 클래스"""
    def __init__(self, master, folder, db_folder, typecode_mapping, excel_cache):
        self.folder = folder
        self.db_folder = db_folder
        self.typecode_mapping = typecode_mapping
        self.cache = excel_cache
        self.top = Toplevel(master)
        self.top.title("📋 전체 RewardGroup 목록")
        self.top.geometry("1400x900")
        self._is_closing = False  # 창 닫힘 상태 추적
        
        # RewardGroupType 매핑
        self.reward_group_types = {
            1: "스테이지 - 클리어 최초 보상",
            2: "스테이지 - 클리어 보상",
            3: "스테이지 - 미션 클리어",
            10: "분해 보상",
            30: "퀘스트 보상",
            40: "결투장 최초 티어 달성 보상",
            41: "결투장 서브 시즌 정산 보상",
            42: "결투장 메인 시즌 정산 보상",
            43: "결투장 입장 기본 보상",
            44: "결투장 클리어 승리 보상",
            45: "결투장 스페셜 보상",
            50: "보스 스코어 최초 티어 달성 보상",
            51: "보스 스코어 시즌 정산 보상 (점수)",
            52: "보스 스코어 시즌 정산 보상 (랭킹)",
            53: "보스 스코어 시즌 정산 보상 (티어)",
            60: "계정 레벨 보상",
            61: "길드 레벨 출석 보상",
            70: "EventDirection 보상",
            80: "월정액 보상",
            90: "길드 레이드 티어 달성 보상",
            91: "길드 레이드 시즌 정산 보상 (점수)",
            92: "길드 레이드 시즌 정산 보상 (랭킹)",
            100: "포인트 리워드 보상",
            110: "거울의 환영 보상",
            120: "이벤트 버프",
            130: "아이템 판매",
            140: "달성 보상 (누적 포인트)",
            141: "시즌 보상 (층)",
            142: "정복자",
            150: "결제 누적 금액 보상"
        }
        
        # 엑셀 데이터 인메모리 캐시
        self.excel_data_cache = {}  # {(file_name, sheet_name): DataFrame}
        self.relationship_cache = {}
        self.relationship_cache_building = set()

        # 창 닫힘 이벤트 바인딩
        self.top.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self._build_ui()
        self._load_reward_groups()
        self._preload_rgi_data()
        
    def _on_close(self):
        """창 닫힘 이벤트 처리"""
        self._is_closing = True
        self.top.destroy()

#1
    def _safe_after(self, delay, func):
        """안전한 after 호출 (창이 닫혀있거나 속성이 없으면 무시)"""
        if not self._is_closing:
            try:
                self.top.after(delay, func)
            except tk.TclError:
                # 창이 이미 삭제된 경우 무시
                pass
            except AttributeError as attr_error:
                # 속성이 없는 경우 로그만 출력하고 무시
                print(f"[UI 업데이트 건너뛰기] 속성 없음: {attr_error}")
            except Exception as e:
                # 기타 오류는 로그만 출력
                print(f"[UI 업데이트 오류] {e}")


    def _preload_rgi_data(self):
        """RewardGroupID 관련 엑셀 데이터를 미리 메모리에 로드"""
        print("[초기화] RewardGroupID 관련 엑셀 데이터 미리 로딩 시작...")
        
        # 백그라운드에서 실행
        threading.Thread(target=self._preload_rgi_data_thread, daemon=True).start()


    def _search_actual_usage_files_ultrafast(self, reward_group_id):
        """딕셔너리 인덱스를 활용한 초고속 검색 (타입 안전성 강화)"""
        try:
            # 숫자형으로 변환
            if isinstance(reward_group_id, str):
                numeric_reward_group_id = int(reward_group_id)
            else:
                numeric_reward_group_id = int(reward_group_id)
                
            print(f"[초고속 검색] RewardGroupID {reward_group_id} -> {numeric_reward_group_id}")
            print(f"[초고속 검색] 인덱스 상태: {len(self.rgi_to_files) if hasattr(self, 'rgi_to_files') else 0}개")
            
            # 인덱스가 없으면 빈 결과 반환
            if not hasattr(self, 'rgi_to_files') or not self.rgi_to_files:
                print(f"[초고속 검색] ❌ 인덱스 없음")
                return []
            
            # 딕셔너리에서 즉시 조회 (O(1) 시간복잡도)
            if numeric_reward_group_id in self.rgi_to_files:
                result = self.rgi_to_files[numeric_reward_group_id]
                print(f"[초고속 검색] ✅ 즉시 조회 완료: {len(result)}개 사용처 발견")
                
                # 결과를 복사해서 반환 (참조 문제 방지)
                return [item.copy() for item in result]
            else:
                print(f"[초고속 검색] ❌ 해당 RewardGroupID 사용처 없음: {numeric_reward_group_id}")
                return []
            
        except Exception as e:
            print(f"[초고속 검색] 오류: {e}")
            return []


    def _search_actual_usage_files_parallel(self, reward_group_id):
        """병렬 처리를 통한 더 빠른 검색"""
        try:
            import concurrent.futures
            
            print(f"[관계 정보] {reward_group_id} 병렬 검색 시작")
            
            actual_usage_files = []
            numeric_reward_group_id = int(reward_group_id)
            
            def search_single_cache(cache_item):
                """단일 캐시 항목 검색"""
                cache_key, cache_data = cache_item
                try:
                    df = cache_data['df']
                    rgi_col = cache_data['rgi_col']
                    file_info = cache_data['file_info']
                    
                    # 해당 RewardGroupID 값 검색
                    matched = df[df[rgi_col] == numeric_reward_group_id]
                    count = len(matched)
                    
                    if count > 0:
                        # 테이블명 추출
                        sheet_name = file_info['sheet_name']
                        table_name = sheet_name
                        if '@' in sheet_name:
                            table_name = sheet_name.split('@')[0]
                        
                        # PK 정보 조회
                        pk_info = self._get_pk_info_simple(file_info['sheet_meta'])
                        
                        return {
                            'table_name': table_name,
                            'file_name': file_info['file_name'],
                            'sheet_name': file_info['sheet_name'],
                            'pk_info': f"{pk_info} ({count}개)"
                        }
                    
                    return None
                    
                except Exception as e:
                    print(f"[병렬 검색] 오류: {cache_key}: {e}")
                    return None
            
            # 병렬 실행 (최대 4개 스레드)
            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                # 모든 캐시 항목에 대해 병렬 검색 실행
                cache_items = list(self.excel_data_cache.items())
                future_to_cache = {executor.submit(search_single_cache, item): item for item in cache_items}
                
                for future in concurrent.futures.as_completed(future_to_cache):
                    result = future.result()
                    if result:
                        actual_usage_files.append(result)
            
            print(f"[관계 정보] 병렬 검색 완료: {len(actual_usage_files)}개 사용처 발견")
            return actual_usage_files
            
        except Exception as e:
            print(f"[관계 정보] 병렬 검색 오류: {e}")
            # 병렬 처리 실패 시 기본 방식으로 폴백
            return self._search_actual_usage_files_optimized(reward_group_id)
    
    
    def _search_actual_usage_files_optimized(self, reward_group_id):
        """최적화된 실제 사용처 검색 (인메모리 캐시 사용)"""
        try:
            print(f"[관계 정보] {reward_group_id} 실제 사용처 검색 - 인메모리 캐시 사용")
            
            actual_usage_files = []
            numeric_reward_group_id = int(reward_group_id)
            
            # 인메모리 캐시에서 검색 (매우 빠름)
            for cache_key, cache_data in self.excel_data_cache.items():
                try:
                    df = cache_data['df']
                    rgi_col = cache_data['rgi_col']
                    file_info = cache_data['file_info']
                    
                    # 해당 RewardGroupID 값 검색 (이미 전처리된 데이터이므로 매우 빠름)
                    matched = df[df[rgi_col] == numeric_reward_group_id]
                    count = len(matched)
                    
                    if count > 0:
                        print(f"[관계 정보] ✅ 캐시에서 발견: {file_info['file_name']}/{file_info['sheet_name']} ({count}개)")
                        
                        # 테이블명 추출
                        sheet_name = file_info['sheet_name']
                        table_name = sheet_name
                        if '@' in sheet_name:
                            table_name = sheet_name.split('@')[0]
                        
                        # PK 정보 조회
                        pk_info = self._get_pk_info_simple(file_info['sheet_meta'])
                        
                        actual_usage_files.append({
                            'table_name': table_name,
                            'file_name': file_info['file_name'],
                            'sheet_name': file_info['sheet_name'],
                            'pk_info': f"{pk_info} ({count}개)"
                        })
                
                except Exception as cache_error:
                    print(f"[관계 정보] 캐시 검색 오류: {cache_key}: {cache_error}")
            
            print(f"[관계 정보] 캐시 검색 완료: {len(actual_usage_files)}개 사용처 발견")
            return actual_usage_files
            
        except Exception as e:
            print(f"[관계 정보] 캐시 검색 전체 오류: {e}")
            return []

    def _build_relationship_cache(self, reward_group_id):
        """최적화된 관계 정보 구축"""
        try:
            print(f"[관계 정보] {reward_group_id} 관계 정보 구축 시작")
            
            # 1순위: 딕셔너리 인덱스 사용 (가장 빠름)
            if hasattr(self, 'rgi_to_files') and self.rgi_to_files:
                relationship_data = self._search_actual_usage_files_ultrafast(reward_group_id)
            # 2순위: 인메모리 캐시 사용
            elif self.excel_data_cache:
                relationship_data = self._search_actual_usage_files_optimized(reward_group_id)
            # 3순위: 기본 방식 (폴백)
            else:
                relationship_data = self._search_actual_usage_files(reward_group_id)
            
            # 캐시에 저장
            self.relationship_cache[reward_group_id] = relationship_data
            self.relationship_cache_building.discard(reward_group_id)
            
            print(f"[관계 정보] ✅ {reward_group_id} 캐시 구축 완료: {len(relationship_data)}개")
            
            # UI 업데이트
            self._safe_after(0, lambda: self._display_cached_relationship_info(relationship_data))
            
        except Exception as e:
            print(f"[관계 정보] 캐시 구축 오류: {e}")
            self.relationship_cache_building.discard(reward_group_id)

    def _display_cached_relationship_info(self, relationship_data):
        """캐시된 관계 정보를 화면에 표시"""
        try:
            # 트리뷰 초기화
            self.relationship_tree.delete(*self.relationship_tree.get_children())
            
            if not relationship_data:
                self.relationship_tree.insert("", "end", values=("사용처 없음", "-", "-", "-"))
                return
            
            # 캐시된 데이터 표시
            for data in relationship_data:
                values = (
                    data['table_name'],
                    data['file_name'], 
                    data['sheet_name'],
                    data['pk_info']
                )
                self.relationship_tree.insert("", "end", values=values)
            
            print(f"[관계 정보] 캐시된 데이터 표시 완료: {len(relationship_data)}개")
            
        except Exception as e:
            print(f"[관계 정보] 캐시된 데이터 표시 오류: {e}")


    def _build_ui(self):
        """UI를 구성합니다."""
        # 상단 프레임
        top_frame = tk.Frame(self.top)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(top_frame, text="전체 RewardGroup 목록", font=("Helvetica", 12, "bold")).pack(side="left")

        # 검색 프레임 추가
        search_frame = tk.Frame(self.top)
        search_frame.pack(fill="x", padx=10, pady=5)
        
        # RewardID 검색
        tk.Label(search_frame, text="RewardID 검색:").pack(side="left")
        self.reward_id_search_var = tk.StringVar()
        reward_id_entry = tk.Entry(search_frame, textvariable=self.reward_id_search_var, width=15)
        reward_id_entry.pack(side="left", padx=5)
        
        search_btn = tk.Button(search_frame, text="검색", command=self._search_by_reward_id)
        search_btn.pack(side="left", padx=5)
        
        clear_search_btn = tk.Button(search_frame, text="검색 초기화", command=self._clear_search)
        clear_search_btn.pack(side="left", padx=5)

        # 필터 프레임
        filter_frame = tk.Frame(top_frame)
        filter_frame.pack(side="right")
        
        # RewardGroupType 선택
        tk.Label(filter_frame, text="RewardGroupType:").pack(side="left")
        self.group_type_var = tk.StringVar(value="40")  # 기본값 40으로 변경 (이미지에 맞춤)
        group_type_combo = ttk.Combobox(filter_frame, textvariable=self.group_type_var,
                                       values=list(self.reward_group_types.keys()),
                                       state="readonly", width=15)
        group_type_combo.pack(side="left", padx=5)
        group_type_combo.bind("<<ComboboxSelected>>", self._on_group_type_change)
        
        test_btn = tk.Button(filter_frame, text="관계 테스트", command=self._test_db_relationships)
        test_btn.pack(side="left", padx=5)

        # 새로고침 버튼
        refresh_btn = tk.Button(filter_frame, text="새로고침", command=self._load_reward_groups)
        refresh_btn.pack(side="left", padx=5)
        
        # PanedWindow로 좌/우 분할
        self.paned = PanedWindow(self.top, orient=tk.HORIZONTAL)
        self.paned.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 좌측: RewardGroup 목록 (크기 줄임)
        left_frame = tk.Frame(self.paned)
        self.paned.add(left_frame, width=450)  # 600에서 450으로 줄임
        
        tk.Label(left_frame, text="RewardGroup 목록", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
        # 목록 프레임 (스크롤바 포함)
        list_frame = tk.Frame(left_frame)
        list_frame.pack(fill="both", expand=True)
        
        # 수직 스크롤바
        list_scroll_y = tk.Scrollbar(list_frame)
        list_scroll_y.pack(side="right", fill="y")
        
        # 트리뷰로 구현 - RewardGroupType으로 수정
        columns = ["RewardGroupType", "RewardGroupID", "RewardPayType"]
        self.group_tree = ttk.Treeview(list_frame, columns=columns, show="headings", 
                                     yscrollcommand=list_scroll_y.set)
        
        # 컬럼 설정
        self.group_tree.heading("RewardGroupType", text="RewardGroupType")
        self.group_tree.column("RewardGroupType", width=120, anchor="center")
        self.group_tree.heading("RewardGroupID", text="RewardGroupID")
        self.group_tree.column("RewardGroupID", width=120, anchor="center")
        self.group_tree.heading("RewardPayType", text="RewardPayType")
        self.group_tree.column("RewardPayType", width=120, anchor="center")
        
        self.group_tree.pack(fill="both", expand=True)
        list_scroll_y.config(command=self.group_tree.yview)
        
        # 우측: 상세 정보 (크기 늘림)
        right_frame = tk.Frame(self.paned)
        self.paned.add(right_frame, width=950)  # 800에서 950으로 늘림
        
        tk.Label(right_frame, text="RewardGroup 상세 정보", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
         # 상세 정보와 관계 정보를 수직으로 분할
        detail_paned = PanedWindow(right_frame, orient=tk.VERTICAL)
        detail_paned.pack(fill="both", expand=True)
        
        # 상단: 상세 정보 프레임
        detail_frame = tk.Frame(detail_paned)
        detail_paned.add(detail_frame, height=400)
        
        # 수직 스크롤바
        detail_scroll_y = tk.Scrollbar(detail_frame)
        detail_scroll_y.pack(side="right", fill="y")
        
        # 트리뷰로 구현
        columns = ["RewardGroupID", "RewardPayType", "RewardPayGroup", "Probability", 
                "RewardType", "RewardID", "TypeName", "RewardName", "MinCount", "MaxCount"]
        self.detail_tree = ttk.Treeview(detail_frame, columns=columns, show="headings", 
                                    yscrollcommand=detail_scroll_y.set)
        
        # 컬럼 설정
        for col in columns:
            self.detail_tree.heading(col, text=col)
            if col in ["RewardGroupID", "RewardType", "RewardID", "MinCount", "MaxCount"]:
                self.detail_tree.column(col, width=80, anchor="center")
            elif col == "RewardName":
                self.detail_tree.column(col, width=200, anchor="w")
            else:
                self.detail_tree.column(col, width=120, anchor="center")
        
        self.detail_tree.pack(fill="both", expand=True)
        detail_scroll_y.config(command=self.detail_tree.yview)
        
        # 하단: 관계 정보 프레임
        relationship_frame = tk.Frame(detail_paned)
        detail_paned.add(relationship_frame, height=300)
        
        tk.Label(relationship_frame, text="관계 정보 (RewardGroupID 사용처)", 
                font=("Helvetica", 10, "bold")).pack(anchor="w", pady=5)
        
        # 관계 정보 트리뷰
        rel_frame = tk.Frame(relationship_frame)
        rel_frame.pack(fill="both", expand=True)
        
        rel_scroll_y = tk.Scrollbar(rel_frame)
        rel_scroll_y.pack(side="right", fill="y")
        
        # 관계타입 컬럼 제거
        rel_columns = ["테이블명", "파일명", "시트명", "PK정보"]
        self.relationship_tree = ttk.Treeview(rel_frame, columns=rel_columns, show="headings", 
                                            yscrollcommand=rel_scroll_y.set)
        
        for col in rel_columns:
            self.relationship_tree.heading(col, text=col)
            if col == "파일명":
                self.relationship_tree.column(col, width=200, anchor="w")
            elif col == "PK정보":
                self.relationship_tree.column(col, width=150, anchor="w")
            else:
                self.relationship_tree.column(col, width=120, anchor="center")
        
        self.relationship_tree.pack(fill="both", expand=True)
        rel_scroll_y.config(command=self.relationship_tree.yview)
        
        # 더블클릭 이벤트 바인딩
        self.relationship_tree.bind("<Double-1>", self._open_excel_from_relationship)
        
        # 상태 표시
        self.status_label = tk.Label(self.top, text="")
        self.status_label.pack(anchor="w", padx=10, pady=5)
        
        # 이벤트 연결
        self.group_tree.bind("<<TreeviewSelect>>", self._on_group_select)

    def _search_by_reward_id(self):
        """RewardID로 검색하여 필터링합니다."""
        search_text = self.reward_id_search_var.get().strip()
        if not search_text:
            messagebox.showwarning("입력 오류", "검색할 RewardID를 입력하세요.")
            return
        
        self.status_label.config(text=f"🔍 RewardID '{search_text}' 검색 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=lambda: self._search_reward_id_thread(search_text), daemon=True).start()

    def _search_reward_id_thread(self, search_text):
        """백그라운드 스레드에서 RewardID 검색을 수행합니다."""
        try:
            print(f"[RewardID 검색] 전체 RewardGroupType에서 '{search_text}' 검색 시작")
            
            # 기존 목록 초기화
            self._safe_after(0, lambda: self.group_tree.delete(*self.group_tree.get_children()))
            
            # DB에서 RewardID가 포함된 데이터 검색 (전체 RewardGroupType 대상)
            db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
            if not os.path.exists(db_path):
                self._safe_after(0, lambda: self.status_label.config(
                    text="❌ RewardGroupTemplate.db 파일을 찾을 수 없습니다."))
                return
            
            conn = sqlite3.connect(db_path)
            
            # 전체 RewardGroupType에서 검색하도록 WHERE 조건 수정
            query = """
                SELECT DISTINCT RewardGroupType, RewardGroupID, RewardPayType
                FROM RewardGroupTemplate 
                WHERE CAST(RewardGroupID AS TEXT) LIKE ?
                ORDER BY RewardGroupType, RewardGroupID
            """
            
            # 부분일치를 위한 패턴
            search_pattern = f"%{search_text}%"
            
            print(f"[RewardID 검색] 실행 쿼리: {query}")
            print(f"[RewardID 검색] 검색 패턴: {search_pattern}")
            
            df = pd.read_sql_query(query, conn, params=(search_pattern,))
            conn.close()
            
            if df.empty:
                self._safe_after(0, lambda: self.status_label.config(
                    text=f"⚠️ 전체 RewardGroupType에서 RewardID '{search_text}'를 포함하는 데이터가 없습니다."))
                return
            
            print(f"[RewardID 검색] 검색 결과: {len(df)}개 발견")
            
            # RewardGroupType별로 결과 분류
            group_types_found = df['RewardGroupType'].unique()
            print(f"[RewardID 검색] 발견된 RewardGroupType: {group_types_found}")
            
            # 트리뷰에 추가
            for _, row in df.iterrows():
                values = (row["RewardGroupType"], row["RewardGroupID"], row["RewardPayType"])
                self._safe_after(0, lambda v=values: self.group_tree.insert("", "end", values=v))
            
            # 상태 메시지에 발견된 RewardGroupType 정보 포함
            group_type_names = []
            for gt in group_types_found:
                type_name = self.reward_group_types.get(gt, f"Type-{gt}")
                group_type_names.append(f"{gt}({type_name})")
            
            status_message = f"✅ RewardID '{search_text}' 검색 완료: {len(df)}개 발견 (타입: {', '.join(group_type_names)})"
            self._safe_after(0, lambda: self.status_label.config(text=status_message))
            
        except Exception as e:
            error_msg = f"❌ 검색 오류: {str(e)}"
            print(f"[RewardID 검색] 오류: {e}")
            self._safe_after(0, lambda: self.status_label.config(text=error_msg))

            
    def _clear_search(self):
        """검색을 초기화하고 전체 목록을 다시 로드합니다."""
        self.reward_id_search_var.set("")
        self._load_reward_groups()


    def _on_group_type_change(self, event=None):
        """RewardGroupType 변경 시 검색 상태 확인 후 동작 결정"""
        search_text = self.reward_id_search_var.get().strip()
        
        if search_text:
            # 검색어가 있으면 사용자에게 확인
            response = messagebox.askyesno(
                "검색 상태 확인", 
                f"현재 RewardID '{search_text}' 검색 결과가 표시되어 있습니다.\n"
                f"선택한 RewardGroupType만 표시하시겠습니까?\n\n"
                f"예: 선택 타입만 표시\n"
                f"아니오: 전체 검색 결과 유지"
            )
            
            if response:
                # 선택한 타입만 표시하도록 검색어 초기화 후 로드
                self.reward_id_search_var.set("")
                self._load_reward_groups()
            # else: 전체 검색 결과 유지 (아무것도 하지 않음)
        else:
            # 검색어가 없으면 기존대로 동작
            self._load_reward_groups()



    def _debug_cache_status(self):
        """캐시 상태를 확인하는 디버그 함수"""
        print(f"\n[캐시 상태] =====================================")
        print(f"[캐시 상태] relationship_cache: {len(self.relationship_cache)}개")
        print(f"[캐시 상태] relationship_cache_building: {self.relationship_cache_building}")
        print(f"[캐시 상태] rgi_to_files: {len(self.rgi_to_files) if hasattr(self, 'rgi_to_files') else 'None'}")
        print(f"[캐시 상태] excel_data_cache: {len(self.excel_data_cache)}개")
        
        if hasattr(self, 'rgi_to_files') and self.rgi_to_files:
            sample_keys = list(self.rgi_to_files.keys())[:5]
            print(f"[캐시 상태] rgi_to_files 샘플 키: {sample_keys}")
        
        if self.relationship_cache:
            cache_keys = list(self.relationship_cache.keys())[:5]
            print(f"[캐시 상태] relationship_cache 샘플 키: {cache_keys}")
        print(f"[캐시 상태] =====================================\n")


    def _load_relationship_info(self, reward_group_id):
        """캐시를 활용한 관계 정보 로드 (디버깅 강화)"""
        try:
            # 캐시 키를 숫자형으로 통일
            cache_key = int(reward_group_id)
            
            print(f"[관계 정보] RewardGroupID {reward_group_id} (캐시키: {cache_key}) 캐시 확인")
            print(f"[관계 정보] 현재 캐시 상태: {len(self.relationship_cache)}개 항목")
            print(f"[관계 정보] 구축 중: {self.relationship_cache_building}")
            
            # 1. 캐시에 있으면 바로 사용
            if cache_key in self.relationship_cache:
                print(f"[관계 정보] ✅ 캐시 HIT: {cache_key}")
                cached_data = self.relationship_cache[cache_key]
                self._display_cached_relationship_info(cached_data)
                return
            
            # 2. 현재 구축 중이면 스킵
            if cache_key in self.relationship_cache_building:
                print(f"[관계 정보] ⏳ 이미 구축 중: {cache_key}")
                return
            
            # 3. 초고속 인덱스가 준비되어 있으면 즉시 처리
            if hasattr(self, 'rgi_to_files') and self.rgi_to_files:
                print(f"[관계 정보] 🚀 초고속 인덱스 사용: {cache_key}")
                relationship_data = self._search_actual_usage_files_ultrafast(cache_key)
                
                # 즉시 캐시에 저장
                self.relationship_cache[cache_key] = relationship_data
                self._display_cached_relationship_info(relationship_data)
                return
            
            # 4. 인덱스가 없으면 백그라운드에서 검색
            print(f"[관계 정보] 🔍 백그라운드 검색 시작: {cache_key}")
            self.relationship_cache_building.add(cache_key)
            
            # 관계 정보 트리뷰 초기화
            self._safe_after(0, lambda: self.relationship_tree.delete(*self.relationship_tree.get_children()))
            self._safe_after(0, lambda: self.relationship_tree.insert("", "end", 
                        values=("검색 중...", "-", "-", "-")))
            
            # 백그라운드에서 검색
            threading.Thread(target=lambda: self._build_relationship_cache(cache_key), daemon=True).start()
            
        except Exception as e:
            print(f"[관계 정보] 캐시 처리 오류: {e}")
            if int(reward_group_id) in self.relationship_cache_building:
                self.relationship_cache_building.remove(int(reward_group_id))


    def _get_pk_info_simple(self, sheet_meta):
        """시트 메타데이터에서 간단한 PK 정보를 조회합니다."""
        try:
            primary_keys = sheet_meta.get("pk", [])
            
            if primary_keys:
                return f"PK: {', '.join(primary_keys)}"
            else:
                row_count = sheet_meta.get("rows", 0)
                return f"행 수: {row_count}"
                
        except Exception as e:
            return f"PK 오류: {str(e)[:20]}"
        

    def _test_db_relationships(self):
        """db_relationships 모듈 테스트"""
        print(f"\n[테스트] db_relationships 모듈 테스트 시작")
        
        try:
            from db_relationships import get_deletion_impact, search_relationships
            print(f"[테스트] ✅ 모듈 임포트 성공")
            
            # 테스트 호출
            impacts = get_deletion_impact("RewardGroupTemplate", "RewardGroupID")
            print(f"[테스트] get_deletion_impact 결과: {len(impacts)}개")
            
            if impacts:
                print(f"[테스트] 첫 번째 결과: {impacts[0]}")
            
            # 검색 테스트
            search_results = search_relationships("Reward")
            print(f"[테스트] search_relationships 결과: {len(search_results)}개")
            
        except Exception as e:
            print(f"[테스트] ❌ 오류: {e}")
            import traceback
            print(f"[테스트] 상세:\n{traceback.format_exc()}")



    def _get_detailed_cache_info(self, file_name, sheet_name):
        """캐시에서 상세한 시트 정보를 조회합니다."""
        try:
            if file_name not in self.cache:
                return {}
            
            file_data = self.cache[file_name]
            sheets = file_data.get("sheets", {})
            
            if sheet_name not in sheets:
                return {}
            
            sheet_meta = sheets[sheet_name]
            
            # 유용한 캐시 정보 정리
            cache_info = {
                "primary_keys": sheet_meta.get("primary_keys", []),
                "row_count": sheet_meta.get("row_count", 0),
                "header_row": sheet_meta.get("header_row", 0),
                "columns": sheet_meta.get("columns", []),
                "has_reward_group_id": sheet_meta.get("has_reward_group_id", False),
                "file_path": file_data.get("path", "")
            }
            
            return cache_info
            
        except Exception as e:
            print(f"[캐시 정보 조회 오류] {file_name}/{sheet_name}: {e}")
            return {}


    def _find_file_sheet_info(self, table_name):
        """테이블명으로 파일명과 시트명을 찾습니다."""
        file_info_list = []
        
        try:
            # 캐시에서 해당 테이블과 관련된 파일/시트 찾기
            for file_name, file_data in self.cache.items():
                for sheet_name, sheet_meta in file_data.get("sheets", {}).items():
                    # 1. 시트명과 테이블명 매칭 확인
                    sheet_match = table_name.lower() in sheet_name.lower()
                    file_match = table_name.lower() in file_name.lower()
                    
                    # 2. has_reward_group_id 플래그 확인 (RewardGroupID 관련 테이블인 경우)
                    has_reward_group_id = sheet_meta.get("has_reward_group_id", False)
                    
                    # 3. 컬럼에 RewardGroupID가 있는지 확인
                    columns = sheet_meta.get("columns", [])
                    has_rgi_column = any("rewardgroupid" in col.lower() for col in columns)
                    
                    # 매칭 조건: 이름 매칭 또는 RewardGroupID 관련
                    if sheet_match or file_match or has_reward_group_id or has_rgi_column:
                        # 캐시 정보도 함께 포함
                        file_info_list.append({
                            'file_name': file_name,
                            'sheet_name': sheet_name,
                            'file_path': file_data['path'],
                            'header_row': sheet_meta.get('header_row', 0),
                            'cache_info': {
                                "primary_keys": sheet_meta.get("pk", []),
                                "row_count": sheet_meta.get("rows", 0),
                                "header_row": sheet_meta.get("header_row", 0),
                                "columns": columns,
                                "has_reward_group_id": has_reward_group_id,
                                "file_path": file_data.get("path", "")
                            }
                        })
            
            # 매칭 정확도에 따른 정렬 (우선순위)
            def get_priority_score(info):
                score = 0
                sheet_name = info['sheet_name'].lower()
                file_name = info['file_name'].lower()
                table_lower = table_name.lower()
                
                # 시트명 정확 일치 (가장 높은 우선순위)
                if sheet_name == table_lower:
                    score += 100
                # 시트명 포함
                elif table_lower in sheet_name:
                    score += 80
                # 파일명 포함
                elif table_lower in file_name:
                    score += 60
                
                # RewardGroupID 관련 보너스
                if info['cache_info'].get('has_reward_group_id', False):
                    score += 20
                
                # 행 수 보너스 (데이터가 많을수록 우선)
                row_count = info['cache_info'].get('row_count', 0)
                if row_count > 0:
                    score += min(10, row_count / 100)  # 최대 10점 보너스
                
                return score
            
            file_info_list.sort(key=get_priority_score, reverse=True)
            
            # 디버깅용 로그
            if file_info_list:
                print(f"[파일 검색] {table_name}에 대해 {len(file_info_list)}개 파일/시트 발견")
                for i, info in enumerate(file_info_list[:3]):  # 상위 3개만 로그
                    score = get_priority_score(info)
                    print(f"  {i+1}. {info['file_name']}/{info['sheet_name']} (점수: {score})")
            
            return file_info_list
            
        except Exception as e:
            print(f"[파일/시트 정보 찾기 오류] {table_name}: {e}")
            return []
            
    def _table_references_reward_group_id(self, table_name, reward_group_id):
        """테이블이 실제로 해당 RewardGroupID 값을 참조하는지 확인합니다."""
        try:
            print(f"[참조 확인] {table_name} 테이블에서 RewardGroupID {reward_group_id} 실제 데이터 확인")
            
            # 1. 먼저 엑셀 파일에서 확인 (캐시 활용)
            found_in_excel = False
            excel_files_checked = 0
            
            for file_name, file_data in self.cache.items():
                for sheet_name, sheet_meta in file_data.get("sheets", {}).items():
                    # 테이블명과 관련된 파일/시트만 확인
                    if (table_name.lower() in sheet_name.lower() or 
                        table_name.lower() in file_name.lower()):
                        
                        excel_files_checked += 1
                        print(f"[참조 확인] 엑셀 검사: {file_name}/{sheet_name}")
                        
                        # RewardGroupID 컬럼이 있는지 확인
                        has_rgi = sheet_meta.get("has_reward_group_id", False)
                        if has_rgi:
                            print(f"[참조 확인] RewardGroupID 컬럼 존재 확인됨, 실제 데이터 검사 시작")
                            
                            # 실제 엑셀 파일에서 해당 값이 존재하는지 확인
                            try:
                                file_path = file_data['path']
                                header_row = sheet_meta.get('header_row', 0)
                                
                                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                                
                                # RewardGroupID 컬럼 찾기
                                rgi_col = None
                                for col in df.columns:
                                    if 'rewardgroupid' in str(col).lower():
                                        rgi_col = col
                                        break
                                
                                if rgi_col is not None:
                                    # 실제 데이터에서 해당 RewardGroupID 값 검색
                                    try:
                                        # 숫자형 검색
                                        df[rgi_col] = pd.to_numeric(df[rgi_col], errors='coerce')
                                        matched = df[df[rgi_col] == int(reward_group_id)]
                                        
                                        if not matched.empty:
                                            found_in_excel = True
                                            print(f"[참조 확인] ✅ {file_name}/{sheet_name}에서 RewardGroupID {reward_group_id} 실제 데이터 {len(matched)}개 발견")
                                            return True
                                    except:
                                        # 문자열 검색
                                        df[rgi_col] = df[rgi_col].astype(str)
                                        matched = df[df[rgi_col] == str(reward_group_id)]
                                        
                                        if not matched.empty:
                                            found_in_excel = True
                                            print(f"[참조 확인] ✅ {file_name}/{sheet_name}에서 RewardGroupID {reward_group_id} 실제 데이터 {len(matched)}개 발견")
                                            return True
                                
                            except Exception as excel_error:
                                print(f"[참조 확인] 엑셀 데이터 검사 오류: {file_name}/{sheet_name}: {excel_error}")
            
            print(f"[참조 확인] 엑셀 검사 완료: {excel_files_checked}개 파일/시트 확인, 실제 데이터 발견: {found_in_excel}")
            
            # 2. 엑셀에서 찾지 못했으면 DB에서 확인
            if not found_in_excel:
                print(f"[참조 확인] 엑셀에서 실제 데이터 없음, DB 확인 시작")
                db_path = os.path.join(self.db_folder, f"{table_name}.db")
                
                if not os.path.exists(db_path):
                    print(f"[참조 확인] ❌ DB 파일 없음: {db_path}")
                    return False
                
                print(f"[참조 확인] DB 파일 열기: {db_path}")
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # 테이블 구조 확인
                cursor.execute(f"PRAGMA table_info({table_name})")
                columns = [row[1] for row in cursor.fetchall()]
                print(f"[참조 확인] DB 테이블 컬럼: {columns}")
                
                # RewardGroupID 관련 컬럼 찾기
                reward_group_columns = [col for col in columns 
                                    if 'rewardgroup' in col.lower() and 'id' in col.lower()]
                print(f"[참조 확인] RewardGroupID 관련 컬럼: {reward_group_columns}")
                
                if not reward_group_columns:
                    print(f"[참조 확인] ❌ RewardGroupID 관련 컬럼 없음")
                    conn.close()
                    return False
                
                # 실제 데이터에서 해당 RewardGroupID가 사용되는지 확인
                for col in reward_group_columns:
                    print(f"[참조 확인] {col} 컬럼에서 RewardGroupID {reward_group_id} 실제 데이터 검색...")
                    cursor.execute(f"SELECT COUNT(*) FROM {table_name} WHERE {col} = ?", (reward_group_id,))
                    count = cursor.fetchone()[0]
                    print(f"[참조 확인] {col} 컬럼 검색 결과: {count}개")
                    
                    if count > 0:
                        print(f"[참조 확인] ✅ {table_name}.{col}에서 RewardGroupID {reward_group_id}를 실제로 {count}번 참조함")
                        conn.close()
                        return True
                
                print(f"[참조 확인] ❌ DB에서 RewardGroupID {reward_group_id} 실제 데이터 없음")
                conn.close()
                return False
            
            return found_in_excel
            
        except Exception as e:
            print(f"[참조 확인] ❌ 오류 발생: {table_name}: {e}")
            return False     

    def _get_pk_info(self, table_name, reward_group_id, file_info_list):
        """캐시에서 간단한 PK 정보를 조회합니다."""
        try:
            if not file_info_list:
                print(f"[PK 정보] {table_name}: 파일 정보 없음")
                return "파일 정보 없음"
            
            # 첫 번째 파일 정보 사용 (가장 관련성이 높은 것)
            file_info = file_info_list[0]
            file_name = file_info['file_name']
            sheet_name = file_info['sheet_name']
            
            print(f"[PK 정보] {table_name}: {file_name}/{sheet_name}에서 PK 정보 조회")
            
            # 캐시에서 해당 파일의 시트 정보 조회
            if file_name not in self.cache:
                print(f"[PK 정보] ❌ 캐시에 파일 정보 없음: {file_name}")
                return "캐시 정보 없음"
            
            file_data = self.cache[file_name]
            sheets = file_data.get("sheets", {})
            
            if sheet_name not in sheets:
                print(f"[PK 정보] ❌ 캐시에 시트 정보 없음: {sheet_name}")
                return "시트 정보 없음"
            
            sheet_meta = sheets[sheet_name]
            print(f"[PK 정보] 캐시 메타데이터 로드 성공")
            
            # PK 정보 조회 (캐시에서 직접 가져오기)
            primary_keys = sheet_meta.get("pk", [])
            
            print(f"[PK 정보] PK: {primary_keys}")
            
            # 간단한 PK 정보만 반환
            if primary_keys:
                pk_info = f"PK: {', '.join(primary_keys)}"
            else:
                # PK가 없으면 행 수만 표시
                row_count = sheet_meta.get("rows", 0)
                pk_info = f"행 수: {row_count}"
            
            print(f"[PK 정보] 최종 결과: {pk_info}")
            return pk_info
            
        except Exception as e:
            print(f"[PK 정보] ❌ 오류 발생: {table_name}: {e}")
            return f"오류: {str(e)[:20]}"


    def _open_excel_from_relationship(self, event):
        """관계 정보에서 더블클릭 시 엑셀 파일을 심플하게 엽니다."""
        item = self.relationship_tree.focus()
        if not item:
            return
        
        values = self.relationship_tree.item(item, "values")
        if not values or len(values) < 3:
            return
        
        table_name = values[0]
        file_name = values[1]
        sheet_name = values[2]
        
        # "사용처 없음"이나 오류인 경우 제외
        if table_name in ["사용처 없음", "오류 발생"] or file_name == "-":
            messagebox.showinfo("알림", "열 수 있는 파일이 아닙니다.")
            return
        
        try:
            # 파일 경로 찾기
            file_path = None
            for cached_file, file_data in self.cache.items():
                if cached_file == file_name:
                    file_path = file_data['path']
                    break
            
            if file_path and os.path.exists(file_path):
                # ExcelFileManager 심플 버전 사용
                success = ExcelFileManager.open_excel_file_simple(file_path, sheet_name)
                
                if success:
                    self.status_label.config(text=f"✅ {file_name} 파일을 열었습니다.")
                else:
                    self.status_label.config(text=f"❌ {file_name} 파일 열기 실패")
            else:
                messagebox.showerror("오류", f"파일을 찾을 수 없습니다: {file_name}")
                
        except Exception as e:
            messagebox.showerror("오류", f"파일 열기 실패: {str(e)}")
            print(f"[엑셀 열기 오류] {file_name}/{sheet_name}: {e}")
            

    def safe_update_status(self, message):
        """상태 메시지를 안전하게 업데이트합니다."""
        if hasattr(self, 'status_label') and self.status_label:
            try:
                self._safe_after(0, lambda: self.status_label.config(text=message))
            except:
                print(f"[상태 업데이트] {message}")
        else:
            print(f"[상태 업데이트] {message}")

    def _preload_rgi_data_thread(self):
        """백그라운드에서 엑셀 데이터 미리 로딩 (개선된 버전)"""
        try:
            # 총 작업량 계산
            rgi_files = [(file_name, sheet_name, file_data, sheet_meta) 
                        for file_name, file_data in self.cache.items()
                        for sheet_name, sheet_meta in file_data.get("sheets", {}).items()
                        if sheet_meta.get("has_reward_group_id", False)]
            
            total_files = len(rgi_files)
            
            self.safe_update_status(f"관계 정보 초고속 인덱스 구축 중... (0/{total_files})")
            
            loaded_count = 0
            self.rgi_to_files = {}
            
            print(f"[인덱스 구축] 시작: {total_files}개 파일/시트")
            
            for file_name, sheet_name, file_data, sheet_meta in rgi_files:
                try:
                    file_path = file_data['path']
                    header_row = sheet_meta.get('header_row', 0)
                    
                    # 엑셀 데이터 로드
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                    
                    # RewardGroupID 컬럼 찾기
                    rgi_col = None
                    for col in df.columns:
                        if 'rewardgroupid' in str(col).lower():
                            rgi_col = col
                            break
                    
                    if rgi_col:
                        # 숫자형으로 변환
                        df[rgi_col] = pd.to_numeric(df[rgi_col], errors='coerce').fillna(0).astype(int)
                        
                        # RewardGroupID별 인덱스 구축
                        unique_rgis = df[rgi_col].unique()
                        for rgi in unique_rgis:
                            if rgi > 0:  # 0은 제외
                                if rgi not in self.rgi_to_files:
                                    self.rgi_to_files[rgi] = []
                                
                                # 해당 RGI가 몇 개 있는지 미리 계산
                                count = len(df[df[rgi_col] == rgi])
                                
                                # 테이블명 추출
                                table_name = sheet_name
                                if '@' in sheet_name:
                                    table_name = sheet_name.split('@')[0]
                                
                                # PK 정보 조회
                                pk_info = self._get_pk_info_simple(sheet_meta)
                                
                                self.rgi_to_files[rgi].append({
                                    'table_name': table_name,
                                    'file_name': file_name,
                                    'sheet_name': sheet_name,
                                    'pk_info': f"{pk_info} ({count}개)"
                                })
                        
                        loaded_count += 1
                        
                        # 진행상황 업데이트 (5개마다)
                        if loaded_count % 5 == 0:
                            progress = f"관계 정보 초고속 인덱스 구축 중... ({loaded_count}/{total_files})"
                            self.safe_update_status(progress)
                            print(f"[인덱스 구축] 진행: {loaded_count}/{total_files}")
                
                except Exception as e:
                    print(f"[인덱스 구축] 파일 오류: {file_name}/{sheet_name}: {e}")
                    loaded_count += 1  # 오류가 있어도 카운트는 증가
            
            # 완료 메시지
            completion_msg = f"✅ 초고속 인덱스 구축 완료 ({loaded_count}개 파일, {len(self.rgi_to_files)}개 RewardGroupID)"
            self.safe_update_status(completion_msg)
            
            print(f"[인덱스 구축] 완료: {len(self.rgi_to_files)}개 RewardGroupID 인덱스")
            print(f"[인덱스 구축] 샘플 키 10개: {list(self.rgi_to_files.keys())[:10]}")
            
        except Exception as e:
            print(f"[인덱스 구축] 전체 오류: {e}")
            self.safe_update_status("❌ 인덱스 구축 실패")


    #심플 엑셀 열기 버전
    def _open_excel(self, tree, column, is_group):
        item = tree.focus()
        if not item:
            return
        values = tree.item(item, "values")
        try:
            file, sheet = values[0], values[1]
            path = os.path.join(self.folder, file)
            
            if os.path.exists(path):
                # 파일 핸들 해제 시도
                self._close_excel_handles(path)
                
                # 잠시 대기
                import time
                time.sleep(0.5)
                
                # 파일 열기
                import subprocess
                import sys
                
                if sys.platform.startswith('win'):
                    os.startfile(path)
                else:
                    subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', path])
                
                # 상태 메시지 업데이트
                status_widget = self.group_status if is_group else self.reward_status
                status_widget.config(text=f"✅ {file} 파일을 열었습니다.")
            else:
                print(f"[엑셀 열기 오류] 파일을 찾을 수 없음: {path}")
                
        except Exception as e:
            print(f"[엑셀 열기 오류] {file} / {sheet}: {e}")


    def _load_reward_groups(self):
        """RewardGroup 목록을 로드합니다."""
        selected_type = self.group_type_var.get()
        if not selected_type:
            selected_type = "40"  # 기본값
            
        type_name = self.reward_group_types.get(int(selected_type), f"Type-{selected_type}")
        self.status_label.config(text=f"🔍 RewardGroupType {selected_type} ({type_name}) 로딩 중...")
        
        # 백그라운드 스레드로 실행
        threading.Thread(target=lambda: self._load_reward_groups_thread(int(selected_type)), daemon=True).start()


    def _load_reward_groups_thread(self, group_type):
        """백그라운드 스레드에서 RewardGroup 목록을 로드합니다."""
        try:
            self._safe_after(0, lambda: self.group_tree.delete(*self.group_tree.get_children()))
            
            # DB에서 RewardGroupTemplate 데이터 로드
            db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
            if not os.path.exists(db_path):
                self._safe_after(0, lambda: self.status_label.config(
                    text="❌ RewardGroupTemplate.db 파일을 찾을 수 없습니다."))
                return
            
            conn = sqlite3.connect(db_path)
            
            # 선택한 RewardGroupType에 해당하는 데이터만 조회 - RewardGroupType을 표시
            query = """
                SELECT DISTINCT RewardGroupType, RewardGroupID, RewardPayType
                FROM RewardGroupTemplate 
                WHERE RewardGroupType = ?
                ORDER BY RewardGroupID
            """
            
            df = pd.read_sql_query(query, conn, params=(group_type,))
            conn.close()
            
            if df.empty:
                self._safe_after(0, lambda: self.status_label.config(
                    text=f"⚠️ RewardGroupType {group_type}에 해당하는 데이터가 없습니다."))
                return
            
            # 트리뷰에 추가
            for _, row in df.iterrows():
                values = (row["RewardGroupType"], row["RewardGroupID"], row["RewardPayType"])
                self._safe_after(0, lambda v=values: self.group_tree.insert("", "end", values=v))


            # 목록 로딩 완료 후 관계 정보 미리 구축 시작
            if len(df) > 0:
                reward_group_ids = df['RewardGroupID'].unique()
                print(f"[관계 정보] {len(reward_group_ids)}개 RewardGroupID 관계 정보 미리 구축 시작")
                
                # 백그라운드에서 모든 관계 정보 미리 구축
                threading.Thread(target=lambda: self._prebuild_relationship_cache(reward_group_ids), daemon=True).start()

            type_name = self.reward_group_types.get(group_type, f"Type-{group_type}")
            self._safe_after(0, lambda: self.status_label.config(
                text=f"✅ RewardGroupType {group_type} ({type_name}) 로딩 완료: {len(df)}개"))
            
        except Exception as e:
            error_msg = f"❌ 오류 발생: {str(e)}"
            self._safe_after(0, lambda: self.status_label.config(text=error_msg))

    
    def _prebuild_relationship_cache(self, reward_group_ids):
        """모든 RewardGroupID의 관계 정보를 미리 구축"""
        try:
            total = len(reward_group_ids)
            for i, reward_group_id in enumerate(reward_group_ids):
                if reward_group_id not in self.relationship_cache:
                    print(f"[사전 구축] {i+1}/{total} RewardGroupID {reward_group_id} 처리 중")
                    
                    try:
                        relationship_data = self._search_actual_usage_files(reward_group_id)
                        self.relationship_cache[reward_group_id] = relationship_data
                        
                        # 진행상황 업데이트 (너무 자주 하지 않도록)
                        if i % 10 == 0:
                            progress = (i + 1) / total * 100
                            self._safe_after(0, lambda p=progress: 
                                self.status_label.config(text=f"관계 정보 사전 구축 중... {p:.1f}%"))
                            
                    except Exception as e:
                        print(f"[사전 구축] RewardGroupID {reward_group_id} 오류: {e}")
            
            print(f"[사전 구축] 완료: {len(self.relationship_cache)}개 관계 정보 구축됨")
            self._safe_after(0, lambda: self.status_label.config(text="✅ 관계 정보 사전 구축 완료"))
            
        except Exception as e:
            print(f"[사전 구축] 오류: {e}")


    def _on_group_select(self, event):
        """RewardGroup 선택 시 상세 정보를 표시합니다."""
        selected_item = self.group_tree.focus()
        if not selected_item:
            return
        
        values = self.group_tree.item(selected_item, "values")
        if not values:
            return
        
        reward_group_id = values[1]  # RewardGroupID
        
        # 이전 검색 취소 (있다면)
        self._cancel_previous_searches()
        
        # 상세 정보 로드
        self.status_label.config(text=f"🔍 RewardGroupID {reward_group_id} 상세 정보 로딩 중...")
        
        # 상세 정보는 즉시 로드 (빠름)
        threading.Thread(target=lambda: self._load_group_detail(reward_group_id), daemon=True).start()
        
        # 관계 정보는 약간의 지연 후 로드 (사용자가 빠르게 선택을 바꾸면 스킵)
        self.current_reward_group_id = reward_group_id
        self.top.after(200, lambda: self._delayed_load_relationship_info(reward_group_id))

    def _delayed_load_relationship_info(self, reward_group_id):
        """지연된 관계 정보 로드 (사용자가 다른 항목 선택하면 취소)"""
        # 현재 선택된 ID와 다르면 취소
        if hasattr(self, 'current_reward_group_id') and self.current_reward_group_id != reward_group_id:
            print(f"[관계 정보] 선택 변경으로 인한 취소: {reward_group_id}")
            return
        
        # 관계 정보 로드
        self._load_relationship_info(reward_group_id)

    def _cancel_previous_searches(self):
        """이전 검색들을 취소합니다."""
        # 현재 구축 중인 캐시들 취소 표시
        if hasattr(self, 'relationship_cache_building'):
            for building_id in list(self.relationship_cache_building):
                print(f"[관계 정보] 구축 중단: {building_id}")
                

    def _load_group_detail(self, reward_group_id):
        """선택한 RewardGroup의 상세 정보를 로드합니다."""
        try:
            # 상세 정보 트리뷰 초기화
            self.top.after(0, lambda: self.detail_tree.delete(*self.detail_tree.get_children()))
            
            # DB에서 해당 RewardGroupID의 모든 데이터 조회
            db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
            conn = sqlite3.connect(db_path)
            
            query = """
                SELECT RewardGroupID, RewardPayType, RewardPayGroup, Probability, 
                       RewardType, RewardID, RewardMinCount, RewardMaxCount
                FROM RewardGroupTemplate 
                WHERE RewardGroupID = ?
                ORDER BY RewardPayGroup, RewardType, RewardID
            """
            
            df = pd.read_sql_query(query, conn, params=(reward_group_id,))
            conn.close()
            
            if df.empty:
                self.top.after(0, lambda: self.status_label.config(
                    text=f"⚠️ RewardGroupID {reward_group_id}에 해당하는 상세 정보가 없습니다."))
                return
            
            # 각 행에 대해 RewardType과 RewardID에 따른 이름 조회
            for _, row in df.iterrows():
                reward_type = row["RewardType"]
                reward_id = row["RewardID"]
                
                # RewardType에 따른 타입명과 보상명 조회
                type_name, reward_name = self._resolve_reward_info(reward_type, reward_id)
                
                values = (
                    row["RewardGroupID"],
                    row["RewardPayType"],
                    row["RewardPayGroup"],
                    row["Probability"],
                    reward_type,
                    reward_id,
                    type_name,
                    reward_name,
                    row["RewardMinCount"],
                    row["RewardMaxCount"]
                )
                
                self.top.after(0, lambda v=values: self.detail_tree.insert("", "end", values=v))
            
            self.top.after(0, lambda: self.status_label.config(
                text=f"✅ RewardGroupID {reward_group_id} 상세 정보 로드 완료 ({len(df)}개 항목)"))
            
        except Exception as e:
            error_msg = f"❌ 상세 정보 로드 오류: {str(e)}"
            self.top.after(0, lambda: self.status_label.config(text=error_msg))

    def _resolve_reward_info(self, reward_type, reward_id):
        """RewardType에 따라 RewardID의 타입명과 이름을 조회합니다."""
        try:
            # type_mappings.py를 사용하여 타입 정보 조회
            table_name = get_table_name_for_type(reward_type, "reward")
            column_name = get_column_for_type(reward_type, "reward")
            type_description = get_description_for_type(reward_type, "reward")
            
            print(f"[RewardInfo] Type={reward_type}, ID={reward_id}")
            print(f"[RewardInfo] 매핑 정보: 테이블={table_name}, 컬럼={column_name}, 설명={type_description}")
            
            # 테이블 정보가 없으면 기본값 반환
            if not table_name or not column_name:
                return (type_description, f"매핑 없음: {reward_id}")
            
            # DB 파일 경로 확인
            db_path = os.path.join(self.db_folder, f"{table_name}.db")
            print(f"[RewardInfo] DB 경로: {db_path}, 존재={os.path.exists(db_path)}")
            
            if not os.path.exists(db_path):
                return (type_description, f"{table_name} DB 없음")
            
            # RewardID를 숫자로 변환 시도
            try:
                if isinstance(reward_id, str) and reward_id.isdigit():
                    numeric_reward_id = int(reward_id)
                else:
                    numeric_reward_id = int(float(reward_id))
            except:
                print(f"[RewardInfo] RewardID 숫자 변환 실패: {reward_id}")
                numeric_reward_id = reward_id
            
            # DB에서 정보 조회
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 1. 테이블 구조 확인
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns_info = cursor.fetchall()
            column_names = [col[1] for col in columns_info]
            print(f"[RewardInfo] 테이블 컬럼: {column_names}")
            
            # 대상 컬럼이 테이블에 있는지 확인
            if column_name not in column_names:
                print(f"[RewardInfo] 경고: {column_name} 컬럼이 테이블에 없습니다. 대체 컬럼 검색 중...")
                # ID 관련 컬럼 자동 탐지
                for col_name in column_names:
                    if 'id' in col_name.lower() or 'template' in col_name.lower() or 'type' in col_name.lower():
                        column_name = col_name
                        print(f"[RewardInfo] 대체 컬럼 발견: {column_name}")
                        break
            
            # 2. 데이터 조회 (여러 방식 시도)
            rows = []
            try:
                # 정확한 숫자 일치로 검색
                query = f"SELECT * FROM {table_name} WHERE {column_name} = ?"
                print(f"[RewardInfo] 쿼리1: {query}, 값={numeric_reward_id}")
                cursor.execute(query, (numeric_reward_id,))
                rows = cursor.fetchall()
                
                if not rows:
                    # 문자열 일치로 검색
                    str_reward_id = str(reward_id)
                    print(f"[RewardInfo] 쿼리2: {query}, 값={str_reward_id}")
                    cursor.execute(query, (str_reward_id,))
                    rows = cursor.fetchall()
                    
            except Exception as qe:
                print(f"[RewardInfo] 쿼리 실행 오류: {qe}")
                # pandas로 시도
                try:
                    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
                    # 타입별로 다른 검색 방식 적용
                    if str(reward_type) == "30":  # GoodsMaxValue의 경우
                        # GoodsType 컬럼으로 검색
                        if 'GoodsType' in df.columns:
                            filtered = df[df['GoodsType'] == numeric_reward_id]
                        else:
                            filtered = df[df[column_name] == numeric_reward_id]
                    else:
                        # 일반적인 검색
                        df[column_name] = pd.to_numeric(df[column_name], errors='coerce')
                        filtered = df[df[column_name] == numeric_reward_id]
                    
                    if not filtered.empty:
                        rows = [tuple(row) for _, row in filtered.iterrows()]
                        print(f"[RewardInfo] DataFrame 필터링 성공: {len(filtered)} 행")
                except Exception as dfe:
                    print(f"[RewardInfo] DataFrame 처리 오류: {dfe}")
            
            # 3. 결과 처리
            if rows:
                print(f"[RewardInfo] 결과 행 수: {len(rows)}")
                # 첫 번째 행 사용
                row_data = rows[0]
                
                # 이름 컬럼 찾기 (우선순위: Name, DisplayName, Title, Description, Desc)
                name_cols = ["Name", "DisplayName", "Title", "Description", "Desc"]
                result_name = None
                
                for name_col in name_cols:
                    if name_col in column_names:
                        col_index = column_names.index(name_col)
                        if col_index < len(row_data) and row_data[col_index]:
                            result_name = str(row_data[col_index])
                            print(f"[RewardInfo] 이름 찾음: {result_name} (컬럼: {name_col})")
                            break
                
                if result_name:
                    conn.close()
                    return (type_description, result_name)
                else:
                    # 이름 컬럼이 없으면 첫 번째 컬럼 값 사용
                    first_value = row_data[0] if row_data else numeric_reward_id
                    conn.close()
                    return (type_description, f"{table_name}: {first_value}")
            else:
                print(f"[RewardInfo] 결과 없음: {table_name} 테이블에서 {reward_id} 찾을 수 없음")
                conn.close()
                return (type_description, f"{table_name}에 없음")
                
        except Exception as e:
            import traceback
            print(f"[RewardInfo] 전체 처리 오류: {e}")
            print(traceback.format_exc())
            return ("오류", f"조회 실패: {str(e)[:30]}")


class RewardGroupDetailPopup:
    def __init__(self, master, group_id, db_folder, typecode_mapping):
        self.group_id = group_id
        self.db_folder = db_folder
        self.typecode_mapping = typecode_mapping
        self.top = Toplevel(master)
        self.top.title(f"🎁 Reward Group 상세 - GroupID: {group_id}")
        self.top.geometry("900x600")
        self.tree = None

        self._build_ui()
        self._load_group_rewards()

    def _build_ui(self):
        tk.Label(self.top, text=f"GroupID {self.group_id} 에 속한 보상 목록",
                 font=("Helvetica", 12, "bold")).pack(pady=10)

        frame = tk.Frame(self.top)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        scroll_y = tk.Scrollbar(frame)
        scroll_y.pack(side="right", fill="y")

        columns = ["RewardGroupType",  "RewardTypeName", "RewardGroupID", "RewardType", "RewardID", "RewardName", "RewardMinCount", "RewardMaxCount"]
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", yscrollcommand=scroll_y.set)
        scroll_y.config(command=self.tree.yview)

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=150)
        self.tree.pack(fill="both", expand=True)

    def _load_group_rewards(self):
        db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
        if not os.path.exists(db_path):
            messagebox.showerror("DB 오류", f"DB 파일이 없습니다: {db_path}")
            return

        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query(
            f"""
            SELECT RewardGroupType, RewardGroupID, RewardType, RewardID, RewardMinCount, RewardMaxCount
            FROM RewardGroupTemplate
            WHERE RewardGroupID = {self.group_id}
            """,
            conn
        )
        conn.close()

        for _, row in df.iterrows():
            rtype = int(row["RewardType"])
            rid = int(row["RewardID"])
            type_name, reward_name = self._resolve_reward_info(rtype, rid)

            values = (
                row["RewardGroupType"], type_name,
                row["RewardGroupID"], rtype, rid, reward_name,
                row["RewardMinCount"], row["RewardMaxCount"]
            )
            self.tree.insert("", "end", values=values)

    def _resolve_reward_info(self, reward_type, reward_id):
        """RewardType에 따라 이름을 매핑"""
        try:
            # type_mappings.py를 사용하여 타입 정보 조회
            table_name = get_table_name_for_type(reward_type, "reward")
            column_name = get_column_for_type(reward_type, "reward")
            type_description = get_description_for_type(reward_type, "reward")
            
            print(f"[RewardGroupDetail] Type={reward_type}, ID={reward_id}")
            print(f"[RewardGroupDetail] 매핑 정보: 테이블={table_name}, 컬럼={column_name}, 설명={type_description}")
            
            # 테이블 정보가 없으면 기본값 반환
            if not table_name or not column_name:
                return (type_description, f"매핑 없음: {reward_id}")
            
            # DB 파일 경로 확인
            db_path = os.path.join(self.db_folder, f"{table_name}.db")
            if not os.path.exists(db_path):
                return (type_description, f"{table_name} DB 없음")

            try:
                conn = sqlite3.connect(db_path)
                df = pd.read_sql_query(f"SELECT * FROM {table_name} WHERE {column_name} = ?", conn, params=(reward_id,))
                conn.close()
                
                if not df.empty:
                    # 이름 컬럼 찾기: Name, DisplayName, Title, Description 등 우선순위 적용
                    name_cols = ["Name", "DisplayName", "Title", "Description", "Desc"]
                    for col in name_cols:
                        if col in df.columns and pd.notna(df.iloc[0][col]):
                            return (type_description, str(df.iloc[0][col]))
                    return (type_description, f"{table_name} 매칭")
                
                return (type_description, "데이터 없음")
                
            except Exception as e:
                print(f"[RewardGroupDetail] DB 조회 실패 {table_name}: {e}")
                return (type_description, "조회 실패")
                
        except Exception as e:
            print(f"[RewardGroupDetail] 전체 오류: {e}")
            return ("오류", "처리 실패")


class RelationInfoPopup:
    """관계 정보를 보여주는 팝업 클래스"""
    def __init__(self, master, reward_id, folder, db_folder, typecode_mapping, is_group=True):
        self.top = Toplevel(master)
        self.top.title(f"🔗 관계 정보 - {'RewardGroupID' if is_group else 'RewardID'}: {reward_id}")
        self.top.geometry("900x600")
        self.folder = folder
        self.db_folder = db_folder
        self.reward_id = int(reward_id)
        self.is_group = is_group
        self.typecode_mapping = typecode_mapping
        
        self.build_ui()
        self.load_relation_data()
                
    def build_ui(self):
        # 상단 정보 프레임
        info_frame = tk.Frame(self.top)
        info_frame.pack(fill="x", padx=10, pady=5)
        
        column_name = "RewardGroupID" if self.is_group else "RewardID"
        tk.Label(info_frame, text=f"{column_name}: {self.reward_id}의 관계 정보", 
                font=("Helvetica", 12, "bold")).pack(anchor="w")
        
        # 관계 유형 선택 프레임
        filter_frame = tk.Frame(self.top)
        filter_frame.pack(fill="x", padx=10, pady=5)
        
        self.relation_type_var = tk.StringVar(value="모든 관계")
        relation_types = ["모든 관계", "외래키"]
        if self.is_group:
            relation_types.extend(["RewardType별 관계"])
            
        relation_type_label = tk.Label(filter_frame, text="관계 유형:")
        relation_type_label.pack(side="left")
        
        relation_type_combo = ttk.Combobox(filter_frame, textvariable=self.relation_type_var,
                                         values=relation_types, state="readonly", width=15)
        relation_type_combo.pack(side="left", padx=5)
        
        refresh_btn = tk.Button(filter_frame, text="새로고침", command=self.load_relation_data)
        refresh_btn.pack(side="left", padx=5)
        
        # 관계 정보 트리뷰
        tree_frame = tk.Frame(self.top)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 스크롤바
        y_scrollbar = tk.Scrollbar(tree_frame)
        y_scrollbar.pack(side="right", fill="y")
        
        x_scrollbar = tk.Scrollbar(tree_frame, orient="horizontal")
        x_scrollbar.pack(side="bottom", fill="x")
        
        # 트리뷰
        columns = ("관계 유형", "원본 테이블", "원본 컬럼", "대상 테이블", "대상 컬럼", "필터 조건", "참조값")
        self.relation_tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                         yscrollcommand=y_scrollbar.set,
                                         xscrollcommand=x_scrollbar.set)
        
        for col in columns:
            self.relation_tree.heading(col, text=col)
            self.relation_tree.column(col, anchor="w", width=120)
            
        self.relation_tree.pack(fill="both", expand=True)
        
        y_scrollbar.config(command=self.relation_tree.yview)
        x_scrollbar.config(command=self.relation_tree.xview)
        
        # 선택 정보
        self.status_label = tk.Label(self.top, text="", anchor="w")
        self.status_label.pack(fill="x", padx=10, pady=5)
        
        # 이벤트 바인딩
        self.relation_tree.bind("<Double-1>", self.open_target_excel)
        self.relation_type_var.trace("w", lambda *args: self.filter_relation_tree())
        
    def load_relation_data(self):
        """관계 정보를 로드하여 트리뷰에 표시합니다."""
        # 트리뷰 초기화
        self.relation_tree.delete(*self.relation_tree.get_children())
        self.status_label.config(text="🔍 관계 정보 검색 중...")
        
        # 스레드로 실행
        threading.Thread(target=self._load_relation_data_thread, daemon=True).start()
        
    def _load_relation_data_thread(self):
        """관계 정보를 백그라운드에서 로드합니다."""
        try:
            relations = []
            
            # 1. RewardGroupTemplate 테이블의 관계 정보 로드
            if self.is_group:
                # GroupID로 검색하는 경우
                db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
                if os.path.exists(db_path):
                    conn = sqlite3.connect(db_path)
                    # GroupID에 해당하는 레코드 조회
                    query = "SELECT * FROM RewardGroupTemplate WHERE RewardGroupID = ?"
                    df = pd.read_sql_query(query, conn, params=(self.reward_id,))
                    conn.close()
                    
                    if not df.empty:
                        # 각 레코드별 관계 정보 생성
                        for _, row in df.iterrows():
                            reward_type = row.get("RewardType", 0)
                            reward_id = row.get("RewardID", 0)
                            
                            # typecode_mapping에서 관련 매핑 찾기
                            for mapping in self.typecode_mapping:
                                if (mapping["source_table"] == "RewardGroupTemplate" and 
                                    mapping["filter_column"] == "RewardType" and 
                                    float(mapping["filter_value"]) == float(reward_type)):
                                    
                                    # 관계 정보 생성
                                    relation = {
                                        "relation_type": mapping["relation_type"],
                                        "source_table": mapping["source_table"],
                                        "source_column": mapping["source_column"],
                                        "target_table": mapping["target_table"],
                                        "target_column": mapping["target_column"],
                                        "filter_condition": f"{mapping['filter_column']}={mapping['filter_value']}",
                                        "reference_value": reward_id
                                    }
                                    
                                    # 대상 테이블에서 추가 정보 조회 (Name 컬럼이 있으면)
                                    target_info = self._get_target_name(
                                        mapping["target_table"], 
                                        mapping["target_column"], 
                                        reward_id
                                    )
                                    if target_info:
                                        relation["reference_value"] = f"{reward_id} ({target_info})"
                                        
                                    relations.append(relation)
            else:
                # RewardID로 검색하는 경우
                db_path = os.path.join(self.db_folder, "RewardGroupTemplate.db")
                if os.path.exists(db_path):
                    conn = sqlite3.connect(db_path)
                    # RewardID에 해당하는 레코드 조회
                    query = "SELECT * FROM RewardGroupTemplate WHERE RewardID = ?"
                    df = pd.read_sql_query(query, conn, params=(self.reward_id,))
                    conn.close()
                    
                    if not df.empty:
                        # 각 레코드별 관계 정보 생성
                        for _, row in df.iterrows():
                            group_id = row.get("GroupID", 0)
                            reward_type = row.get("RewardType", 0)
                            
                            # 관계 정보 생성 (RewardGroupTemplate -> 다른 테이블)
                            relation = {
                                "relation_type": "참조됨",
                                "source_table": "RewardGroupTemplate",
                                "source_column": "RewardID",
                                "target_table": "RewardGroupTemplate",
                                "target_column": "GroupID",
                                "filter_condition": f"RewardType={reward_type}",
                                "reference_value": group_id
                            }
                            relations.append(relation)
                            
                            # typecode_mapping에서 관련 매핑 찾기 (RewardID가 대상인 경우)
                            for mapping in self.typecode_mapping:
                                if (mapping["target_table"] == "RewardGroupTemplate" and 
                                    mapping["source_column"] == "RewardID"):
                                    
                                    relation = {
                                        "relation_type": mapping["relation_type"],
                                        "source_table": mapping["source_table"],
                                        "source_column": mapping["source_column"],
                                        "target_table": "RewardGroupTemplate",
                                        "target_column": "RewardID",
                                        "filter_condition": "",
                                        "reference_value": self.reward_id
                                    }
                                    relations.append(relation)
            
            # 2. 관계 정보를 트리뷰에 추가
            for relation in relations:
                values = (
                    relation["relation_type"],
                    relation["source_table"],
                    relation["source_column"],
                    relation["target_table"],
                    relation["target_column"],
                    relation["filter_condition"],
                    relation["reference_value"]
                )
                final_values = values  # 값 복사
                self.top.after(0, lambda v=final_values: self.relation_tree.insert("", "end", values=v))
   
            # 3. 상태 업데이트
            status_msg = f"✅ 총 {len(relations)}개의 관계 정보가 로드되었습니다."
            self.top.after(0, lambda: self.status_label.config(text=status_msg))
            
        except Exception as e:
            # 오류 메시지 미리 생성
            error_msg = f"❌ 오류 발생: {str(e)}"
            self.top.after(0, lambda: self.status_label.config(text=error_msg))
                
    def _get_target_name(self, target_table, target_column, target_id):
        """대상 테이블에서 이름 정보를 조회합니다."""
        try:
            db_path = os.path.join(self.db_folder, f"{target_table}.db")
            if not os.path.exists(db_path):
                return None
                
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 테이블에 Name 컬럼이 있는지 확인
            cursor.execute(f"PRAGMA table_info({target_table})")
            columns = [row[1] for row in cursor.fetchall()]
            
            name_column = None
            for col in ["Name", "DESC", "Description", "Title"]:
                if col in columns:
                    name_column = col
                    break
            
            if name_column:
                # ID와 이름 조회
                cursor.execute(
                    f"SELECT {name_column} FROM {target_table} WHERE {target_column} = ?", 
                    (target_id,)
                )
                result = cursor.fetchone()
                conn.close()
                
                if result and result[0]:
                    return result[0]
            
            conn.close()
            return None
        except Exception as e:
            print(f"대상 테이블 정보 조회 오류: {e}")
            return None
    def open_target_excel(self, event):
        """대상 테이블 관련 엑셀 파일을 엽니다."""
        item = self.relation_tree.focus()
        if not item:
            return
            
        values = self.relation_tree.item(item, "values")
        if len(values) < 7:
            return
            
        target_table = values[3]
        target_column = values[4]
        reference_value = values[6]
        
        # 참조 값이 있으면 숫자 부분만 추출
        if reference_value and "(" in reference_value:
            reference_value = reference_value.split("(")[0].strip()
        
        try:
            # 엑셀 파일 열기 시도
            # 캐시가 없는 경우 DB 파일을 직접 열어서 정보 표시
            db_path = os.path.join(self.db_folder, f"{target_table}.db")
            if os.path.exists(db_path):
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute(f"SELECT * FROM {target_table} WHERE {target_column} = ?", (reference_value,))
                rows = cursor.fetchall()
                
                # 컬럼 이름 가져오기
                cursor.execute(f"PRAGMA table_info({target_table})")
                columns = [row[1] for row in cursor.fetchall()]
                conn.close()
                
                if rows:
                    # 결과 팝업 표시
                    result_popup = Toplevel(self.top)
                    result_popup.title(f"테이블 데이터: {target_table}")
                    result_popup.geometry("800x400")
                    
                    # 결과 트리뷰
                    tree_frame = tk.Frame(result_popup)
                    tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
                    
                    result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
                    for col in columns:
                        result_tree.heading(col, text=col)
                        result_tree.column(col, width=100)
                    
                    result_tree.pack(fill="both", expand=True)
                    
                    # 데이터 추가
                    for row in rows:
                        result_tree.insert("", "end", values=row)
                    
                    self.status_label.config(text=f"✅ {target_table} 테이블 데이터를 조회했습니다.")
                else:
                    self.status_label.config(text=f"⚠️ {target_table} 테이블에서 {reference_value} 값을 찾을 수 없습니다.")
            else:
                self.status_label.config(text=f"⚠️ {target_table}.db 파일을 찾을 수 없습니다.")
        except Exception as e:
            self.status_label.config(text=f"❌ 오류 발생: {e}")
    
    def filter_relation_tree(self):
        """관계 유형에 따라 트리뷰를 필터링합니다."""
        try:
            filter_type = self.relation_type_var.get()
            
            # 현재 숨겨진 항목 저장
            hidden_items = []
            for item in self.relation_tree.detached():
                hidden_items.append(item)
            
            # 모든 항목 표시
            for item in hidden_items:
                self.relation_tree.reattach(item, "", "end")
            
            # 선택된 필터에 따라 항목 숨기기
            if filter_type != "모든 관계":
                to_detach = []
                for item in self.relation_tree.get_children():
                    values = self.relation_tree.item(item, "values")
                    relation_type = values[0]
                    
                    if filter_type == "외래키" and relation_type != "외래키":
                        to_detach.append(item)
                    elif filter_type == "RewardType별 관계":
                        # RewardType 필터 조건이 있는지 확인
                        filter_condition = values[5]
                        if not filter_condition.startswith("RewardType"):
                            to_detach.append(item)
                
                # 항목 숨기기
                for item in to_detach:
                    self.relation_tree.detach(item)
                    
            self.status_label.config(text=f"✅ 필터 적용: {filter_type}")
        except Exception as e:
            # 오류 메시지 표시 (람다 없이)
            self.status_label.config(text=f"❌ 필터 오류: {str(e)}")