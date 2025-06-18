# quest_search_popup.py
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
import pandas as pd
import sqlite3
import threading

from utils.excel_utils import ExcelFileManager
from utils.config_utils import load_search_history, save_search_history

class QuestSearchPopup:
    def __init__(self, master, folder, db_folder, excel_cache=None):
        self.top = tk.Toplevel(master)
        self.top.title("🔍 퀘스트 검색기")
        self.top.geometry("1200x700")
        self.folder = folder
        self.db_folder = db_folder
        self.cache = excel_cache
        self._quest_detached = []
        self.quest_type_mapping = self.load_quest_type_mapping()
        self.mission_type_mapping = self.load_mission_type_mapping()
        self.table_relationships = self.load_table_relationships()
        
        # 검색 히스토리
        self.quest_id_history = load_search_history("quest_id")
        self.quest_name_history = load_search_history("quest_name")
        
        # UI 구성
        self._build_ui()
        
    def load_quest_type_mapping(self):
        """QuestType 매핑 정보 로드"""
        try:
            # 기본 매핑 정의
            quest_types = {
                "1": "메인 퀘스트",
                "2": "스토리 이벤트",
                "4": "일반 업적",
                "5": "대륙 업적",
                "11": "일일 미션",
                "12": "주간 미션",
                "21": "가이드 퀘스트",
                "31": "알선소 퀘스트",
                "41": "레벨 패스 일일",
                "42": "레벨 패스 주간",
                "61": "이벤트 미션",
            }
            
            # typecode_mapping.json에서 추가 정보 확인
            mapping_path = os.path.join(self.folder, "typecode_mapping.json")
            if not os.path.exists(mapping_path):
                mapping_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "typecode_mapping.json")
            
            if os.path.exists(mapping_path):
                with open(mapping_path, 'r', encoding='utf-8') as f:
                    typecode_data = json.load(f)
                    # QuestType 관련 매핑 추가 처리 (필요시)
            
            return quest_types
        except Exception as e:
            print(f"QuestType 매핑 로드 오류: {e}")
            return {}
            
    def load_mission_type_mapping(self):
        """QuestMissionType 매핑 정보 로드"""
        try:
                    # 기본 매핑 정의
            mission_types = {
                "1": "몬스터 처치 ",
                "3": "스테이지 미션 클리어",
                "10": "스테이지 클리어",
                "11": "스테이지 플레이",
                "12": "결투장 승리하기",
                "13": "보스전 단계 클리어",
                "14": "미믹 클리어하기",
                "15": "시간의 균열 클리어",
                "19": "배치 변경.",
                "101": "영웅 습득",
                "102": "영웅 랭크 달성",
                "103": "영웅 잠재력 달성",
                "104": "영웅 스킬 강화",
                "106": "영웅 레벨 달성",
                "107": "영웅진화 달성",
                "108": "영웅초월 달성",
                "109": "영웅레벨업 시도",
                "201": "대장간 슬롯 강화 시도",
                "202": "대장간 슬롯 강화 성공",
                "203": "대장간 슬롯 강화 실패 ",
                "204": "대장간 슬롯 연마 단계",
                "205": "장비 초월 시도",
                "206": "장비 초월 달성",
                "207": "장비 품질 시도",
                "208": "장비 품질 달성",
                "209": "장비 세공하기",
                "210": "장비 장착하기",
                "211": "구매후 장비 장착",
                "301": "아이템 습득",
                "302": "아이템 습득",
                "303": "아이템 습득",
                "304": "아이템 분해",
                "305": "아이템 제작",
                "306": "지식 획득",
                "307": "지식 등록",
                "308": "지식 장착",
                "309": "도구 사용",
                "401": "마을 입장",
                "402": "지역 입장",
                "403": "지역 퇴장",
                "404": "npc 인터렉션",
                "405": "상점 구매",
                "406": "상점 판매",
                "502": "길드 가입",
                "504": "친구포인트 선물하기 ",
                "505": "길드 출석",
                "510": "소환(영웅)",
                "511": "소환(장비)",
                "601": "퀘스트완료",
                "602": "특정 타입 퀘스트 수행",
                "604": "튜토리얼 클리어",
                "606": "스토리 모드 클리어",
                "701": "재화 획득",
                "702": "재화 사용",
                "703": "티켓 사용",
                "802": "접속 일수",
                "803": "계정연동",
                "901": "미니게임 - 수리하기",
                "1000": "스테이지 스즌별 달성도",
                "1099": "StageSeason_End",
                "10001": "레이드 클리어 ( ) 스코어",
                "10002": "특정 영웅 전투력 달성",
                "10003": "시리즈별 영웅 전투력",
                "10004": "팀 전투력 달성",
                "10005": "계정 레벨",
                "10006": "스테이지 클리어",
                "10007": "메인 퀘스트 클리어",
                "10008": "영웅 랭크업",
                "10009": "영웅 잠재력 달성",
                "10010": "영웅 레벨 달성",
                "10011": "영웅진화 달성",
                "10012": "영웅초월 달성",
                "10013": "보스전 단계 클리어",
                "10014": "길드 레이드 클리어시 ( ) 점수",
                "10100": "영웅 소환 누적  - MissionType(510)",
                "10101": "장비 소환 누적  - MissionType(511)",
                "10102": "영웅 + 장비 누적  - 510 + 511",
                "10200": "별 갯수 획득 ( 1 : 시리즈(3,4,5) ) 습득 별갯수",
                "10201": "별 갯수 획득 통합",
                "30001": "스테이지 클리어",
                "30002": "영웅 생존",
                "30003": "특정 시간내에 스테이지 클리어",
                "30004": "특정 웨이브 클리어",
                "30005": "아이템 사용 ",
                "30006": "해당 시리즈 영웅만 사용."
            }
            
            # Excel 파일에서 추가 정의 로드 (필요시)
            # TODO: Excel 파일 참조 구현
            
            return mission_types
        except Exception as e:
            print(f"MissionType 매핑 로드 오류: {e}")
            return {}
    
    def load_table_relationships(self):
        """table_relationships.json 파일을 로드합니다."""
        try:
            # 파일 경로 탐색
            rel_path = os.path.join(self.folder, "table_relationships.json")
            if not os.path.exists(rel_path):
                rel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "table_relationships.json")
            
            if os.path.exists(rel_path):
                with open(rel_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            
            print(f"테이블 관계 파일을 찾을 수 없습니다: {rel_path}")
            return {}
        except Exception as e:
            print(f"테이블 관계 로드 오류: {e}")
            return {}
    
    def _build_ui(self):
        """UI 구성"""
        # 검색 영역
        search_frame = tk.Frame(self.top)
        search_frame.pack(fill="x", padx=10, pady=5)
        
        # ID로 검색
        id_frame = tk.Frame(search_frame)
        id_frame.pack(side="left", padx=5)
        
        tk.Label(id_frame, text="퀘스트 ID:").pack(side="left")
        self.id_entry = tk.Entry(id_frame, width=15)
        self.id_entry.pack(side="left", padx=5)
        
        # 이름으로 검색
        name_frame = tk.Frame(search_frame)
        name_frame.pack(side="left", padx=5)
        
        tk.Label(name_frame, text="퀘스트 이름:").pack(side="left")
        self.name_entry = tk.Entry(name_frame, width=25)
        self.name_entry.pack(side="left", padx=5)
        
        # 타입으로 필터링
        type_frame = tk.Frame(search_frame)
        type_frame.pack(side="left", padx=5)
        
        tk.Label(type_frame, text="퀘스트 타입:").pack(side="left")
        quest_types = ["전체"] + [f"{k}:{v}" for k, v in self.quest_type_mapping.items()]
        self.type_combo = ttk.Combobox(type_frame, values=quest_types, width=15, state="readonly")
        self.type_combo.current(0)
        self.type_combo.pack(side="left", padx=5)
        
        # 검색 버튼
        search_btn = tk.Button(search_frame, text="검색", 
                             command=self._run_search)
        search_btn.pack(side="left", padx=5)
        
        # 히스토리 프레임
        history_frame = tk.Frame(self.top)
        history_frame.pack(fill="x", padx=10, pady=5)
        
        # ID 히스토리
        id_history_frame = tk.Frame(history_frame)
        id_history_frame.pack(side="left", fill="y", padx=5)
        
        tk.Label(id_history_frame, text="ID 검색 기록").pack(anchor="w")
        self.id_history_listbox = tk.Listbox(id_history_frame, height=4, width=15)
        self.id_history_listbox.pack(fill="x")
        self.id_history_listbox.bind("<<ListboxSelect>>", 
                                    lambda e: self._on_history_select(e, self.id_entry, True))
        
        # 이름 히스토리
        name_history_frame = tk.Frame(history_frame)
        name_history_frame.pack(side="left", fill="y", padx=5)
        
        tk.Label(name_history_frame, text="이름 검색 기록").pack(anchor="w")
        self.name_history_listbox = tk.Listbox(name_history_frame, height=4, width=25)
        self.name_history_listbox.pack(fill="x")
        self.name_history_listbox.bind("<<ListboxSelect>>", 
                                      lambda e: self._on_history_select(e, self.name_entry, False))
        
        # 히스토리 삭제 버튼
        delete_frame = tk.Frame(history_frame)
        delete_frame.pack(side="right", padx=5)
        
        tk.Button(delete_frame, text="ID 기록 삭제", 
                command=lambda: self._delete_history(True)).pack(pady=2)
        tk.Button(delete_frame, text="이름 기록 삭제", 
                command=lambda: self._delete_history(False)).pack(pady=2)
        
        # 결과 영역
        result_frame = tk.Frame(self.top)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 결과 목록 (왼쪽)
        list_frame = tk.Frame(result_frame)
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # 필터 프레임
        filter_frame = tk.Frame(list_frame)
        filter_frame.pack(fill="x", pady=5)
        
        tk.Label(filter_frame, text="결과 필터:").pack(side="left")
        self.filter_var = tk.StringVar()
        filter_entry = tk.Entry(filter_frame, textvariable=self.filter_var, width=20)
        filter_entry.pack(side="left", padx=5)
        
        tk.Button(filter_frame, text="필터 적용", 
                command=self._apply_filter).pack(side="left", padx=2)
        tk.Button(filter_frame, text="필터 초기화", 
                command=self._reset_filter).pack(side="left", padx=2)
        
        # 트리뷰 (결과 목록)
        tree_frame = tk.Frame(list_frame)
        tree_frame.pack(fill="both", expand=True)
        
        columns = ("ID", "QuestType", "퀘스트 이름", "상태")
        self.result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # 컬럼 설정
        self.result_tree.heading("ID", text="ID")
        self.result_tree.column("ID", width=70, anchor="w")
        self.result_tree.heading("QuestType", text="타입")
        self.result_tree.column("QuestType", width=100, anchor="w")
        self.result_tree.heading("퀘스트 이름", text="퀘스트 이름")
        self.result_tree.column("퀘스트 이름", width=250, anchor="w")
        self.result_tree.heading("상태", text="상태")
        self.result_tree.column("상태", width=100, anchor="w")
        
        self.result_tree.pack(side="left", fill="both", expand=True)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        # 이벤트 바인딩
        self.result_tree.bind("<<TreeviewSelect>>", self._on_quest_select)
        self.result_tree.bind("<Double-1>", self._open_excel)
        
        # 상세 정보 (오른쪽)
        self.detail_frame = QuestDetailFrame(result_frame, self)
        
        # 상태 표시
        self.status_var = tk.StringVar()
        status_bar = tk.Label(self.top, textvariable=self.status_var, anchor="w")
        status_bar.pack(fill="x", padx=10, pady=5)
        
        # 초기화
        self._load_history()
    
    def _load_history(self):
        """검색 히스토리 로드"""
        # ID 히스토리
        self.id_history_listbox.delete(0, tk.END)
        for item in self.quest_id_history:
            self.id_history_listbox.insert(tk.END, item)
        
        # 이름 히스토리
        self.name_history_listbox.delete(0, tk.END)
        for item in self.quest_name_history:
            self.name_history_listbox.insert(tk.END, item)
    
    def _on_history_select(self, event, entry, is_id):
        """히스토리 항목 선택 시 처리"""
        listbox = event.widget
        selection = listbox.curselection()
        if selection:
            value = listbox.get(selection[0])
            entry.delete(0, tk.END)
            entry.insert(0, value)
    
    def _delete_history(self, is_id):
        """히스토리 삭제"""
        if is_id:
            self.quest_id_history.clear()
            save_search_history(self.quest_id_history, "quest_id")
            self.id_history_listbox.delete(0, tk.END)
        else:
            self.quest_name_history.clear()
            save_search_history(self.quest_name_history, "quest_name")
            self.name_history_listbox.delete(0, tk.END)
    
    def _update_history(self, keyword, is_id):
        """검색 히스토리 업데이트"""
        if not keyword:
            return
            
        if is_id:
            # ID 히스토리 업데이트
            if keyword in self.quest_id_history:
                self.quest_id_history.remove(keyword)
            self.quest_id_history.insert(0, keyword)
            self.quest_id_history = self.quest_id_history[:10]  # 최대 10개
            save_search_history(self.quest_id_history, "quest_id")
        else:
            # 이름 히스토리 업데이트
            if keyword in self.quest_name_history:
                self.quest_name_history.remove(keyword)
            self.quest_name_history.insert(0, keyword)
            self.quest_name_history = self.quest_name_history[:10]  # 최대 10개
            save_search_history(self.quest_name_history, "quest_name")
        
        # 히스토리 목록 갱신
        self._load_history()
    
    def _run_search(self):
        """검색 실행"""
        # 검색 조건 가져오기
        quest_id = self.id_entry.get().strip()
        quest_name = self.name_entry.get().strip()
        quest_type = self.type_combo.get().strip()
        
        # 검색 조건 검증
        if not quest_id and not quest_name:
            messagebox.showwarning("검색 조건", "퀘스트 ID 또는 이름을 입력하세요.")
            return
        
        # 히스토리 업데이트
        if quest_id:
            self._update_history(quest_id, True)
        if quest_name:
            self._update_history(quest_name, False)
        
        # 타입 필터 처리
        type_filter = None
        if quest_type != "전체":
            type_filter = quest_type.split(":")[0]
        
        # 상태 표시
        self.status_var.set("🔍 검색 중...")
        self.result_tree.delete(*self.result_tree.get_children())
        
        # 스레드로 검색 실행
        threading.Thread(
            target=self._search_quest,
            args=(quest_id, quest_name, type_filter),
            daemon=True
        ).start()
    
    def _search_quest(self, quest_id, quest_name, type_filter):
        """퀘스트 검색 실행 (스레드에서 실행)"""
        try:
            # DB 연결
            db_path = os.path.join(self.db_folder, "QuestTemplate.db")
            if not os.path.exists(db_path):
                self.top.after(0, lambda: self.status_var.set("❌ QuestTemplate.db 파일을 찾을 수 없습니다."))
                return
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # SQL 쿼리 구성
            params = []
            sql = "SELECT * FROM QuestTemplate WHERE 1=1"
            
            if quest_id:
                sql += " AND TemplateID = ?"
                params.append(int(quest_id))
            
            if quest_name:
                # String 테이블 참조 (String_Quest_Name)
                string_db_path = os.path.join(self.db_folder, "String_Quest_Name.db")
                if os.path.exists(string_db_path):
                    string_conn = sqlite3.connect(string_db_path)
                    string_cursor = string_conn.cursor()
                    string_cursor.execute("SELECT STRING_ID FROM String_Quest_Name WHERE DESCRIPTION LIKE ?", (f"%{quest_name}%",))
                    string_ids = [row[0] for row in string_cursor.fetchall()]
                    string_conn.close()
                    
                    if string_ids:
                        placeholders = ", ".join(["?"] * len(string_ids))
                        sql += f" AND TemplateID IN ({placeholders})"
                        params.extend(string_ids)
                    else:
                        # 이름 검색 결과가 없는 경우
                        self.top.after(0, lambda: self.status_var.set("❌ 검색 결과가 없습니다."))
                        return
                else:
                    # String DB가 없으면 모든 퀘스트에서 검색
                    sql += " AND TemplateID LIKE ?"
                    params.append(f"%{quest_name}%")
            
            if type_filter:
                sql += " AND QuestType = ?"
                params.append(int(type_filter))
            
            # 쿼리 실행
            cursor.execute(sql, params)
            quests = cursor.fetchall()
            
            # 컬럼 이름 가져오기
            column_names = [description[0] for description in cursor.description]
            
            # 결과가 없는 경우
            if not quests:
                self.top.after(0, lambda: self.status_var.set("❌ 검색 결과가 없습니다."))
                return
            
            # 결과 처리 및 UI 업데이트
            results = []
            for quest in quests:
                # 딕셔너리로 변환
                quest_dict = {column_names[i]: value for i, value in enumerate(quest)}
                
                # 이름 가져오기
                quest_name = self._get_quest_name(quest_dict["TemplateID"])
                
                # 타입 이름 가져오기
                quest_type = self.quest_type_mapping.get(str(quest_dict["QuestType"]), f"타입 {quest_dict['QuestType']}")
                
                # 결과 추가
                results.append({
                    "id": quest_dict["TemplateID"],
                    "type": quest_type,
                    "name": quest_name,
                    "status": "활성" if quest_dict.get("IsActive", 1) == 1 else "비활성",
                    "data": quest_dict
                })
            
            # UI 업데이트 (메인 스레드에서)
            self.top.after(0, lambda: self._update_results(results))
            
            conn.close()
        except Exception as e:
            self.top.after(0, lambda: self.status_var.set(f"❌ 검색 오류: {str(e)}"))
    
    def _get_quest_name(self, template_id):
        """퀘스트 이름 가져오기"""
        try:
            string_db_path = os.path.join(self.db_folder, "String_Quest_Name.db")
            if os.path.exists(string_db_path):
                conn = sqlite3.connect(string_db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT DESCRIPTION FROM String_Quest_Name WHERE STRING_ID = ?", (template_id,))
                result = cursor.fetchone()
                conn.close()
                
                if result:
                    return result[0]
            
            return f"퀘스트 {template_id}"
        except Exception as e:
            print(f"퀘스트 이름 조회 오류: {e}")
            return f"퀘스트 {template_id}"
    
    def _update_results(self, results):
        """검색 결과 UI 업데이트"""
        # 트리뷰 초기화
        self.result_tree.delete(*self.result_tree.get_children())
        
        # 결과 추가
        for result in results:
            self.result_tree.insert("", "end", 
                                  values=(result["id"], result["type"], result["name"], result["status"]), 
                                  tags=(str(result["id"]),))
        
        # 상태 표시 업데이트
        self.status_var.set(f"✅ 검색 완료: {len(results)}개의 퀘스트를 찾았습니다.")
        
        # 첫 번째 항목 선택
        if self.result_tree.get_children():
            first_item = self.result_tree.get_children()[0]
            self.result_tree.selection_set(first_item)
            self.result_tree.focus(first_item)
            self._on_quest_select(None)
    
    def _apply_filter(self):
        """결과 필터링"""
        keyword = self.filter_var.get().strip().lower()
        if not keyword:
            self._reset_filter()
            return
        
        # 필터링
        self._quest_detached.clear()
        for item in self.result_tree.get_children():
            values = self.result_tree.item(item, "values")
            if not any(keyword in str(v).lower() for v in values):
                self._quest_detached.append(item)
                self.result_tree.detach(item)
        
        # 상태 표시 업데이트
        visible_count = len(self.result_tree.get_children()) - len(self._quest_detached)
        self.status_var.set(f"✅ 필터 적용: {visible_count}개 표시 중")
    
    def _reset_filter(self):
        """필터 초기화"""
        # 모든 항목 다시 표시
        for item in self._quest_detached:
            self.result_tree.reattach(item, "", "end")
        self._quest_detached.clear()
        
        # 필터 입력창 초기화
        self.filter_var.set("")
        
        # 상태 표시 업데이트
        self.status_var.set(f"✅ 필터 초기화: {len(self.result_tree.get_children())}개 표시 중")
    
    def _on_quest_select(self, event):
        """퀘스트 선택 시 처리"""
        selection = self.result_tree.selection()
        if not selection:
            return
        
        # 선택된 퀘스트 ID 가져오기
        item = selection[0]
        quest_id = self.result_tree.item(item, "tags")[0]
        
        # 퀘스트 상세 정보 로드
        self.detail_frame.load_quest_details(quest_id)
    
    def _open_excel(self, event):
        """엑셀 파일 열기"""
        selection = self.result_tree.selection()
        if not selection:
            return
        
        # 선택된 퀘스트 ID 가져오기
        item = selection[0]
        quest_id = self.result_tree.item(item, "tags")[0]
        
        # 엑셀 파일 찾기
        for file_path, file_info in self.cache.items():
            for sheet_name, sheet_info in file_info.get("sheets", {}).items():
                if "has_quest" in sheet_info and sheet_info["has_quest"]:
                    # 엑셀 파일 열기
                    try:
                        full_path = os.path.join(self.folder, file_path)
                        ExcelFileManager.highlight_excel_by_value(full_path, sheet_name, "TemplateID", quest_id)
                        return
                    except Exception as e:
                        print(f"엑셀 열기 오류: {e}")
        
        # 엑셀 파일을 찾지 못한 경우
        messagebox.showinfo("정보", "해당 퀘스트의 엑셀 파일을 찾을 수 없습니다.")
    
    def get_quest_data(self, quest_id):
        """퀘스트 데이터 가져오기"""
        try:
            db_path = os.path.join(self.db_folder, "QuestTemplate.db")
            if not os.path.exists(db_path):
                return None
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 퀘스트 정보 조회
            cursor.execute("SELECT * FROM QuestTemplate WHERE TemplateID = ?", (quest_id,))
            quest = cursor.fetchone()
            
            if not quest:
                conn.close()
                return None
            
            # 컬럼 이름 가져오기
            column_names = [description[0] for description in cursor.description]
            
            # 딕셔너리로 변환
            quest_dict = {column_names[i]: value for i, value in enumerate(quest)}
            
            conn.close()
            return quest_dict
        except Exception as e:
            print(f"퀘스트 데이터 조회 오류: {e}")
            return None
    
    def get_quest_missions(self, quest_id):
        """퀘스트 미션 데이터 가져오기"""
        try:
            db_path = os.path.join(self.db_folder, "QuestMissionTemplate.db")
            if not os.path.exists(db_path):
                return []
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 미션 정보 조회
            cursor.execute("SELECT * FROM QuestMissionTemplate WHERE TemplateID = ?", (quest_id,))
            missions = cursor.fetchall()
            
            if not missions:
                conn.close()
                return []
            
            # 컬럼 이름 가져오기
            column_names = [description[0] for description in cursor.description]
            
            # 결과 변환
            result = []
            for mission in missions:
                mission_dict = {column_names[i]: value for i, value in enumerate(mission)}
                
                # 미션 타입 이름 변환
                mission_type = mission_dict.get("MissionType", 0)
                type_name = self.mission_type_mapping.get(str(mission_type), f"미션 타입 {mission_type}")
                mission_dict["MissionTypeName"] = type_name
                
                result.append(mission_dict)
            
            conn.close()
            return result
        except Exception as e:
            print(f"퀘스트 미션 조회 오류: {e}")
            return []
    
    def find_quest_references(self, quest_id):
        """퀘스트 참조 정보 찾기"""
        references = {
            "event_direction": [],
            "condition_template": [],
            "other_quests": []
        }
        
        try:
            # 1. EventDirection 테이블에서 참조 찾기
            event_db_path = os.path.join(self.db_folder, "EventDirection.db")
            if os.path.exists(event_db_path):
                conn = sqlite3.connect(event_db_path)
                cursor = conn.cursor()
                
                # RequireType이 40, 41, 42인 경우 참조
                cursor.execute("""
                    SELECT * FROM EventDirection 
                    WHERE (RequireType IN (40, 41, 42) AND RequireOption = ?) 
                    OR (HideType IN (40, 41, 42) AND HideOption = ?)
                """, (quest_id, quest_id))
                
                events = cursor.fetchall()
                
                # 컬럼 이름 가져오기
                if events:
                    column_names = [description[0] for description in cursor.description]
                    for event in events:
                        event_dict = {column_names[i]: value for i, value in enumerate(event)}
                        references["event_direction"].append(event_dict)
                
                conn.close()
            
            # 2. ConditionTemplate 테이블에서 참조 찾기
            condition_db_path = os.path.join(self.db_folder, "ConditionTemplate.db")
            if os.path.exists(condition_db_path):
                conn = sqlite3.connect(condition_db_path)
                cursor = conn.cursor()
                
                # ConditionType이 1020, 1021인 경우 참조
                cursor.execute("""
                    SELECT * FROM ConditionTemplate 
                    WHERE (ConditionType IN (1020, 1021) AND Condition1 = ?)
                """, (quest_id,))
                
                conditions = cursor.fetchall()
                
                # 컬럼 이름 가져오기
                if conditions:
                    column_names = [description[0] for description in cursor.description]
                    for condition in conditions:
                        condition_dict = {column_names[i]: value for i, value in enumerate(condition)}
                        references["condition_template"].append(condition_dict)
                
                conn.close()
            
            # 3. QuestTemplate 테이블에서 참조 찾기
            quest_db_path = os.path.join(self.db_folder, "QuestTemplate.db")
            if os.path.exists(quest_db_path):
                conn = sqlite3.connect(quest_db_path)
                cursor = conn.cursor()
                
                # OpenTemplateID 또는 NeedTemplateID가 quest_id인 경우
                cursor.execute("""
                    SELECT * FROM QuestTemplate 
                    WHERE OpenTemplateID = ? OR NeedTemplateID = ?
                """, (quest_id, quest_id))
                
                related_quests = cursor.fetchall()
                
                # 컬럼 이름 가져오기
                if related_quests:
                    column_names = [description[0] for description in cursor.description]
                    for quest in related_quests:
                        quest_dict = {column_names[i]: value for i, value in enumerate(quest)}
                        # 이름 추가
                        quest_dict["QuestName"] = self._get_quest_name(quest_dict["TemplateID"])
                        references["other_quests"].append(quest_dict)
                
                conn.close()
            
            return references
        except Exception as e:
            print(f"퀘스트 참조 검색 오류: {e}")
            return references
        
class QuestDetailFrame:
    """퀘스트 상세 정보 표시 프레임"""
    def __init__(self, parent, quest_search):
        self.frame = tk.Frame(parent, relief=tk.RIDGE, bd=2)
        self.frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.quest_search = quest_search
        
        # 노트북 생성
        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 기본 정보 탭
        self.basic_frame = tk.Frame(self.notebook)
        self.notebook.add(self.basic_frame, text="기본 정보")
        
        # 미션 정보 탭
        self.mission_frame = tk.Frame(self.notebook)
        self.notebook.add(self.mission_frame, text="미션 정보")
        
        # 참조 정보 탭
        self.reference_frame = tk.Frame(self.notebook)
        self.notebook.add(self.reference_frame, text="참조 정보")
        
        # 현재 로드된 퀘스트 ID
        self.current_quest_id = None
    
    def load_quest_details(self, quest_id):
        """퀘스트 상세 정보 로드"""
        self.current_quest_id = quest_id
        
        # 데이터 로드
        quest_data = self.quest_search.get_quest_data(quest_id)
        if not quest_data:
            self._show_error("퀘스트 정보를 찾을 수 없습니다.")
            return
        
        # 각 탭 업데이트
        self._update_basic_info(quest_data)
        self._update_mission_info(quest_id)
        self._update_reference_info(quest_id)
    
    def _show_error(self, message):
        """오류 메시지 표시"""
        # 모든 탭 초기화
        for frame in [self.basic_frame, self.mission_frame, self.reference_frame]:
            for widget in frame.winfo_children():
                widget.destroy()
            tk.Label(frame, text=message, fg="red").pack(pady=20)
    
    def _update_basic_info(self, quest_data):
        """기본 정보 탭 업데이트"""
        # 기존 위젯 제거
        for widget in self.basic_frame.winfo_children():
            widget.destroy()
        
        # 스크롤 영역 생성
        scroll_frame = tk.Frame(self.basic_frame)
        scroll_frame.pack(fill="both", expand=True)
        
        canvas = tk.Canvas(scroll_frame)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        content_frame = tk.Frame(canvas)
        
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 마우스 휠 스크롤 바인딩
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 퀘스트 이름 표시
        quest_name = self.quest_search._get_quest_name(quest_data["TemplateID"])
        quest_type = self.quest_search.quest_type_mapping.get(
            str(quest_data.get("QuestType", 0)), 
            f"타입 {quest_data.get('QuestType', 0)}"
        )
        
        title_frame = tk.Frame(content_frame)
        title_frame.pack(fill="x", pady=5)
        
        tk.Label(title_frame, 
                text=f"[{quest_data['TemplateID']}] {quest_name}", 
                font=("Helvetica", 12, "bold")).pack(anchor="w")
        
        tk.Label(title_frame, 
                text=f"퀘스트 타입: {quest_type}").pack(anchor="w")
        
        # 구분선
        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=5)
        
        # 퀘스트 설명 표시 (String_Quest_Desc에서 가져오기)
        quest_desc = self._get_quest_description(quest_data["TemplateID"])
        if quest_desc:
            desc_frame = tk.Frame(content_frame)
            desc_frame.pack(fill="x", pady=5)
            
            tk.Label(desc_frame, text="퀘스트 설명:", font=("Helvetica", 10, "bold")).pack(anchor="w")
            desc_text = tk.Text(desc_frame, wrap="word", height=3, width=50)
            desc_text.insert("1.0", quest_desc)
            desc_text.config(state="disabled")
            desc_text.pack(fill="x", pady=3)
            
            ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=5)
        
        # 기본 정보 표시
        info_frame = tk.Frame(content_frame)
        info_frame.pack(fill="x", pady=5)
        
        tk.Label(info_frame, text="기본 정보:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        
        # 그리드 레이아웃으로 정보 표시
        grid_frame = tk.Frame(info_frame)
        grid_frame.pack(fill="x", pady=3)
        
        # 표시할 필드와 레이블
        fields = [
            ("GroupID", "그룹 ID"),
            ("QuestType", "퀘스트 타입"),
            ("NeedLevel", "필요 레벨"),
            ("IsActive", "활성 여부"),
            ("OpenTemplateID", "오픈 퀘스트 ID"),
            ("NeedTemplateID", "필요 퀘스트 ID"),
            ("RewardGroupID", "보상 그룹 ID"),
            ("AcceptScenarioID", "수락 시나리오 ID"),
            ("CompleteScenarioID", "완료 시나리오 ID")
        ]
        
        for idx, (field, label) in enumerate(fields):
            row = idx // 2
            col = (idx % 2) * 2
            
            # 레이블
            tk.Label(grid_frame, text=f"{label}:", 
                    width=15, anchor="e").grid(row=row, column=col, sticky="e", padx=5, pady=2)
            
            # 값
            value = quest_data.get(field, "")
            
            # 특수 처리
            if field == "IsActive":
                value = "활성" if value == 1 else "비활성"
            
            # 연결된 퀘스트 ID에 이름 추가
            if field in ["OpenTemplateID", "NeedTemplateID"] and value:
                related_name = self.quest_search._get_quest_name(value)
                value = f"{value} ({related_name})"
            
            tk.Label(grid_frame, text=str(value), 
                    width=25, anchor="w").grid(row=row, column=col+1, sticky="w", padx=5, pady=2)
        
        # 추가 정보가 있으면 표시
        other_fields = [k for k in quest_data.keys() if k not in [f[0] for f in fields] and k != "TemplateID"]
        
        if other_fields:
            ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=5)
            
            other_frame = tk.Frame(content_frame)
            other_frame.pack(fill="x", pady=5)
            
            tk.Label(other_frame, text="추가 정보:", font=("Helvetica", 10, "bold")).pack(anchor="w")
            
            other_grid = tk.Frame(other_frame)
            other_grid.pack(fill="x", pady=3)
            
            for idx, field in enumerate(other_fields):
                row = idx // 2
                col = (idx % 2) * 2
                
                # 레이블
                tk.Label(other_grid, text=f"{field}:", 
                        width=15, anchor="e").grid(row=row, column=col, sticky="e", padx=5, pady=2)
                
                # 값
                value = quest_data.get(field, "")
                tk.Label(other_grid, text=str(value), 
                        width=25, anchor="w").grid(row=row, column=col+1, sticky="w", padx=5, pady=2)
    
    def _get_quest_description(self, template_id):
        """퀘스트 설명 가져오기"""
        try:
            string_db_path = os.path.join(self.quest_search.db_folder, "String_Quest_Desc.db")
            if os.path.exists(string_db_path):
                conn = sqlite3.connect(string_db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT DESCRIPTION FROM String_Quest_Desc WHERE STRING_ID = ?", (template_id,))
                result = cursor.fetchone()
                conn.close()
                
                if result:
                    return result[0]
            
            return ""
        except Exception as e:
            print(f"퀘스트 설명 조회 오류: {e}")
            return ""
    
    def _update_mission_info(self, quest_id):
        """미션 정보 탭 업데이트"""
        # 기존 위젯 제거
        for widget in self.mission_frame.winfo_children():
            widget.destroy()
        
        # 미션 정보 로드
        missions = self.quest_search.get_quest_missions(quest_id)
        
        if not missions:
            tk.Label(self.mission_frame, text="미션 정보가 없습니다.").pack(pady=20)
            return
        
        # 미션 정보 표시
        tk.Label(self.mission_frame, text=f"총 {len(missions)}개의 미션", 
                font=("Helvetica", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        
        # 트리뷰로 미션 목록 표시
        columns = ("MissionID", "MissionType", "MissionTarget", "CountValue")
        mission_tree = ttk.Treeview(self.mission_frame, columns=columns, show="headings")
        
        mission_tree.heading("MissionID", text="미션 ID")
        mission_tree.column("MissionID", width=70, anchor="w")
        mission_tree.heading("MissionType", text="미션 타입")
        mission_tree.column("MissionType", width=120, anchor="w")
        mission_tree.heading("MissionTarget", text="미션 대상")
        mission_tree.column("MissionTarget", width=120, anchor="w")
        mission_tree.heading("CountValue", text="필요 수량")
        mission_tree.column("CountValue", width=80, anchor="w")
        
        mission_tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(self.mission_frame, orient="vertical", command=mission_tree.yview)
        scrollbar.pack(side="right", fill="y")
        mission_tree.configure(yscrollcommand=scrollbar.set)
        
        # 미션 데이터 추가
        for mission in missions:
            target_id = mission.get("TargetID", 0)
            target_name = self._get_target_name(mission.get("MissionType", 0), target_id)
            target_display = f"{target_id} ({target_name})" if target_name else target_id
            
            values = (
                mission.get("MissionID", 0),
                mission.get("MissionTypeName", ""),
                target_display,
                mission.get("CountValue", 0)
            )
            
            mission_tree.insert("", "end", values=values)
        
        # 미션 선택 이벤트 처리
        def on_mission_select(event):
            selection = mission_tree.selection()
            if not selection:
                return
            
            # 선택된 미션 정보 표시
            item = selection[0]
            values = mission_tree.item(item, "values")
            mission_id = values[0]
            
            # 해당 미션 찾기
            selected_mission = next((m for m in missions if str(m.get("MissionID", 0)) == mission_id), None)
            if not selected_mission:
                return
            
            # 미션 상세 정보 표시
            detail_toplevel = tk.Toplevel(self.mission_frame)
            detail_toplevel.title(f"미션 {mission_id} 상세 정보")
            detail_toplevel.geometry("500x400")
            
            # 미션 정보 표시
            detail_frame = tk.Frame(detail_toplevel)
            detail_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # 미션 타입과 타겟 정보
            mission_type = selected_mission.get("MissionType", 0)
            target_id = selected_mission.get("TargetID", 0)
            target_name = self._get_target_name(mission_type, target_id)
            
            tk.Label(detail_frame, 
                    text=f"미션 ID: {mission_id}", 
                    font=("Helvetica", 12, "bold")).pack(anchor="w", pady=(0, 5))
            
            tk.Label(detail_frame, 
                    text=f"미션 타입: {selected_mission.get('MissionTypeName', '')} ({mission_type})").pack(anchor="w")
            
            tk.Label(detail_frame, 
                    text=f"대상 ID: {target_id}").pack(anchor="w")
            
            if target_name:
                tk.Label(detail_frame, 
                        text=f"대상 이름: {target_name}").pack(anchor="w")
            
            tk.Label(detail_frame, 
                    text=f"필요 수량: {selected_mission.get('CountValue', 0)}").pack(anchor="w")
            
            # 추가 정보 표시
            ttk.Separator(detail_frame, orient="horizontal").pack(fill="x", pady=10)
            tk.Label(detail_frame, text="미션 상세 정보:", font=("Helvetica", 10, "bold")).pack(anchor="w")
            
            # 그리드로 모든 정보 표시
            grid_frame = tk.Frame(detail_frame)
            grid_frame.pack(fill="both", expand=True, pady=5)
            
            row = 0
            for key, value in selected_mission.items():
                if key in ["MissionID", "MissionType", "TargetID", "CountValue", "MissionTypeName"]:
                    continue  # 이미 표시된 정보는 제외
                
                tk.Label(grid_frame, text=f"{key}:", 
                        width=15, anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
                
                tk.Label(grid_frame, text=str(value), 
                        width=30, anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
                
                row += 1
        
        mission_tree.bind("<<TreeviewSelect>>", on_mission_select)
    
    def _get_target_name(self, mission_type, target_id):
        """미션 대상 이름 가져오기"""
        try:
            if not target_id:
                return ""
            
            # 미션 타입별 대상 테이블
            target_tables = {
                10: "HeroTemplate",   # NPC 대화
                20: "StageTemplate",  # 스테이지 클리어
                30: "ItemTemplate",   # 아이템 획득
                40: "HeroTemplate",   # 적 처치
                50: "ItemTemplate"    # 아이템 사용
            }
            
            # 테이블 조회
            table = target_tables.get(mission_type)
            if not table:
                return ""
            
            db_path = os.path.join(self.quest_search.db_folder, f"{table}.db")
            if not os.path.exists(db_path):
                return ""
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # ID 컬럼 (테이블별로 다름)
            id_column = "TemplateID" if table in ["StageTemplate", "ItemTemplate"] else "BaseHeroID"
            
            # 이름 가져오기
            cursor.execute(f"SELECT * FROM {table} WHERE {id_column} = ?", (target_id,))
            result = cursor.fetchone()
            
            if not result:
                conn.close()
                return ""
            
            # 컬럼 이름 가져오기
            column_names = [description[0] for description in cursor.description]
            entity_dict = {column_names[i]: value for i, value in enumerate(result)}
            
            conn.close()
            
            # 이름 필드 (테이블별로 다름)
            if table == "HeroTemplate":
                # HeroTemplate의 경우 이름은 String 테이블에 있음
                string_id = target_id
                string_table = "String"
            elif table == "ItemTemplate":
                # ItemTemplate의 경우 이름은 String 테이블에 있음
                string_id = target_id
                string_table = "String"
            elif table == "StageTemplate":
                # StageTemplate의 경우 이름은 해당 테이블에 있음
                return entity_dict.get("StageName", "")
            else:
                return ""
            
            # String 테이블에서 이름 가져오기
            string_db_path = os.path.join(self.quest_search.db_folder, f"{string_table}.db")
            if os.path.exists(string_db_path):
                conn = sqlite3.connect(string_db_path)
                cursor = conn.cursor()
                cursor.execute(f"SELECT DESCRIPTION FROM {string_table} WHERE STRING_ID = ?", (string_id,))
                result = cursor.fetchone()
                conn.close()
                
                if result:
                    return result[0]
            
            return ""
        except Exception as e:
            print(f"대상 이름 조회 오류: {e}")
            return ""
    
    def _update_reference_info(self, quest_id):
        """참조 정보 탭 업데이트"""
        # 기존 위젯 제거
        for widget in self.reference_frame.winfo_children():
            widget.destroy()
        
        # 참조 정보 로드
        references = self.quest_search.find_quest_references(quest_id)
        
        # 참조 정보가 없는 경우
        if not any(references.values()):
            tk.Label(self.reference_frame, text="이 퀘스트를 참조하는 정보가 없습니다.").pack(pady=20)
            return
        
        # 참조 정보 표시
        # 스크롤 영역 생성
        scroll_frame = tk.Frame(self.reference_frame)
        scroll_frame.pack(fill="both", expand=True)
        
        canvas = tk.Canvas(scroll_frame)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        content_frame = tk.Frame(canvas)
        
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 마우스 휠 스크롤 바인딩
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 1. 다른 퀘스트에서의 참조
        if references["other_quests"]:
            tk.Label(content_frame, text="다른 퀘스트에서 참조:", 
                   font=("Helvetica", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
            
            quest_frame = tk.Frame(content_frame)
            quest_frame.pack(fill="x", padx=10, pady=5)
            
            for idx, quest in enumerate(references["other_quests"]):
                relation_type = "선행 퀘스트" if quest.get("NeedTemplateID") == int(quest_id) else "후속 퀘스트"
                
                tk.Label(quest_frame, 
                       text=f"{idx+1}. [{quest['TemplateID']}] {quest.get('QuestName', '')} - {relation_type}").pack(anchor="w")
            
            ttk.Separator(content_frame, orient="horizontal").pack(fill="x", padx=10, pady=5)
        
        # 2. EventDirection에서의 참조
        if references["event_direction"]:
            tk.Label(content_frame, text="이벤트 디렉션에서 참조:", 
                   font=("Helvetica", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
            
            event_frame = tk.Frame(content_frame)
            event_frame.pack(fill="x", padx=10, pady=5)
            
            for idx, event in enumerate(references["event_direction"]):
                relation_type = "필요 조건" if event.get("RequireOption") == int(quest_id) else "숨김 조건"
                group_id = event.get("GroupID", "")
                
                tk.Label(event_frame, 
                       text=f"{idx+1}. 그룹 ID: {group_id} - {relation_type}").pack(anchor="w")
            
            ttk.Separator(content_frame, orient="horizontal").pack(fill="x", padx=10, pady=5)
        
        # 3. ConditionTemplate에서의 참조
        if references["condition_template"]:
            tk.Label(content_frame, text="조건 템플릿에서 참조:", 
                   font=("Helvetica", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
            
            condition_frame = tk.Frame(content_frame)
            condition_frame.pack(fill="x", padx=10, pady=5)
            
            for idx, condition in enumerate(references["condition_template"]):
                template_id = condition.get("TemplateID", "")
                condition_type = condition.get("ConditionType", "")
                
                tk.Label(condition_frame, 
                       text=f"{idx+1}. 템플릿 ID: {template_id} - 조건 타입: {condition_type}").pack(anchor="w")
                
                # 이 조건을 참조하는 다른 테이블 검색 (추가 기능)
                self._add_condition_references(condition_frame, template_id)
            
            ttk.Separator(content_frame, orient="horizontal").pack(fill="x", padx=10, pady=5)
    
    def _add_condition_references(self, parent_frame, condition_id):
        """조건 ID를 참조하는 다른 테이블 정보 추가"""
        # 조건을 참조할 수 있는 테이블과 컬럼
        reference_tables = [
            {"table": "StageTemplate", "column": "UnlockConditionTID"},
            {"table": "MapSpawn", "column": "ShowConditionTID"},
            {"table": "MapSpawn", "column": "HIdeConditionTID"},
            {"table": "MapObject", "column": "ShowConditionTID"},
            {"table": "MapObject", "column": "HIdeConditionTID"},
            {"table": "MapTeleport", "column": "ShowConditionTID"},
            {"table": "MapTeleport", "column": "HIdeConditionTID"},
            {"table": "Tutorial", "column": "UnlockConditionTID"}
        ]
        
        references = []
        
        try:
            for ref in reference_tables:
                db_path = os.path.join(self.quest_search.db_folder, f"{ref['table']}.db")
                if not os.path.exists(db_path):
                    continue
                
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # 참조 검색
                cursor.execute(f"SELECT * FROM {ref['table']} WHERE {ref['column']} = ?", (condition_id,))
                results = cursor.fetchall()
                
                if results:
                    # 컬럼 이름 가져오기
                    column_names = [description[0] for description in cursor.description]
                    
                    for result in results:
                        entity_dict = {column_names[i]: value for i, value in enumerate(result)}
                        
                        # 기본 ID 컬럼 찾기
                        id_column = next((col for col in ["UniqueID", "TemplateID", "StageID", "ObjectID"] 
                                        if col in column_names), None)
                        
                        if id_column:
                            entity_id = entity_dict.get(id_column)
                            references.append({
                                "table": ref["table"],
                                "id": entity_id,
                                "column": ref["column"]
                            })
                
                conn.close()
            
            # 참조 정보 표시
            if references:
                ref_frame = tk.Frame(parent_frame)
                ref_frame.pack(fill="x", padx=(20, 0), pady=3)
                
                tk.Label(ref_frame, text="참조 항목:", font=("Helvetica", 9, "italic")).pack(anchor="w")
                
                for idx, ref in enumerate(references):
                    tk.Label(ref_frame, 
                           text=f"  - {ref['table']} ID: {ref['id']} ({ref['column']})").pack(anchor="w")
        
        except Exception as e:
            print(f"조건 참조 검색 오류: {e}")


if __name__ == "__main__":
    # 테스트용 코드
    root = tk.Tk()
    root.title("퀘스트 검색기 테스트")
    
    folder_path = "./"  # 엑셀 폴더 경로
    db_folder_path = "./"  # DB 폴더 경로
    
    QuestSearchPopup(root, folder_path, db_folder_path)
    
    root.mainloop()