import tkinter as tk
from tkinter import ttk
import os
import sqlite3
import threading
import logging
import time
from utils.cache_utils import hash_paths

class ReplaceUI:
    """치환 관련 UI 및 기능을 담당하는 클래스"""
    
    def __init__(self, parent):
        """
        ReplaceUI 초기화
        
        Args:
            parent: 부모 클래스 (StringSearchPopup)
        """
        self.parent = parent
        self.replace_input_frame = None
        self.word_replace_frame = None
        self.bulk_replace_frame = None
        self.unique_replace_frame = None
        self.replace_lang_frame = None
        self.select_frame = None
        
        # 치환 관련 변수 초기화
        self.replace_from = tk.StringVar()
        self.replace_to = tk.StringVar()
        self.bulk_value = tk.StringVar()
        self.replace_mode = tk.StringVar(value="replace")  # 'replace', 'bulk', 'unique' 중 하나
        self.row_checks = {}  # 체크박스 상태 저장
        self.select_all_var = tk.BooleanVar(value=True)  # 전체 선택 체크박스
        
        # 언어 선택 변수
        self.replace_languages = {
            "ALL": tk.BooleanVar(value=True),
            "KR": tk.BooleanVar(value=False),
            "EN": tk.BooleanVar(value=False),
            "CN": tk.BooleanVar(value=False),
            "TW": tk.BooleanVar(value=False),
            "TH": tk.BooleanVar(value=False),
            "PT": tk.BooleanVar(value=False),
            "ES": tk.BooleanVar(value=False),
            "DE": tk.BooleanVar(value=False),
            "FR": tk.BooleanVar(value=False),
        }
        
        # DB 정보 변수
        self.current_search_var = None
        self.db_info_var = None
    
    def create_replace_input_frame(self, parent_frame):
        """
        치환 입력 프레임 생성
        
        Args:
            parent_frame: 부모 프레임
            
        Returns:
            생성된 치환 입력 프레임
        """
        # 치환 프레임 (두 가지 모드)
        self.replace_input_frame = tk.Frame(parent_frame)
        
        # 1. 단어 치환 모드 UI
        self.word_replace_frame = tk.Frame(self.replace_input_frame)
        self.word_replace_frame.pack(fill="x", padx=5, pady=5)
        
        tk.Label(self.word_replace_frame, text="변경 전:").grid(row=0, column=0, sticky="w")
        replace_from_entry = tk.Entry(self.word_replace_frame, textvariable=self.replace_from, width=30)
        replace_from_entry.grid(row=0, column=1, padx=5, sticky="w")
        
        tk.Label(self.word_replace_frame, text="변경 후:").grid(row=0, column=2, padx=(20, 5), sticky="w")
        replace_to_entry = tk.Entry(self.word_replace_frame, textvariable=self.replace_to, width=30)
        replace_to_entry.grid(row=0, column=3, padx=5, sticky="w")
        
        # 2. 컬럼 일괄 변경 모드 UI
        self.bulk_replace_frame = tk.Frame(self.replace_input_frame)
        # pack은 replace_mode 변경 시 처리
        
        tk.Label(self.bulk_replace_frame, text="변경 값:").pack(side="left")
        bulk_entry = tk.Entry(self.bulk_replace_frame, textvariable=self.bulk_value, width=40)
        bulk_entry.pack(side="left", padx=5)
        
        tk.Label(self.bulk_replace_frame, text="(선택한 컬럼을 이 값으로 일괄 변경합니다. KR 컬럼은 변경되지 않습니다.)").pack(side="left", padx=5)
        
        # 고유 텍스트 치환 프레임은 create_unique_replace_frame 메서드에서 생성
        
        # 언어 선택 프레임
        self.replace_lang_frame = tk.Frame(self.replace_input_frame)
        self.replace_lang_frame.pack(fill="x", padx=5, pady=5)
        
        tk.Label(self.replace_lang_frame, text="컬럼 선택:").pack(side="left")
        for lang in self.replace_languages:
            tk.Checkbutton(self.replace_lang_frame, text=lang, variable=self.replace_languages[lang]).pack(side="left", padx=5)
        
        # 전체 선택 체크박스
        self.select_frame = tk.Frame(self.replace_input_frame)
        self.select_frame.pack(fill="x", padx=5, pady=5)
        
        tk.Checkbutton(self.select_frame, text="전체 선택/해제", variable=self.select_all_var, 
                    command=self.toggle_all_rows).pack(side="left")
                    
        return self.replace_input_frame
    
    def create_unique_replace_frame(self):
        """고유 텍스트 치환 프레임 생성"""
        # 기존 프레임이 있으면 모든 자식 위젯 제거
        if hasattr(self, 'unique_replace_frame') and self.unique_replace_frame:
            for widget in self.unique_replace_frame.winfo_children():
                widget.destroy()
        else:
            self.unique_replace_frame = tk.Frame(self.replace_input_frame)
        
        # 설명 텍스트
        tk.Label(self.unique_replace_frame, 
                text="고유 텍스트 DB를 사용해 선택한 항목의 다국어를 치환합니다.",
                font=("Arial", 9)).pack(side="top", pady=(5, 2), anchor="w")
        
        # DB 정보 표시 영역
        info_frame = tk.Frame(self.unique_replace_frame, relief="groove", bd=1)
        info_frame.pack(fill="x", pady=5, padx=5)
        
        # 현재 검색어 정보
        search_frame = tk.Frame(info_frame)
        search_frame.pack(fill="x", padx=5, pady=3)
        
        tk.Label(search_frame, text="검색어 정보:", font=("Arial", 9, "bold"), width=12, anchor="w").pack(side="left")
        self.current_search_var = tk.StringVar(value="데이터 로드 중...")
        search_label = tk.Label(search_frame, textvariable=self.current_search_var, anchor="w")
        search_label.pack(side="left", fill="x", expand=True, padx=5)
        
        # DB 정보 표시
        db_frame = tk.Frame(info_frame)
        db_frame.pack(fill="x", padx=5, pady=3)
        
        tk.Label(db_frame, text="DB 정보:", font=("Arial", 9, "bold"), width=12, anchor="w").pack(side="left")
        self.db_info_var = tk.StringVar(value="DB 정보 로드 중...")
        db_label = tk.Label(db_frame, textvariable=self.db_info_var, anchor="w")
        db_label.pack(side="left", fill="x", expand=True, padx=5)
        
        # DB 정보 업데이트 버튼
        btn_frame = tk.Frame(self.unique_replace_frame)
        btn_frame.pack(fill="x", pady=5)
        
        refresh_btn = tk.Button(btn_frame, text="DB 정보 새로고침", command=self.update_unique_db_info)
        refresh_btn.pack(side="right", padx=5)
    
    def update_replace_ui(self):
        """치환 모드(단어/일괄/고유)에 따라 UI 업데이트"""
        # 먼저 모든 프레임 숨김
        self.word_replace_frame.pack_forget()
        self.bulk_replace_frame.pack_forget()
        if hasattr(self, 'unique_replace_frame') and self.unique_replace_frame:
            self.unique_replace_frame.pack_forget()
        
        # 현재 모드에 해당하는 프레임만 표시
        if self.replace_mode.get() == "replace":
            self.word_replace_frame.pack(fill="x", padx=5, pady=5)
        elif self.replace_mode.get() == "bulk":
            self.bulk_replace_frame.pack(fill="x", padx=5, pady=5)
        else:  # unique 모드
            if not hasattr(self, 'unique_replace_frame') or self.unique_replace_frame is None:
                self.create_unique_replace_frame()
            self.unique_replace_frame.pack(fill="x", padx=5, pady=5)
            
            # DB 정보 업데이트
            self.update_unique_db_info()
    
    def update_unique_db_info(self):
        """DB 정보 업데이트 및 현재 검색어에 대한 DB 값 표시"""
        if not hasattr(self, 'current_search_var') or not hasattr(self, 'db_info_var'):
            print("DEBUG: update_unique_db_info() - 변수가 초기화되지 않음")
            return
        
        # 현재 검색어 가져오기
        parent = self.parent
        keyword = parent.search_ui.search_text.get().strip()
        
        # DB 정보 및 검색어 정보 가져오기
        try:
            # DB 파일 경로
            db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unique_texts.db")
            
            if not os.path.exists(db_path):
                # 프로그램 실행 경로에서 찾아보기
                alternative_path = "unique_texts.db"
                if os.path.exists(alternative_path):
                    db_path = alternative_path
                else:
                    self.db_info_var.set("DB 파일을 찾을 수 없습니다.")
                    self.current_search_var.set("DB 파일 없음")
                    return
            
            # DB 연결
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 테이블 목록 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            
            if not tables:
                self.db_info_var.set("DB에 테이블이 없습니다.")
                self.current_search_var.set("DB 테이블 없음")
                conn.close()
                return
            
            # 기본 테이블 이름
            table_name = "unique_texts"
            
            # 테이블 이름이 다를 경우 첫 번째 테이블 사용
            if not any(table[0] == table_name for table in tables):
                table_name = tables[0][0]
            
            # 테이블 구조 확인
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns_info = cursor.fetchall()
            columns = [col[1] for col in columns_info]
            
            # 테이블 레코드 개수 확인
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            record_count = cursor.fetchone()[0]
            
            # DB 정보 표시
            language_cols = [col for col in columns if col not in ["id", "STRING_ID"]]
            
            # 언어 컬럼이 많은 경우 일부만 표시
            if len(language_cols) > 5:
                language_cols_str = ", ".join(language_cols[:5]) + f" 외 {len(language_cols) - 5}개"
            else:
                language_cols_str = ", ".join(language_cols)
            
            db_info = f"테이블: {table_name} | 레코드: {record_count:,}개 | 컬럼: {language_cols_str}"
            self.db_info_var.set(db_info)
            
            # 현재 검색어에 대한 DB 값 확인
            if keyword:
                # KR 컬럼 확인
                if "KR" in columns:
                    # 검색어와 일치하는 레코드 찾기
                    cursor.execute(f"SELECT * FROM {table_name} WHERE KR LIKE ?", (f"%{keyword}%",))
                    record = cursor.fetchone()
                    
                    if record:
                        # 컬럼 이름 가져오기
                        cursor.execute(f"PRAGMA table_info({table_name})")
                        col_names = [col[1] for col in cursor.fetchall()]
                        
                        # 결과를 딕셔너리로 변환
                        record_dict = {col_names[i]: record[i] for i in range(len(col_names))}
                        
                        # STRING_ID와 주요 언어 컬럼의 값을 표시
                        string_id = record_dict.get("STRING_ID", "")
                        
                        # 값 표시 (일부 컬럼만)
                        display_cols = ["KR", "EN", "CN", "TW", "TH", "PT", "ES", "DE"]
                        display_values = []
                        
                        for col in display_cols:
                            if col in record_dict and record_dict[col]:
                                shortened_value = record_dict[col]
                                if len(shortened_value) > 20:
                                    shortened_value = shortened_value[:18] + "..."
                                display_values.append(f"{col}: {shortened_value}")
                        
                        if display_values:
                            search_info = "ID: " + str(string_id) + " | " + " | ".join(display_values[:4])
                            if len(display_values) > 4:
                                search_info += f" 외 {len(display_values) - 4}개"
                            self.current_search_var.set(search_info)
                        else:
                            self.current_search_var.set(f"ID: {string_id} (값 없음)")
                    else:
                        self.current_search_var.set(f"'{keyword}'에 대한 값 없음")
                else:
                    self.current_search_var.set("KR 컬럼이 DB에 없음")
            else:
                self.current_search_var.set("검색어를 입력하세요")
            
            conn.close()
            
        except Exception as e:
            self.db_info_var.set(f"DB 정보 로딩 실패: {str(e)}")
            self.current_search_var.set("오류 발생")
            print(f"DEBUG: DB 정보 업데이트 중 오류: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def handle_click(self, event):
        """트리뷰 클릭 처리 - 치환 모드 체크박스"""
        parent = self.parent
        
        if parent.current_mode != "replace":
            return
            
        region = parent.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = parent.tree.identify_column(event.x)
            if column == "#1":  # 첫 번째 컬럼 (선택)
                item = parent.tree.identify_row(event.y)
                if item:
                    self.toggle_check(item)
                    return "break"  # 이벤트 처리 중단
    
    def toggle_check(self, item):
        """체크박스 상태 전환"""
        current = self.row_checks.get(item, True)  # 기본값은 True
        self.row_checks[item] = not current
        
        # 체크박스 상태 업데이트
        self.parent.tree.set(item, "선택", "✓" if not current else "")
        
        # 모든 항목이 체크되었는지 확인하여 전체 선택 체크박스 상태 업데이트
        all_checked = all(self.row_checks.get(item, True) for item in self.parent.tree.get_children())
        self.select_all_var.set(all_checked)
    
    def toggle_all_rows(self):
        """모든 행 체크박스 상태 전환"""
        checked = self.select_all_var.get()
        for item in self.parent.tree.get_children():
            self.row_checks[item] = checked
            self.parent.tree.set(item, "선택", "✓" if checked else "")
    
    def get_selected_replace_languages(self):
        """선택된 언어 목록을 반환합니다."""
        langs = []
        all_selected = self.replace_languages["ALL"].get()
        
        if all_selected:
            # ALL이 선택된 경우, 엑셀에 있는 모든 언어 컬럼을 사용하도록 함
            return ["ALL"]
        
        for lang, var in self.replace_languages.items():
            if lang != "ALL" and var.get():
                langs.append(lang)
        
        if not langs:
            # 선택된 언어가 없으면 기본적으로 모든 언어 사용
            return ["ALL"]
        
        return langs
    
    def show_unique_replacement_confirmation(self, replacement_data, langs):
        """고유 텍스트 치환 확인 팝업 표시"""
        print("DEBUG: 확인 팝업 함수 시작")
        parent = self.parent
        
        confirm_dialog = tk.Toplevel(parent.top)
        confirm_dialog.title("고유 텍스트 치환 확인")
        confirm_dialog.geometry("600x400")
        confirm_dialog.transient(parent.top)
        confirm_dialog.grab_set()  # 모달 다이얼로그로 설정
        
        # 항상 최상위에 표시
        confirm_dialog.attributes('-topmost', True)
        
        # 모달 다이얼로그 설정
        confirm_dialog.protocol("WM_DELETE_WINDOW", lambda: None)  # 닫기 버튼 비활성화
        
        # 스크롤 가능한 프레임 구성
        main_frame = tk.Frame(confirm_dialog)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 설명 레이블
        tk.Label(main_frame, text="고유 텍스트 대상의 변경할 컬럼 목록:", 
                font=("Arial", 10, "bold")).pack(pady=(0, 5))
        
        # 스크롤 영역
        canvas = tk.Canvas(main_frame)
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 각 KR 텍스트별 정보 표시
        row_count = 0
        for kr_text, data in replacement_data.items():
            # KR 텍스트 표시
            frame = tk.Frame(scrollable_frame)
            frame.pack(fill="x", pady=5)
            
            tk.Label(
                frame, text=f"\"{kr_text}\" 텍스트 다국어를 변경합니다:",
                anchor="w", justify="left", font=("Arial", 9, "bold")
            ).pack(fill="x")
            
            # 대상 컬럼 목록
            available_langs = list(data.keys())
            if available_langs:
                lang_text = f"대상 컬럼: {', '.join(available_langs)}"
                tk.Label(
                    frame, text=lang_text,
                    anchor="w", justify="left"
                ).pack(fill="x", padx=20)
            
            # 구분선
            separator = ttk.Separator(scrollable_frame, orient="horizontal")
            separator.pack(fill="x", pady=5)
            
            row_count += 1
        
        # 버튼 프레임
        button_frame = tk.Frame(confirm_dialog)
        button_frame.pack(pady=10)
        
        result = {"confirmed": False}  # 딕셔너리를 사용하여 참조로 값 변경
        
        def on_confirm():
            print("DEBUG: 실행하기 버튼 클릭")
            result["confirmed"] = True
            confirm_dialog.destroy()
        
        def on_cancel():
            print("DEBUG: 취소 버튼 클릭")
            confirm_dialog.destroy()
        
        confirm_button = tk.Button(button_frame, text="실행하기", command=on_confirm, width=10)
        confirm_button.pack(side="left", padx=10)
        cancel_button = tk.Button(button_frame, text="취소", command=on_cancel, width=10)
        cancel_button.pack(side="left", padx=10)
        
        print("DEBUG: 팝업 설정 완료, 대기 시작")
        
        # 다이얼로그 focus 설정
        confirm_dialog.focus_set()
        confirm_button.focus_set()
        
        # 중앙 정렬
        confirm_dialog.update_idletasks()
        width = confirm_dialog.winfo_width()
        height = confirm_dialog.winfo_height()
        x = (confirm_dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (confirm_dialog.winfo_screenheight() // 2) - (height // 2)
        confirm_dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # UI 갱신 강제 실행
        confirm_dialog.update()
        
        # 대화상자가 닫힐 때까지 대기
        confirm_dialog.wait_window()
        
        print(f"DEBUG: 팝업 종료, 결과: {result['confirmed']}")
        return result["confirmed"]
    
    def process_selected(self):
        """선택된 항목들을 파일+시트 단위로 그룹핑 후, 엑셀 수정하고 트리뷰 갱신"""
        import logging
        import os
        import time
        
        parent = self.parent
        
        # 로깅 설정
        log_dir = "excel_process_logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_file = os.path.join(log_dir, f"process_selected_{time.strftime('%Y%m%d_%H%M%S')}.log")
        logging.basicConfig(filename=log_file, level=logging.DEBUG, 
                            format='%(asctime)s - %(levelname)s - %(message)s')
        
        logging.info("===== 항목 처리 시작 =====")
        
        if parent.current_mode != "replace":
            logging.error("치환 모드가 아님, 함수 종료")
            if not parent.tree.get_children():
                parent.show_warning("모드 오류", "치환 모드에서만 사용 가능한 기능입니다.")
                parent.top.focus_force()
            return
                
        logging.info(f"모드: {self.replace_mode.get()}")

        # 선택 항목 확인 방법 변경 - 실제 값을 직접 수집
        selected_items_data = []
        for item in parent.tree.get_children():
            try:
                if self.row_checks.get(item, False):
                    values = parent.tree.item(item, "values")
                    tags = parent.tree.item(item, "tags")
                    selected_items_data.append((values, tags))
                    logging.debug(f"선택된 항목: {values}")
            except Exception as e:
                logging.error(f"항목 확인 중 오류: {str(e)}")
                continue

        logging.info(f"선택된 항목 수: {len(selected_items_data)}")

        if not selected_items_data:
            logging.warning("선택된 항목 없음, 함수 종료")
            parent.show_warning("선택 오류", "변경할 항목을 선택하세요.")
            return

        # 진행 창 먼저 시작
        parent.progress_manager.start_progress_window(0)
        
        # 열려 있는 엑셀 파일 확인
        parent.progress_manager.update_progress("열려 있는 엑셀 파일 확인 중...", success=True)
        open_files = parent.file_manager.check_open_excel_files()

        if open_files:
            # 사용자에게 열려 있는 파일 목록 표시 및 확인
            confirm_msg = f"{len(open_files)}개의 엑셀 파일이 현재 열려 있습니다.\n"
            confirm_msg += "해당 파일들은 처리에서 제외됩니다.\n\n"
            confirm_msg += "• " + "\n• ".join([os.path.basename(f) for f in open_files[:10]])
            
            if len(open_files) > 10:
                confirm_msg += f"\n\n... 외 {len(open_files) - 10}개 파일"
            
            confirm_msg += "\n\n계속 진행하시겠습니까?"
            
            proceed = parent.show_confirm("열린 파일 감지", confirm_msg, parent.progress_manager.progress_window)
            
            if not proceed:
                logging.info("사용자가 작업을 취소함")
                parent.progress_manager.update_progress("사용자가 작업을 취소했습니다.", success=False)
                parent.progress_manager.finish_progress()
                return

        # 고유 텍스트 치환 모드 처리
        replacement_data = None
        if self.replace_mode.get() == "unique":
            logging.info("고유 텍스트 치환 모드 처리 시작")
            
            # 선택된 KR 텍스트 수집
            kr_texts = []
            for values, _ in selected_items_data:
                if len(values) >= 4:
                    kr_index = 4 if parent.current_mode == "replace" else 3
                    if len(values) > kr_index:
                        kr_text = values[kr_index]
                        if kr_text:
                            kr_texts.append(kr_text)
            
            logging.info(f"수집된 KR 텍스트 수: {len(kr_texts)}")
            
            # DB에서 KR 텍스트로 다국어 검색
            replacement_data = parent.db_operations.query_unique_string_db(kr_texts)
            logging.info(f"치환 데이터 반환 결과: {replacement_data is not None}")
            
            # 치환 데이터가 없으면 종료
            if not replacement_data:
                logging.warning("치환 데이터 없음, 처리 종료")
                parent.progress_manager.update_progress("치환할 데이터가 없습니다.", success=False)
                parent.progress_manager.finish_progress()
                return
            
            # 확인 팝업 표시
            logging.info("확인 팝업 표시 시작")
            confirmed = self.show_unique_replacement_confirmation(
                    replacement_data, self.get_selected_replace_languages())
            
            if not confirmed:
                logging.info("사용자가 확인을 취소함")
                parent.progress_manager.update_progress("사용자가 작업을 취소했습니다.", success=False)
                parent.progress_manager.finish_progress()
                return
            
            logging.info("확인 완료, 계속 진행")
        
        # 항목 처리를 위한 변수 초기화
        file_task_map = {}
        parent.file_to_items = {}
        
        # 각 항목 처리
        for idx, (values, tags) in enumerate(selected_items_data):
            try:
                if len(values) < 4:
                    logging.warning(f"항목 {idx}의 값이 부족함: {values}")
                    parent.progress_manager.update_progress(f"항목 {idx+1} 정보 부족", success=False)
                    continue

                # 파일 경로는 태그에서 가져옴
                if tags and len(tags) > 0 and not any(tag in ["success", "error", "external_link"] for tag in tags):
                    file_path = tags[0]  # ✅ 태그에서 전체 파일 경로 가져오기
                else:
                    file_path = values[1]  # 태그가 없거나 특수 태그면 표시명 사용
                    
                display_name = values[1]  # 화면에 표시된 파일명
                sheet_name = values[2]
                string_id = values[3]
                
                # 데이터 열 위치 확인 (모드에 따라 위치 조정)
                kr_index = 4 if parent.current_mode == "replace" else 3
                kr_text = values[kr_index] if len(values) > kr_index else ""
                
                # 진행 창 업데이트
                parent.progress_manager.update_progress(f"항목 {idx+1}/{len(selected_items_data)} 처리: {display_name}", success=True)
                
                if not os.path.isabs(file_path):
                    file_path = os.path.join(parent.folder_path, file_path)
                    logging.info(f"  - 절대 경로로 변환: {file_path}")

                # 모드에 따른 작업 정보 생성
                if self.replace_mode.get() == "replace":
                    from_text = self.replace_from.get().strip()
                    to_text = self.replace_to.get().strip()
                    langs = self.get_selected_replace_languages()
                    task_info = ("replace", string_id, from_text, to_text, langs)
                    logging.info(f"  - 치환 태스크: 변경 전='{from_text}', 변경 후='{to_text}', 언어={langs}")
                    
                elif self.replace_mode.get() == "bulk":
                    new_value = self.bulk_value.get()
                    langs = self.get_selected_replace_languages()
                    task_info = ("bulk", string_id, None, new_value, langs)
                    logging.info(f"  - 일괄 변경 태스크: 새 값='{new_value}', 언어={langs}")
                    
                else:  # unique 모드
                    langs = self.get_selected_replace_languages()
                    # kr_text에 해당하는 다국어 매핑 정보 가져오기
                    if kr_text in replacement_data:
                        task_info = ("unique", string_id, kr_text, replacement_data[kr_text], langs)
                        logging.info(f"  - 고유 텍스트 치환 태스크: KR='{kr_text}', 언어={langs}")
                    else:
                        # 매칭 정보가 없으면 건너뛰기
                        logging.warning(f"  - KR='{kr_text}'에 대한 고유 텍스트 매핑 정보 없음, 건너뜀")
                        parent.progress_manager.update_progress(f"매칭 정보 없음: {kr_text}", success=False)
                        continue

                # 태스크 맵에 추가 (현재 아이템 정보 저장)
                item_info = (values, tags)
                
                # 파일 경로, 시트 이름, 태스크 정보 저장
                key = (file_path, sheet_name)
                if key not in file_task_map:
                    file_task_map[key] = []
                
                file_task_map[key].append(task_info)
                parent.file_to_items.setdefault(key, []).append(item_info)
            
            except Exception as e:
                logging.error(f"항목 {idx} 처리 중 오류: {str(e)}")
                import traceback
                logging.error(traceback.format_exc())
                parent.progress_manager.update_progress(f"항목 {idx+1} 처리 중 오류: {str(e)}", success=False)
                continue

        logging.info(f"생성된 파일-시트 그룹 수: {len(file_task_map)}")
        
        # 외부 링크 있는 파일 처리 결과 출력
        if parent.file_manager.files_with_external_links:
            external_links_count = len(parent.file_manager.files_with_external_links)
            parent.progress_manager.update_progress(f"외부 링크 파일: {external_links_count}개", success=False)
            parent.show_warning("외부 링크 발견", 
              f"{external_links_count}개 파일에서 외부 링크가 발견되어 처리되지 않습니다.")
        
        if not file_task_map:
            parent.progress_manager.update_progress("처리할 항목이 없습니다.", success=False)
            parent.show_info("처리 완료", "처리할 항목이 없습니다.")
            parent.progress_manager.finish_progress()
            return

        # 진행 창 업데이트 - 전체 작업 수
        parent.progress_manager.total_items = len(file_task_map)
        parent.progress_manager.processed_items = 0  # 진행률 초기화
        parent.progress_manager.update_progress(f"총 {parent.progress_manager.total_items}개 파일-시트 그룹 처리 시작", success=True)

        # 엑셀 작업 실행
        processed_files = self.execute_grouped_tasks(file_task_map)
        
        # DB 최신화 및 검색 결과 갱신
        parent.progress_manager.update_progress("String DB 최신화 중...", success=True)
        parent.db_operations.update_string_db_and_refresh_search()

        parent.progress_manager.update_progress("결과 업데이트 완료", success=True)
        parent.progress_manager.finish_progress(external_link_files=parent.file_manager.files_with_external_links)

        result_msg = f"✅ {self.replace_mode.get()} 모드 작업 완료"
        parent.status_var.set(result_msg)
        
        # 로그 파일 정보 UI에 표시
        parent.progress_manager.update_progress(f"\n📝 로그 파일 위치: {log_file}", success=True)
        
        logging.info("===== 항목 처리 종료 =====")
        
        # 검색 모드로 전환하여 UI 갱신
        parent.switch_to_search_mode(force=True)
        
        # 검색 실행하여 결과 갱신
        if parent.search_ui.search_text.get().strip():
            parent.search_ui.start_search()
        
        return processed_files
    
    def execute_grouped_tasks(self, file_task_map):
        """
        파일별로 그룹화된 작업 실행
        
        Args:
            file_task_map: 파일-시트별 작업 매핑
            
        Returns:
            처리된 파일 집합
        """
        parent = self.parent
        processed_files = set()
        
        # 파일별 캐시 정보 확인
        for idx, ((file_path, sheet_name), tasks) in enumerate(file_task_map.items()):
            # 작업 그룹 시작을 UI에 표시
            parent.progress_manager.update_progress(f"작업 그룹 {idx+1}/{len(file_task_map)} 시작: {os.path.basename(file_path)} - {sheet_name}", success=True)
            
            # 캐시 데이터 찾기
            matched_entry = None
            normalized_file_path = os.path.normpath(file_path)
            
            for cache_filename, cache_data in parent.cache.items():
                if os.path.normpath(cache_data.get("path", "")) == normalized_file_path:
                    matched_entry = cache_data
                    break
            
            if not matched_entry:
                parent.progress_manager.update_progress(f"캐시 매칭 실패: {os.path.basename(file_path)}", success=False)
                continue
            
            sheet_info = matched_entry["sheets"].get(sheet_name)
            if not sheet_info:
                parent.progress_manager.update_progress(f"시트 캐시 없음: {sheet_name}", success=False)
                continue
            
            header_row = sheet_info.get("header_row", 1)
            col_map = sheet_info.get("column_positions", {})
            
            # 엑셀 파일 수정 작업 실행
            success, _ = parent.file_manager.modify_excel_file(
                file_path, sheet_name, tasks, header_row, col_map)
            
            if success:
                processed_files.add(file_path)
                
            # 진행률 업데이트
            progress = min(100, int((idx + 1) / len(file_task_map) * 100))
            parent.progress_manager.progress_var.set(progress)
            
        # 완료된 파일 추가 저장 처리 (xlwings 사용)
        if processed_files:
            parent.progress_manager.update_progress("저장된 파일 확인 중...", success=True)
            parent.file_manager.ensure_files_saved(processed_files)
        
        # 처리 결과 요약
        processed_count = len(processed_files)
        open_count = len(parent.file_manager.open_excel_files)
        external_count = len(parent.file_manager.files_with_external_links)
        
        result_msg = f"✅ 작업 완료: {processed_count}개 파일 처리"
        if open_count > 0 or external_count > 0:
            result_msg += f" (제외: 열린 파일 {open_count}개, 외부 링크 {external_count}개)"
        
        parent.progress_manager.update_progress(result_msg, success=True)
        
        return processed_files