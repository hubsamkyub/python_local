import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from utils import ScrollableCheckList
from dialogs.selection_dialogs import LanguageSelectionDialog

class TabSetups:
    def __init__(self, manager):
        self.manager = manager
    
    def setup_translation_tab(self):
        """번역 대상 탭 설정 (동적 컬럼 생성)"""
        # 검색 및 필터 (기존과 동일)
        filter_frame = ttk.Frame(self.manager.translation_tab)
        filter_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(filter_frame, text="검색:").pack(side="left")
        self.manager.search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self.manager.search_var, width=30)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.manager.filter_translations())
        
        ttk.Label(filter_frame, text="상태 필터:").pack(side="left", padx=20)
        self.manager.filter_vars = {
            "신규": tk.BooleanVar(value=True), 
            "변경": tk.BooleanVar(value=True),
            "확인필요": tk.BooleanVar(value=True), 
            "확정": tk.BooleanVar(value=True),      # False → True로 변경
            "완료": tk.BooleanVar(value=True),
            "재번역완료": tk.BooleanVar(value=True)  # ← 새로 추가
        }
        for status, var in self.manager.filter_vars.items():
            ttk.Checkbutton(filter_frame, text=status, variable=var,
                        command=self.manager.filter_translations).pack(side="left", padx=5)
        
        # 번역 방법 필터 추가
        ttk.Label(filter_frame, text="방법 필터:").pack(side="left", padx=(20, 5))
        self.manager.method_filter_var = tk.StringVar(value="전체")
        self.manager.method_filter_combo = ttk.Combobox(filter_frame, textvariable=self.manager.method_filter_var, state="readonly", width=15)
        self.manager.method_filter_combo.pack(side="left")
        self.manager.method_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.manager.filter_translations())
        
        self.manager.stats_frame = ttk.Frame(filter_frame)
        self.manager.stats_frame.pack(side="right", padx=10)
        self.manager.update_stats_label()  
            
        # 전체 선택/해제 체크박스 추가
        select_all_frame = ttk.Frame(self.manager.translation_tab)
        select_all_frame.pack(fill="x", padx=5, pady=(5,0))
        self.manager.select_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(select_all_frame, text="전체 선택/해제", variable=self.manager.select_all_var, command=self.manager.toggle_all_selections).pack(side="left")
        
        # 동적 컬럼 리스트 생성
        base_columns = ["선택", "STRING_ID", "KR", "상태"]
        lang_columns = self.manager.VISIBLE_LANGS  # __init__에서 정의한 언어 순서 사용
        end_columns = ["번역방법"]
        columns = base_columns + lang_columns + end_columns

        tree_frame = ttk.Frame(self.manager.translation_tab)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.manager.translation_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings")
        
        # 컬럼 설정
        self.manager.translation_tree.column("#0", width=0, stretch=False)
        self.manager.translation_tree.heading("#0", text="")

        # 고정 컬럼 설정
        self.manager.translation_tree.column("선택", width=40, anchor="center")
        self.manager.translation_tree.column("STRING_ID", width=150)
        self.manager.translation_tree.column("KR", width=250)
        self.manager.translation_tree.column("상태", width=80)
        self.manager.translation_tree.column("번역방법", width=100)

        # 동적 언어 컬럼 설정
        for lang_col in lang_columns:
            self.manager.translation_tree.column(lang_col, width=150)

        for col in columns:
            self.manager.translation_tree.heading(col, text=col)
        
        # 스크롤바 (기존과 동일)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.manager.translation_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.manager.translation_tree.xview)
        self.manager.translation_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.manager.translation_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # 우클릭 컨텍스트 메뉴 생성
        self.manager.context_menu = tk.Menu(self.manager.root, tearoff=0)
        self.manager.context_menu.add_command(label="🔄 이 항목 재번역 (API 강제)", command=self.manager.force_retranslate_selected)
        self.manager.context_menu.add_command(label="✏️ 직접 편집", command=self.manager.edit_translation_inline)
        self.manager.context_menu.add_separator()
        self.manager.context_menu.add_command(label="🗑️ TM에서 삭제", command=self.manager.remove_from_tm)
        self.manager.context_menu.add_command(label="📋 TM 항목 보기", command=self.manager.view_tm_entry)
        
        # 트리뷰에 우클릭 이벤트 바인딩
        self.manager.translation_tree.bind("<Button-3>", self.manager.show_context_menu)  # 우클릭
        self.manager.translation_tree.bind("<Control-Button-1>", self.manager.show_context_menu)  # Ctrl+클릭 (Mac 호환)
        self.manager.translation_tree.bind("<Button-1>", self.manager.on_tree_click)  # 기존 클릭 이벤트
                
        # 키보드 이벤트 바인딩
        self.manager.translation_tree.bind("<KeyPress-space>", self.manager.toggle_selected_checkboxes)
        self.manager.translation_tree.bind("<Return>", self.manager.toggle_selected_checkboxes)  # Enter키도 동일하게
        self.manager.translation_tree.bind("<Control-a>", self.manager.select_all_items)  # Ctrl+A로 전체 선택
        self.manager.translation_tree.bind("<Control-d>", self.manager.deselect_all_items)  # Ctrl+D로 전체 해제
        self.manager.translation_tree.bind("<Control-i>", self.manager.invert_selection)  # Ctrl+I: 선택 반전
        self.manager.translation_tree.bind("<Delete>", self.manager.clear_selected_translations)  # Delete: 선택된 항목의 번역 내용 삭제
        self.manager.translation_tree.bind("<F2>", self.manager.edit_selected_item)  # F2: 선택된 항목 편집

        # 포커스 설정으로 키보드 이벤트가 작동하도록
        self.manager.translation_tree.focus_set()
        
        # 키보드 단축키 안내 추가
        help_frame = ttk.Frame(self.manager.translation_tab)
        help_frame.pack(fill="x", padx=5, pady=2)
        
        help_text = "💡 키보드 단축키: Space/Enter=체크토글 | Ctrl+A=전체선택 | Ctrl+D=전체해제 | Ctrl+I=선택반전 | F2=편집 | Del=번역삭제"
        help_label = ttk.Label(help_frame, text=help_text, font=("맑은 고딕", 8), foreground="gray")
        help_label.pack(side="right")        
        
        self.manager.check_states = {}

    def setup_scenario_tab(self):
        """시나리오 번역 전용 탭 설정 (화자 로드 기능 추가)"""
        main_frame = ttk.Frame(self.manager.scenario_tab, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 시나리오 매니저 초기화 (탭 열 때마다 확인)
        self.manager.ensure_scenario_manager()
        
        # 시나리오 번역 안내
        info_frame = ttk.LabelFrame(main_frame, text="ℹ️ 시나리오 번역이란?")
        info_frame.pack(fill="x", pady=5)
        
        info_text = """화자별 맞춤 번역: 각 캐릭터의 말투와 성격에 맞는 번역을 제공합니다.
    사용법: 1) 레퍼런스 데이터 설정 → 2) 화자 분석 → 3) 메인 설정에서 '시나리오모드' 체크 → 4) 번역 실행"""
        
        ttk.Label(info_frame, text=info_text, font=("맑은 고딕", 9), foreground="navy").pack(padx=10, pady=5)
        
        # === 1. 레퍼런스 데이터 설정 ===
        ref_frame = ttk.LabelFrame(main_frame, text="1️⃣ 레퍼런스 데이터 설정")
        ref_frame.pack(fill="x", padx=5, pady=5)
        
        ref_inner = ttk.Frame(ref_frame, padding="5")
        ref_inner.pack(fill="x")
        
        # 파일 선택
        ttk.Label(ref_inner, text="레퍼런스 파일:").grid(row=0, column=0, sticky="w")
        self.manager.ref_file_var = tk.StringVar()
        ttk.Entry(ref_inner, textvariable=self.manager.ref_file_var, width=40).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(ref_inner, text="엑셀 선택", command=self.manager.select_reference_file).grid(row=0, column=2, padx=2)
        ttk.Button(ref_inner, text="구글 시트", command=self.manager.load_reference_from_gsheet).grid(row=0, column=3, padx=2)
        
        ref_inner.grid_columnconfigure(1, weight=1)
        
        # 분석 버튼들
        analysis_frame = ttk.Frame(ref_inner)
        analysis_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=5)
        
        ttk.Button(analysis_frame, text="🔍 구조확인", command=self.manager.debug_file_structure_detailed).pack(side="left", padx=2)
        ttk.Button(analysis_frame, text="🚀 자동분석", command=self.manager.analyze_reference_data_smart).pack(side="left", padx=2)
        ttk.Button(analysis_frame, text="🔧 수동매핑", command=self.manager.manual_column_mapping_dialog).pack(side="left", padx=2)
        ttk.Button(analysis_frame, text="📜 디버그로그", command=self.manager.show_latest_debug_log).pack(side="left", padx=2)
        ttk.Button(analysis_frame, text="💾 데이터셋관리", command=self.manager.show_reference_dataset_manager).pack(side="left", padx=2)
        
        # === 2. 화자 관리 ===
        speaker_frame = ttk.LabelFrame(main_frame, text="2️⃣ 화자 스타일 관리")
        speaker_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        speaker_inner = ttk.Frame(speaker_frame, padding="5")
        speaker_inner.pack(fill="both", expand=True)
        
        # 화자 상태 표시 및 새로고침
        status_frame = ttk.Frame(speaker_inner)
        status_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        
        self.manager.speaker_status_label = ttk.Label(status_frame, text="화자 정보 로딩 중...", foreground="blue")
        self.manager.speaker_status_label.pack(side="left")
        
        ttk.Button(status_frame, text="🔄 새로고침", command=self.manager.refresh_speaker_list).pack(side="right")
        
        # 화자 리스트
        columns = ("화자", "성별", "말투", "번역 스타일", "레퍼런스 수")
        self.manager.speaker_tree = ttk.Treeview(speaker_inner, columns=columns, show="headings", height=8)
        
        # 컬럼 너비 조정
        widths = {"화자": 80, "성별": 50, "말투": 60, "번역 스타일": 200, "레퍼런스 수": 80}
        for col in columns:
            self.manager.speaker_tree.heading(col, text=col)
            self.manager.speaker_tree.column(col, width=widths.get(col, 100))
        
        speaker_scroll = ttk.Scrollbar(speaker_inner, orient="vertical", command=self.manager.speaker_tree.yview)
        self.manager.speaker_tree.configure(yscrollcommand=speaker_scroll.set)
        
        self.manager.speaker_tree.grid(row=1, column=0, sticky="nsew")
        speaker_scroll.grid(row=1, column=1, sticky="ns")
        
        speaker_inner.grid_rowconfigure(1, weight=1)
        speaker_inner.grid_columnconfigure(0, weight=1)
        
        # 화자 관리 버튼들
        speaker_btn_frame = ttk.Frame(speaker_inner)
        speaker_btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        
        ttk.Button(speaker_btn_frame, text="➕ 추가", command=self.manager.add_speaker).pack(side="left", padx=2)
        ttk.Button(speaker_btn_frame, text="✏️ 편집", command=self.manager.edit_speaker).pack(side="left", padx=2)
        ttk.Button(speaker_btn_frame, text="🗑️ 삭제", command=self.manager.delete_speaker).pack(side="left", padx=2)
        
        # === 3. 사용 안내 ===
        usage_frame = ttk.LabelFrame(main_frame, text="3️⃣ 사용 방법")
        usage_frame.pack(fill="x", padx=5, pady=5)
        
        usage_text = """✅ 설정 완료 후: 메인 설정에서 '시나리오모드' 체크 → LLM 엔진 자동 선택 → 번역 실행
    💡 팁: 레퍼런스 데이터는 한 번 분석하면 자동 저장되어 다음에 재사용 가능합니다."""
        
        ttk.Label(usage_frame, text=usage_text, font=("맑은 고딕", 9), 
                foreground="darkgreen").pack(padx=10, pady=5)
        
        # 초기 화자 리스트 로드
        self.manager.refresh_speaker_list()

    def setup_tm_management_tab(self):
        """번역 메모리(TM) 관리 탭 UI 구성 (하위 탭 구조 적용)"""
        # TM 관리 탭의 메인 프레임
        main_tm_frame = ttk.Frame(self.manager.tm_management_tab, padding="5")
        main_tm_frame.pack(fill="both", expand=True)
        
        # TM 관리 탭 내에 노트북(하위 탭) 생성
        tm_notebook = ttk.Notebook(main_tm_frame)
        tm_notebook.pack(fill="both", expand=True)

        # 탭 1: TM 조회 및 직접 편집
        view_edit_tab = ttk.Frame(tm_notebook)
        tm_notebook.add(view_edit_tab, text="TM 조회/편집")
        self.setup_tm_view_edit_tab(view_edit_tab) # UI 구성 함수 호출

        # 탭 2: Excel로 가져오기/업데이트
        import_tab = ttk.Frame(tm_notebook)
        tm_notebook.add(import_tab, text="Excel로 가져오기/업데이트")
        self.setup_excel_import_tab(import_tab)

    def setup_tm_view_edit_tab(self, parent_tab):
        """'TM 조회/편집' 하위 탭의 UI를 구성합니다."""
        main_frame = ttk.Frame(parent_tab, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        control_frame = ttk.LabelFrame(main_frame, text="TM 관리 도구")
        control_frame.pack(fill="x", pady=5)
        
        ttk.Button(control_frame, text="🔍 TM 상태 확인", command=self.manager.debug_tm_status).pack(side="left", padx=10)

        build_frame = ttk.Frame(control_frame)
        build_frame.pack(side="left", padx=10, pady=5)

        ttk.Button(build_frame, text="소스 DB 폴더로 TM 구축", command=self.manager.start_db_build).pack(side="left")

        mode_frame = ttk.Frame(build_frame)
        mode_frame.pack(side="left", padx=10)

        ttk.Radiobutton(mode_frame, text="충돌 우선 해결 모드 (엄격)", variable=self.manager.db_build_mode_var, value="conflict").pack(anchor="w")
        ttk.Radiobutton(mode_frame, text="빈칸 채우기 모드 (보강)", variable=self.manager.db_build_mode_var, value="fill_blanks").pack(anchor="w")
        ttk.Button(build_frame, text="TM 정리하기 (제외 규칙 적용)", command=self.manager.cleanup_tm_with_rules).pack(side="left", padx=20)
        
        search_frame = ttk.Frame(control_frame)
        search_frame.pack(side="right", padx=10)
        ttk.Label(search_frame, text="KR 검색:").pack(side="left")
        self.manager.tm_view_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.manager.tm_view_search_var, width=40)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.manager.load_tm_view())
        
        list_frame = ttk.LabelFrame(main_frame, text="마스터 번역 메모리 내용")
        list_frame.pack(fill="both", expand=True, pady=10)

        base_columns = ["KR"] + self.manager.VISIBLE_LANGS
        self.manager.tm_view_tree = ttk.Treeview(list_frame, columns=base_columns, show="headings")
        for col in base_columns:
            self.manager.tm_view_tree.heading(col, text=col)
            self.manager.tm_view_tree.column(col, width=150)
        self.manager.tm_view_tree.column("KR", width=250)

        v_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.manager.tm_view_tree.yview)
        h_scroll = ttk.Scrollbar(list_frame, orient="horizontal", command=self.manager.tm_view_tree.xview)
        self.manager.tm_view_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        self.manager.tm_view_tree.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")

    def setup_excel_import_tab(self, parent_tab):
        """'Excel로 가져오기/업데이트' 하위 탭의 UI를 구성합니다."""
        main_frame = ttk.Frame(parent_tab, padding="10")
        main_frame.pack(fill="both", expand=True)

        # 1. 폴더 선택 및 파일 검색 프레임
        folder_frame = ttk.LabelFrame(main_frame, text="1. 파일 검색")
        folder_frame.pack(fill="x", pady=5)
        ttk.Label(folder_frame, text="엑셀 폴더:").pack(side="left", padx=5)
        ttk.Entry(folder_frame, textvariable=self.manager.excel_import_folder_var, width=80).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(folder_frame, text="폴더 선택", command=lambda: self.manager.excel_import_folder_var.set(filedialog.askdirectory() or "")).pack(side="left", padx=5)
        ttk.Button(folder_frame, text="파일 검색 실행", command=self.manager.search_excel_for_import).pack(side="left", padx=5)

        # 2. 파일 및 언어 선택 프레임
        selection_frame = ttk.LabelFrame(main_frame, text="2. 대상 선택")
        selection_frame.pack(fill="both", expand=True, pady=5)
        
        # 2-1. 파일 선택 리스트
        file_list_frame = ttk.Frame(selection_frame)
        file_list_frame.pack(side="left", fill="both", expand=True, padx=5)
        ttk.Label(file_list_frame, text="처리할 파일 선택:").pack(anchor="w")
        self.manager.excel_files_checklist = ScrollableCheckList(file_list_frame, height=15)
        self.manager.excel_files_checklist.pack(fill="both", expand=True)

        # 2-2. 언어 선택 리스트
        lang_list_frame = ttk.Frame(selection_frame)
        lang_list_frame.pack(side="left", fill="y", padx=5)
        ttk.Label(lang_list_frame, text="업데이트할 언어 선택:").pack(anchor="w")
        for lang in self.manager.VISIBLE_LANGS:
            var = tk.BooleanVar(value=True)
            self.manager.excel_import_lang_vars[lang] = var
            ttk.Checkbutton(lang_list_frame, text=lang, variable=var).pack(anchor="w")

        # 3. 실행 버튼 프레임
        action_frame = ttk.LabelFrame(main_frame, text="3. 실행")
        action_frame.pack(fill="x", pady=5)
        ttk.Button(action_frame, text="선택한 파일로 TM 업데이트 시작", command=self.manager.start_excel_import, style="Accent.TButton").pack(pady=10)

    def setup_conflict_tab(self):
        """'충돌 해결' 탭의 UI를 구성합니다."""
        main_frame = ttk.Frame(self.manager.conflict_tab, padding="10")
        main_frame.pack(fill="both", expand=True)

        # 1. 상단 액션 프레임
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", pady=5)
        ttk.Button(action_frame, text="🚀 가장 많이 쓰인 번역으로 전체 자동 해결", command=self.manager.auto_resolve_all_conflicts).pack(side="left")
        ttk.Button(action_frame, text="🔄 목록 새로고침", command=self.manager.load_conflicts_to_view).pack(side="left", padx=10)

        # 2. 콘텐츠 프레임 (충돌 목록 + 해결 패널)
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True, pady=5)

        # 2-1. 충돌 목록 Treeview
        list_frame = ttk.LabelFrame(content_frame, text="충돌 항목 목록")
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        columns = ["KR"] + self.manager.VISIBLE_LANGS
        self.manager.conflict_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.manager.conflict_tree.heading("KR", text="KR (충돌 항목)")
        self.manager.conflict_tree.column("KR", width=200)
        for lang in self.manager.VISIBLE_LANGS:
            self.manager.conflict_tree.heading(lang, text=lang)
            self.manager.conflict_tree.column(lang, width=120)
        
        self.manager.conflict_tree.tag_configure('conflict', foreground='red')
        self.manager.conflict_tree.bind("<<TreeviewSelect>>", self.manager.on_conflict_row_selected)
        self.manager.conflict_tree.pack(fill="both", expand=True)

        # 2-2. 충돌 해결 패널
        resolve_panel = ttk.LabelFrame(content_frame, text="충돌 해결")
        resolve_panel.pack(side="right", fill="y")

        self.manager.conflict_kr_var = tk.StringVar()
        ttk.Label(resolve_panel, text="KR:").pack(anchor="w", padx=5, pady=5)
        ttk.Entry(resolve_panel, textvariable=self.manager.conflict_kr_var, state="readonly", width=40).pack(anchor="w", padx=5)
        
        self.manager.conflict_combos = {}
        for lang in self.manager.VISIBLE_LANGS:
            ttk.Label(resolve_panel, text=f"{lang} 후보:").pack(anchor="w", padx=5, pady=(10, 0))
            combo = ttk.Combobox(resolve_panel, state="disabled", width=38)
            combo.pack(anchor="w", padx=5)
            self.manager.conflict_combos[lang] = combo
        
        ttk.Button(resolve_panel, text="✅ 선택된 값으로 충돌 해결", command=self.manager.resolve_selected_conflict).pack(pady=20, padx=5)

    def setup_glossary_tab(self):
        """용어집 관리 탭 UI 구성 (STRING_ID 제거)"""
        main_frame = ttk.Frame(self.manager.glossary_tab, padding="10")
        main_frame.pack(fill="both", expand=True)

        # 상단 컨트롤 프레임
        control_frame = ttk.LabelFrame(main_frame, text="용어집 관리 도구")
        control_frame.pack(fill="x", pady=5)
        
        # 좌측: 핵심 기능
        left_control = ttk.Frame(control_frame)
        left_control.pack(side="left", padx=5, pady=5)
        
        ttk.Button(left_control, text="🔄 구글 시트와 동기화", command=self.manager.sync_glossary_from_gsheet, style="Accent.TButton").pack(side="left")
        
        # 우측: 검색 기능
        search_frame = ttk.Frame(control_frame)
        search_frame.pack(side="right", padx=10, pady=5)
        
        ttk.Label(search_frame, text="KR 검색:").pack(side="left")
        self.manager.glossary_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.manager.glossary_search_var, width=30)
        search_entry.pack(side="left", padx=5)
        ttk.Button(search_frame, text="🔍 필터", command=self.manager.filter_glossary).pack(side="left", padx=2)
        ttk.Button(search_frame, text="🔄 전체", command=self.manager.clear_glossary_filter).pack(side="left", padx=2)
        
        # 사용자 안내
        info_label = ttk.Label(control_frame, text="ℹ️ 용어 추가/수정/삭제는 마스터 구글 시트에서 진행 후, 동기화 버튼을 눌러주세요.", foreground="blue")
        info_label.pack(side="bottom", padx=10, pady=2)

        # 용어집 목록 Treeview (STRING_ID 제거)
        list_frame = ttk.LabelFrame(main_frame, text="용어집 목록 (마스터 구글 시트의 로컬 사본)")
        list_frame.pack(fill="both", expand=True, pady=10)

        # STRING_ID 제거된 컬럼 목록
        self.manager.glossary_cols = ["kr", "en", "cn", "tw", "contributor", "verified"]
        self.manager.glossary_tree = ttk.Treeview(list_frame, columns=self.manager.glossary_cols, show="headings")
        
        # 컬럼 너비 조정
        col_widths = {"kr": 200, "en": 200, "cn": 120, "tw": 120, "contributor": 100, "verified": 70}
        for col in self.manager.glossary_cols:
            self.manager.glossary_tree.heading(col, text=col.upper())
            self.manager.glossary_tree.column(col, width=col_widths.get(col, 100), anchor="w")

        # 스크롤바
        v_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.manager.glossary_tree.yview)
        h_scroll = ttk.Scrollbar(list_frame, orient="horizontal", command=self.manager.glossary_tree.xview)
        self.manager.glossary_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.manager.glossary_tree.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")

    def setup_exclusion_tab(self):
        """규칙 기반 제외 목록 관리 탭 UI 구성"""
        main_frame = ttk.Frame(self.manager.exclusion_tab, padding="10")
        main_frame.pack(fill="both", expand=True)

        # 1. 새 규칙 추가 프레임
        add_frame = ttk.LabelFrame(main_frame, text="새 규칙 추가")
        add_frame.pack(fill="x", pady=5)
        
        ttk.Label(add_frame, text="규칙 유형:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.manager.rule_type_var = tk.StringVar(value="startswith")
        rule_types = ["startswith", "endswith", "contains", "equals", "length", "regex"]
        ttk.Combobox(add_frame, textvariable=self.manager.rule_type_var, values=rule_types, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(add_frame, text="적용 필드:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.manager.rule_field_var = tk.StringVar(value="KR")
        # 나중에 다른 언어 필드도 추가할 수 있도록 확장 가능하게 구성
        rule_fields = ["KR", "STRING_ID"] + self.manager.VISIBLE_LANGS
        ttk.Combobox(add_frame, textvariable=self.manager.rule_field_var, values=rule_fields, state="readonly").grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        
        ttk.Label(add_frame, text="값:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.manager.rule_value_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.manager.rule_value_var).grid(row=0, column=5, padx=5, pady=5, sticky="ew")

        ttk.Label(add_frame, text="설명:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.manager.rule_desc_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self.manager.rule_desc_var).grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="ew")
        
        ttk.Button(add_frame, text="규칙 추가", command=self.manager.add_exclusion_rule).grid(row=0, column=6, rowspan=2, padx=10, pady=5, ipady=10)
        
        add_frame.grid_columnconfigure(5, weight=1)

        # 2. 규칙 목록 표시 프레임
        list_frame = ttk.LabelFrame(main_frame, text="제외 규칙 목록")
        list_frame.pack(fill="both", expand=True, pady=10)

        columns = ("description", "rule_type", "field", "value", "enabled")
        self.manager.exclusion_rule_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.manager.exclusion_rule_tree.heading("description", text="설명")
        self.manager.exclusion_rule_tree.heading("rule_type", text="규칙 유형")
        self.manager.exclusion_rule_tree.heading("field", text="적용 필드")
        self.manager.exclusion_rule_tree.heading("value", text="값")
        self.manager.exclusion_rule_tree.heading("enabled", text="활성화")
        
        self.manager.exclusion_rule_tree.column("description", width=250)
        self.manager.exclusion_rule_tree.column("rule_type", width=100)
        self.manager.exclusion_rule_tree.column("field", width=100)
        self.manager.exclusion_rule_tree.column("value", width=150)
        self.manager.exclusion_rule_tree.column("enabled", width=80, anchor="center")
        
        self.manager.exclusion_rule_tree.pack(side="left", fill="both", expand=True)
        
        # 3. 관리 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=5)
        
        ttk.Button(button_frame, text="활성화/비활성화", command=self.manager.toggle_exclusion_rule).pack(side="left")
        ttk.Button(button_frame, text="규칙 삭제", command=self.manager.delete_exclusion_rule).pack(side="left", padx=10)
        ttk.Button(button_frame, text="기본값으로 초기화", command=self.manager.reset_default_rules).pack(side="right")

    def setup_history_tab(self):
        """번역 이력 탭 설정"""
        # 검색
        search_frame = ttk.Frame(self.manager.history_tab)
        search_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(search_frame, text="검색:").pack(side="left")
        self.manager.history_search_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.manager.history_search_var, width=40).pack(side="left", padx=5)
        ttk.Button(search_frame, text="검색", command=self.manager.search_history).pack(side="left")
        
        # 이력 테이블
        columns = ["시간", "STRING_ID", "KR", "번역방법", "상태"]
        self.manager.history_tree = ttk.Treeview(self.manager.history_tab, columns=columns, show="headings")
        
        for col in columns:
            self.manager.history_tree.heading(col, text=col)
        
        self.manager.history_tree.column("시간", width=150)
        self.manager.history_tree.column("STRING_ID", width=150)
        self.manager.history_tree.column("KR", width=300)
        self.manager.history_tree.column("번역방법", width=100)
        self.manager.history_tree.column("상태", width=100)
        
        history_scroll = ttk.Scrollbar(self.manager.history_tab, orient="vertical", command=self.manager.history_tree.yview)
        self.manager.history_tree.configure(yscrollcommand=history_scroll.set)
        
        self.manager.history_tree.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        history_scroll.pack(side="right", fill="y")