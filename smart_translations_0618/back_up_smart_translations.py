#백업
#tab_ui로 분리
    # def setup_excel_import_tab(self, parent_tab):
    #     """'Excel로 가져오기/업데이트' 하위 탭의 UI를 구성합니다."""
    #     main_frame = ttk.Frame(parent_tab, padding="10")
    #     main_frame.pack(fill="both", expand=True)

    #     # 1. 폴더 선택 및 파일 검색 프레임
    #     folder_frame = ttk.LabelFrame(main_frame, text="1. 파일 검색")
    #     folder_frame.pack(fill="x", pady=5)
    #     ttk.Label(folder_frame, text="엑셀 폴더:").pack(side="left", padx=5)
    #     ttk.Entry(folder_frame, textvariable=self.excel_import_folder_var, width=80).pack(side="left", fill="x", expand=True, padx=5)
    #     ttk.Button(folder_frame, text="폴더 선택", command=lambda: self.excel_import_folder_var.set(filedialog.askdirectory() or "")).pack(side="left", padx=5)
    #     ttk.Button(folder_frame, text="파일 검색 실행", command=self.search_excel_for_import).pack(side="left", padx=5)

    #     # 2. 파일 및 언어 선택 프레임
    #     selection_frame = ttk.LabelFrame(main_frame, text="2. 대상 선택")
    #     selection_frame.pack(fill="both", expand=True, pady=5)
        
    #     # 2-1. 파일 선택 리스트
    #     file_list_frame = ttk.Frame(selection_frame)
    #     file_list_frame.pack(side="left", fill="both", expand=True, padx=5)
    #     ttk.Label(file_list_frame, text="처리할 파일 선택:").pack(anchor="w")
    #     self.excel_files_checklist = ScrollableCheckList(file_list_frame, height=15)
    #     self.excel_files_checklist.pack(fill="both", expand=True)

    #     # 2-2. 언어 선택 리스트
    #     lang_list_frame = ttk.Frame(selection_frame)
    #     lang_list_frame.pack(side="left", fill="y", padx=5)
    #     ttk.Label(lang_list_frame, text="업데이트할 언어 선택:").pack(anchor="w")
    #     for lang in self.VISIBLE_LANGS:
    #         var = tk.BooleanVar(value=True)
    #         self.excel_import_lang_vars[lang] = var
    #         ttk.Checkbutton(lang_list_frame, text=lang, variable=var).pack(anchor="w")

    #     # 3. 실행 버튼 프레임
    #     action_frame = ttk.LabelFrame(main_frame, text="3. 실행")
    #     action_frame.pack(fill="x", pady=5)
    #     ttk.Button(action_frame, text="선택한 파일로 TM 업데이트 시작", command=self.start_excel_import, style="Accent.TButton").pack(pady=10)
    
    # def setup_translation_tab(self):
    #     """번역 대상 탭 설정 (동적 컬럼 생성)"""
    #     # 검색 및 필터 (기존과 동일)
    #     filter_frame = ttk.Frame(self.translation_tab)
    #     filter_frame.pack(fill="x", padx=5, pady=5)
        
    #     ttk.Label(filter_frame, text="검색:").pack(side="left")
    #     self.search_var = tk.StringVar()
    #     search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=30)
    #     search_entry.pack(side="left", padx=5)
    #     search_entry.bind("<KeyRelease>", lambda e: self.filter_translations())
        
    #     ttk.Label(filter_frame, text="상태 필터:").pack(side="left", padx=20)
    #     self.filter_vars = {
    #         "신규": tk.BooleanVar(value=True), 
    #         "변경": tk.BooleanVar(value=True),
    #         "확인필요": tk.BooleanVar(value=True), 
    #         "확정": tk.BooleanVar(value=True),      # False → True로 변경
    #         "완료": tk.BooleanVar(value=True),
    #         "재번역완료": tk.BooleanVar(value=True)  # ← 새로 추가
    #     }
    #     for status, var in self.filter_vars.items():
    #         ttk.Checkbutton(filter_frame, text=status, variable=var,
    #                     command=self.filter_translations).pack(side="left", padx=5)
        
    #     # <<< 시작: 번역 방법 필터 추가 >>>
    #     ttk.Label(filter_frame, text="방법 필터:").pack(side="left", padx=(20, 5))
    #     self.method_filter_var = tk.StringVar(value="전체")
    #     self.method_filter_combo = ttk.Combobox(filter_frame, textvariable=self.method_filter_var, state="readonly", width=15)
    #     self.method_filter_combo.pack(side="left")
    #     self.method_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_translations())
    #     # <<< 종료: 번역 방법 필터 추가 >>>
        
    #     self.stats_frame = ttk.Frame(filter_frame)
    #     self.stats_frame.pack(side="right", padx=10)
    #     self.update_stats_label()  
            
    #     # <<< 시작: 전체 선택/해제 체크박스 추가 >>>
    #     select_all_frame = ttk.Frame(self.translation_tab)
    #     select_all_frame.pack(fill="x", padx=5, pady=(5,0))
    #     self.select_all_var = tk.BooleanVar(value=True)
    #     ttk.Checkbutton(select_all_frame, text="전체 선택/해제", variable=self.select_all_var, command=self.toggle_all_selections).pack(side="left")
    #     # <<< 종료: 전체 선택/해제 체크박스 추가 >>>
        
    #     # <<< 시작: 동적 컬럼 리스트 생성 >>>
    #     base_columns = ["선택", "STRING_ID", "KR", "상태"]
    #     lang_columns = self.VISIBLE_LANGS  # __init__에서 정의한 언어 순서 사용
    #     end_columns = ["번역방법"]
    #     columns = base_columns + lang_columns + end_columns
    #     # <<< 종료: 동적 컬럼 리스트 생성 >>>

    #     tree_frame = ttk.Frame(self.translation_tab)
    #     tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
    #     self.translation_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings")
        
    #     # 컬럼 설정
    #     self.translation_tree.column("#0", width=0, stretch=False)
    #     self.translation_tree.heading("#0", text="")

    #     # 고정 컬럼 설정
    #     self.translation_tree.column("선택", width=40, anchor="center")
    #     self.translation_tree.column("STRING_ID", width=150)
    #     self.translation_tree.column("KR", width=250)
    #     self.translation_tree.column("상태", width=80)
    #     self.translation_tree.column("번역방법", width=100)

    #     # <<< 시작: 동적 언어 컬럼 설정 >>>
    #     for lang_col in lang_columns:
    #         self.translation_tree.column(lang_col, width=150)
    #     # <<< 종료: 동적 언어 컬럼 설정 >>>

    #     for col in columns:
    #         self.translation_tree.heading(col, text=col)
        
    #     # 스크롤바 (기존과 동일)
    #     v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.translation_tree.yview)
    #     h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.translation_tree.xview)
    #     self.translation_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
    #     self.translation_tree.grid(row=0, column=0, sticky="nsew")
    #     v_scroll.grid(row=0, column=1, sticky="ns")
    #     h_scroll.grid(row=1, column=0, sticky="ew")
        
    #     tree_frame.grid_rowconfigure(0, weight=1)
    #     tree_frame.grid_columnconfigure(0, weight=1)
        
    #     # 우클릭 컨텍스트 메뉴 생성
    #     self.context_menu = tk.Menu(self.root, tearoff=0)
    #     self.context_menu.add_command(label="🔄 이 항목 재번역 (API 강제)", command=self.force_retranslate_selected)
    #     self.context_menu.add_command(label="✏️ 직접 편집", command=self.edit_translation_inline)
    #     self.context_menu.add_separator()
    #     self.context_menu.add_command(label="🗑️ TM에서 삭제", command=self.remove_from_tm)
    #     self.context_menu.add_command(label="📋 TM 항목 보기", command=self.view_tm_entry)
        
    #     # 트리뷰에 우클릭 이벤트 바인딩
    #     self.translation_tree.bind("<Button-3>", self.show_context_menu)  # 우클릭
    #     self.translation_tree.bind("<Control-Button-1>", self.show_context_menu)  # Ctrl+클릭 (Mac 호환)
    #     self.translation_tree.bind("<Button-1>", self.on_tree_click)  # 기존 클릭 이벤트
                
    #     # 키보드 이벤트 바인딩
    #     self.translation_tree.bind("<KeyPress-space>", self.toggle_selected_checkboxes)
    #     self.translation_tree.bind("<Return>", self.toggle_selected_checkboxes)  # Enter키도 동일하게
    #     self.translation_tree.bind("<Control-a>", self.select_all_items)  # Ctrl+A로 전체 선택
    #     self.translation_tree.bind("<Control-d>", self.deselect_all_items)  # Ctrl+D로 전체 해제
    #     self.translation_tree.bind("<Control-i>", self.invert_selection)  # Ctrl+I: 선택 반전
    #     self.translation_tree.bind("<Delete>", self.clear_selected_translations)  # Delete: 선택된 항목의 번역 내용 삭제
    #     self.translation_tree.bind("<F2>", self.edit_selected_item)  # F2: 선택된 항목 편집

    #     # 포커스 설정으로 키보드 이벤트가 작동하도록
    #     self.translation_tree.focus_set()
        
    #     # 키보드 단축키 안내 추가
    #     help_frame = ttk.Frame(self.translation_tab)
    #     help_frame.pack(fill="x", padx=5, pady=2)
        
    #     help_text = "💡 키보드 단축키: Space/Enter=체크토글 | Ctrl+A=전체선택 | Ctrl+D=전체해제 | Ctrl+I=선택반전 | F2=편집 | Del=번역삭제"
    #     help_label = ttk.Label(help_frame, text=help_text, font=("맑은 고딕", 8), foreground="gray")
    #     help_label.pack(side="right")        
        
    #     self.check_states = {}

    # def setup_scenario_tab(self):
    #     """시나리오 번역 전용 탭 설정 (화자 로드 기능 추가)"""
    #     main_frame = ttk.Frame(self.scenario_tab, padding="10")
    #     main_frame.pack(fill="both", expand=True)
        
    #     # 시나리오 매니저 초기화 (탭 열 때마다 확인)
    #     self.ensure_scenario_manager()
        
    #     # 시나리오 번역 안내
    #     info_frame = ttk.LabelFrame(main_frame, text="ℹ️ 시나리오 번역이란?")
    #     info_frame.pack(fill="x", pady=5)
        
    #     info_text = """화자별 맞춤 번역: 각 캐릭터의 말투와 성격에 맞는 번역을 제공합니다.
    # 사용법: 1) 레퍼런스 데이터 설정 → 2) 화자 분석 → 3) 메인 설정에서 '시나리오모드' 체크 → 4) 번역 실행"""
        
    #     ttk.Label(info_frame, text=info_text, font=("맑은 고딕", 9), foreground="navy").pack(padx=10, pady=5)
        
    #     # === 1. 레퍼런스 데이터 설정 ===
    #     ref_frame = ttk.LabelFrame(main_frame, text="1️⃣ 레퍼런스 데이터 설정")
    #     ref_frame.pack(fill="x", padx=5, pady=5)
        
    #     ref_inner = ttk.Frame(ref_frame, padding="5")
    #     ref_inner.pack(fill="x")
        
    #     # 파일 선택
    #     ttk.Label(ref_inner, text="레퍼런스 파일:").grid(row=0, column=0, sticky="w")
    #     self.ref_file_var = tk.StringVar()
    #     ttk.Entry(ref_inner, textvariable=self.ref_file_var, width=40).grid(row=0, column=1, sticky="ew", padx=5)
    #     ttk.Button(ref_inner, text="엑셀 선택", command=self.select_reference_file).grid(row=0, column=2, padx=2)
    #     ttk.Button(ref_inner, text="구글 시트", command=self.load_reference_from_gsheet).grid(row=0, column=3, padx=2)
        
    #     ref_inner.grid_columnconfigure(1, weight=1)
        
    #     # 분석 버튼들
    #     analysis_frame = ttk.Frame(ref_inner)
    #     analysis_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=5)
        
    #     ttk.Button(analysis_frame, text="🔍 구조확인", command=self.debug_file_structure_detailed).pack(side="left", padx=2)
    #     ttk.Button(analysis_frame, text="🚀 자동분석", command=self.analyze_reference_data_smart).pack(side="left", padx=2)
    #     ttk.Button(analysis_frame, text="🔧 수동매핑", command=self.manual_column_mapping_dialog).pack(side="left", padx=2)
    #     ttk.Button(analysis_frame, text="📜 디버그로그", command=self.show_latest_debug_log).pack(side="left", padx=2)
    #     ttk.Button(analysis_frame, text="💾 데이터셋관리", command=self.show_reference_dataset_manager).pack(side="left", padx=2)
        
    #     # === 2. 화자 관리 ===
    #     speaker_frame = ttk.LabelFrame(main_frame, text="2️⃣ 화자 스타일 관리")
    #     speaker_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
    #     speaker_inner = ttk.Frame(speaker_frame, padding="5")
    #     speaker_inner.pack(fill="both", expand=True)
        
    #     # 화자 상태 표시 및 새로고침
    #     status_frame = ttk.Frame(speaker_inner)
    #     status_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        
    #     self.speaker_status_label = ttk.Label(status_frame, text="화자 정보 로딩 중...", foreground="blue")
    #     self.speaker_status_label.pack(side="left")
        
    #     ttk.Button(status_frame, text="🔄 새로고침", command=self.refresh_speaker_list).pack(side="right")
        
    #     # 화자 리스트
    #     columns = ("화자", "성별", "말투", "번역 스타일", "레퍼런스 수")
    #     self.speaker_tree = ttk.Treeview(speaker_inner, columns=columns, show="headings", height=8)
        
    #     # 컬럼 너비 조정
    #     widths = {"화자": 80, "성별": 50, "말투": 60, "번역 스타일": 200, "레퍼런스 수": 80}
    #     for col in columns:
    #         self.speaker_tree.heading(col, text=col)
    #         self.speaker_tree.column(col, width=widths.get(col, 100))
        
    #     speaker_scroll = ttk.Scrollbar(speaker_inner, orient="vertical", command=self.speaker_tree.yview)
    #     self.speaker_tree.configure(yscrollcommand=speaker_scroll.set)
        
    #     self.speaker_tree.grid(row=1, column=0, sticky="nsew")
    #     speaker_scroll.grid(row=1, column=1, sticky="ns")
        
    #     speaker_inner.grid_rowconfigure(1, weight=1)
    #     speaker_inner.grid_columnconfigure(0, weight=1)
        
    #     # 화자 관리 버튼들
    #     speaker_btn_frame = ttk.Frame(speaker_inner)
    #     speaker_btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        
    #     ttk.Button(speaker_btn_frame, text="➕ 추가", command=self.add_speaker).pack(side="left", padx=2)
    #     ttk.Button(speaker_btn_frame, text="✏️ 편집", command=self.edit_speaker).pack(side="left", padx=2)
    #     ttk.Button(speaker_btn_frame, text="🗑️ 삭제", command=self.delete_speaker).pack(side="left", padx=2)
        
    #     # === 3. 사용 안내 ===
    #     usage_frame = ttk.LabelFrame(main_frame, text="3️⃣ 사용 방법")
    #     usage_frame.pack(fill="x", padx=5, pady=5)
        
    #     usage_text = """✅ 설정 완료 후: 메인 설정에서 '시나리오모드' 체크 → LLM 엔진 자동 선택 → 번역 실행
    # 💡 팁: 레퍼런스 데이터는 한 번 분석하면 자동 저장되어 다음에 재사용 가능합니다."""
        
    #     ttk.Label(usage_frame, text=usage_text, font=("맑은 고딕", 9), 
    #             foreground="darkgreen").pack(padx=10, pady=5)
        
    #     # 초기 화자 리스트 로드
    #     self.refresh_speaker_list()

    # def setup_tm_management_tab(self):
    #     """번역 메모리(TM) 관리 탭 UI 구성 (하위 탭 구조 적용)"""
    #     # TM 관리 탭의 메인 프레임
    #     main_tm_frame = ttk.Frame(self.tm_management_tab, padding="5")
    #     main_tm_frame.pack(fill="both", expand=True)
        
    #     # TM 관리 탭 내에 노트북(하위 탭) 생성
    #     tm_notebook = ttk.Notebook(main_tm_frame)
    #     tm_notebook.pack(fill="both", expand=True)

    #     # 탭 1: TM 조회 및 직접 편집
    #     view_edit_tab = ttk.Frame(tm_notebook)
    #     tm_notebook.add(view_edit_tab, text="TM 조회/편집")
    #     self.setup_tm_view_edit_tab(view_edit_tab) # UI 구성 함수 호출

    #     # 탭 2: Excel로 가져오기/업데이트
    #     import_tab = ttk.Frame(tm_notebook)
    #     tm_notebook.add(import_tab, text="Excel로 가져오기/업데이트")
    #     self.setup_excel_import_tab(import_tab)

    # def setup_conflict_tab(self):
    #     """'충돌 해결' 탭의 UI를 구성합니다."""
    #     main_frame = ttk.Frame(self.conflict_tab, padding="10")
    #     main_frame.pack(fill="both", expand=True)

    #     # 1. 상단 액션 프레임
    #     action_frame = ttk.Frame(main_frame)
    #     action_frame.pack(fill="x", pady=5)
    #     ttk.Button(action_frame, text="🚀 가장 많이 쓰인 번역으로 전체 자동 해결", command=self.auto_resolve_all_conflicts).pack(side="left")
    #     ttk.Button(action_frame, text="🔄 목록 새로고침", command=self.load_conflicts_to_view).pack(side="left", padx=10)

    #     # 2. 콘텐츠 프레임 (충돌 목록 + 해결 패널)
    #     content_frame = ttk.Frame(main_frame)
    #     content_frame.pack(fill="both", expand=True, pady=5)

    #     # 2-1. 충돌 목록 Treeview
    #     list_frame = ttk.LabelFrame(content_frame, text="충돌 항목 목록")
    #     list_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

    #     columns = ["KR"] + self.VISIBLE_LANGS
    #     self.conflict_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
    #     self.conflict_tree.heading("KR", text="KR (충돌 항목)")
    #     self.conflict_tree.column("KR", width=200)
    #     for lang in self.VISIBLE_LANGS:
    #         self.conflict_tree.heading(lang, text=lang)
    #         self.conflict_tree.column(lang, width=120)
        
    #     self.conflict_tree.tag_configure('conflict', foreground='red')
    #     self.conflict_tree.bind("<<TreeviewSelect>>", self.on_conflict_row_selected)
    #     self.conflict_tree.pack(fill="both", expand=True)

    #     # 2-2. 충돌 해결 패널
    #     resolve_panel = ttk.LabelFrame(content_frame, text="충돌 해결")
    #     resolve_panel.pack(side="right", fill="y")

    #     self.conflict_kr_var = tk.StringVar()
    #     ttk.Label(resolve_panel, text="KR:").pack(anchor="w", padx=5, pady=5)
    #     ttk.Entry(resolve_panel, textvariable=self.conflict_kr_var, state="readonly", width=40).pack(anchor="w", padx=5)
        
    #     self.conflict_combos = {}
    #     for lang in self.VISIBLE_LANGS:
    #         ttk.Label(resolve_panel, text=f"{lang} 후보:").pack(anchor="w", padx=5, pady=(10, 0))
    #         combo = ttk.Combobox(resolve_panel, state="disabled", width=38)
    #         combo.pack(anchor="w", padx=5)
    #         self.conflict_combos[lang] = combo
        
    #     ttk.Button(resolve_panel, text="✅ 선택된 값으로 충돌 해결", command=self.resolve_selected_conflict).pack(pady=20, padx=5)
    
    # def setup_glossary_tab(self):
    #     """용어집 관리 탭 UI 구성 (STRING_ID 제거)"""
    #     main_frame = ttk.Frame(self.glossary_tab, padding="10")
    #     main_frame.pack(fill="both", expand=True)

    #     # 상단 컨트롤 프레임
    #     control_frame = ttk.LabelFrame(main_frame, text="용어집 관리 도구")
    #     control_frame.pack(fill="x", pady=5)
        
    #     # 좌측: 핵심 기능
    #     left_control = ttk.Frame(control_frame)
    #     left_control.pack(side="left", padx=5, pady=5)
        
    #     ttk.Button(left_control, text="🔄 구글 시트와 동기화", command=self.sync_glossary_from_gsheet, style="Accent.TButton").pack(side="left")
        
    #     # 우측: 검색 기능
    #     search_frame = ttk.Frame(control_frame)
    #     search_frame.pack(side="right", padx=10, pady=5)
        
    #     ttk.Label(search_frame, text="KR 검색:").pack(side="left")
    #     self.glossary_search_var = tk.StringVar()
    #     search_entry = ttk.Entry(search_frame, textvariable=self.glossary_search_var, width=30)
    #     search_entry.pack(side="left", padx=5)
    #     ttk.Button(search_frame, text="🔍 필터", command=self.filter_glossary).pack(side="left", padx=2)
    #     ttk.Button(search_frame, text="🔄 전체", command=self.clear_glossary_filter).pack(side="left", padx=2)
        
    #     # 사용자 안내
    #     info_label = ttk.Label(control_frame, text="ℹ️ 용어 추가/수정/삭제는 마스터 구글 시트에서 진행 후, 동기화 버튼을 눌러주세요.", foreground="blue")
    #     info_label.pack(side="bottom", padx=10, pady=2)

    #     # 용어집 목록 Treeview (STRING_ID 제거)
    #     list_frame = ttk.LabelFrame(main_frame, text="용어집 목록 (마스터 구글 시트의 로컬 사본)")
    #     list_frame.pack(fill="both", expand=True, pady=10)

    #     # 🆕 STRING_ID 제거된 컬럼 목록
    #     self.glossary_cols = ["kr", "en", "cn", "tw", "contributor", "verified"]
    #     self.glossary_tree = ttk.Treeview(list_frame, columns=self.glossary_cols, show="headings")
        
    #     # 컬럼 너비 조정
    #     col_widths = {"kr": 200, "en": 200, "cn": 120, "tw": 120, "contributor": 100, "verified": 70}
    #     for col in self.glossary_cols:
    #         self.glossary_tree.heading(col, text=col.upper())
    #         self.glossary_tree.column(col, width=col_widths.get(col, 100), anchor="w")

    #     # 스크롤바
    #     v_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.glossary_tree.yview)
    #     h_scroll = ttk.Scrollbar(list_frame, orient="horizontal", command=self.glossary_tree.xview)
    #     self.glossary_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
    #     self.glossary_tree.pack(side="left", fill="both", expand=True)
    #     v_scroll.pack(side="right", fill="y")
    #     h_scroll.pack(side="bottom", fill="x")
    
    # def setup_exclusion_tab(self):
    #     """규칙 기반 제외 목록 관리 탭 UI 구성"""
    #     main_frame = ttk.Frame(self.exclusion_tab, padding="10")
    #     main_frame.pack(fill="both", expand=True)

    #     # 1. 새 규칙 추가 프레임
    #     add_frame = ttk.LabelFrame(main_frame, text="새 규칙 추가")
    #     add_frame.pack(fill="x", pady=5)
        
    #     ttk.Label(add_frame, text="규칙 유형:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    #     self.rule_type_var = tk.StringVar(value="startswith")
    #     rule_types = ["startswith", "endswith", "contains", "equals", "length", "regex"]
    #     ttk.Combobox(add_frame, textvariable=self.rule_type_var, values=rule_types, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    #     ttk.Label(add_frame, text="적용 필드:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
    #     self.rule_field_var = tk.StringVar(value="KR")
    #     # 나중에 다른 언어 필드도 추가할 수 있도록 확장 가능하게 구성
    #     rule_fields = ["KR", "STRING_ID"] + self.VISIBLE_LANGS
    #     ttk.Combobox(add_frame, textvariable=self.rule_field_var, values=rule_fields, state="readonly").grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        
    #     ttk.Label(add_frame, text="값:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
    #     self.rule_value_var = tk.StringVar()
    #     ttk.Entry(add_frame, textvariable=self.rule_value_var).grid(row=0, column=5, padx=5, pady=5, sticky="ew")

    #     ttk.Label(add_frame, text="설명:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    #     self.rule_desc_var = tk.StringVar()
    #     ttk.Entry(add_frame, textvariable=self.rule_desc_var).grid(row=1, column=1, columnspan=5, padx=5, pady=5, sticky="ew")
        
    #     ttk.Button(add_frame, text="규칙 추가", command=self.add_exclusion_rule).grid(row=0, column=6, rowspan=2, padx=10, pady=5, ipady=10)
        
    #     add_frame.grid_columnconfigure(5, weight=1)

    #     # 2. 규칙 목록 표시 프레임
    #     list_frame = ttk.LabelFrame(main_frame, text="제외 규칙 목록")
    #     list_frame.pack(fill="both", expand=True, pady=10)

    #     columns = ("description", "rule_type", "field", "value", "enabled")
    #     self.exclusion_rule_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
    #     self.exclusion_rule_tree.heading("description", text="설명")
    #     self.exclusion_rule_tree.heading("rule_type", text="규칙 유형")
    #     self.exclusion_rule_tree.heading("field", text="적용 필드")
    #     self.exclusion_rule_tree.heading("value", text="값")
    #     self.exclusion_rule_tree.heading("enabled", text="활성화")
        
    #     self.exclusion_rule_tree.column("description", width=250)
    #     self.exclusion_rule_tree.column("rule_type", width=100)
    #     self.exclusion_rule_tree.column("field", width=100)
    #     self.exclusion_rule_tree.column("value", width=150)
    #     self.exclusion_rule_tree.column("enabled", width=80, anchor="center")
        
    #     self.exclusion_rule_tree.pack(side="left", fill="both", expand=True)
        
    #     # 3. 관리 버튼 프레임
    #     button_frame = ttk.Frame(main_frame)
    #     button_frame.pack(fill="x", pady=5)
        
    #     ttk.Button(button_frame, text="활성화/비활성화", command=self.toggle_exclusion_rule).pack(side="left")
    #     ttk.Button(button_frame, text="규칙 삭제", command=self.delete_exclusion_rule).pack(side="left", padx=10)
    #     ttk.Button(button_frame, text="기본값으로 초기화", command=self.reset_default_rules).pack(side="right")

    # def setup_history_tab(self):
    #     """번역 이력 탭 설정"""
    #     # 검색
    #     search_frame = ttk.Frame(self.history_tab)
    #     search_frame.pack(fill="x", padx=10, pady=5)
        
    #     ttk.Label(search_frame, text="검색:").pack(side="left")
    #     self.history_search_var = tk.StringVar()
    #     ttk.Entry(search_frame, textvariable=self.history_search_var, width=40).pack(side="left", padx=5)
    #     ttk.Button(search_frame, text="검색", command=self.search_history).pack(side="left")
        
    #     # 이력 테이블
    #     columns = ["시간", "STRING_ID", "KR", "번역방법", "상태"]
    #     self.history_tree = ttk.Treeview(self.history_tab, columns=columns, show="headings")
        
    #     for col in columns:
    #         self.history_tree.heading(col, text=col)
        
    #     self.history_tree.column("시간", width=150)
    #     self.history_tree.column("STRING_ID", width=150)
    #     self.history_tree.column("KR", width=300)
    #     self.history_tree.column("번역방법", width=100)
    #     self.history_tree.column("상태", width=100)
        
    #     history_scroll = ttk.Scrollbar(self.history_tab, orient="vertical", command=self.history_tree.yview)
    #     self.history_tree.configure(yscrollcommand=history_scroll.set)
        
    #     self.history_tree.pack(side="left", fill="both", expand=True, padx=10, pady=5)
    #     history_scroll.pack(side="right", fill="y")
    
    # def setup_tm_view_edit_tab(self, parent_tab):
    #     """'TM 조회/편집' 하위 탭의 UI를 구성합니다."""
    #     # 이 부분은 이전에 'setup_tm_management_tab'에 있던 내용과 거의 동일합니다.
    #     main_frame = ttk.Frame(parent_tab, padding="10")
    #     main_frame.pack(fill="both", expand=True)
        
    #     control_frame = ttk.LabelFrame(main_frame, text="TM 관리 도구")
    #     control_frame.pack(fill="x", pady=5)
        
    #     ttk.Button(control_frame, text="🔍 TM 상태 확인", command=self.debug_tm_status).pack(side="left", padx=10)

    #     build_frame = ttk.Frame(control_frame)
    #     build_frame.pack(side="left", padx=10, pady=5)

    #     ttk.Button(build_frame, text="소스 DB 폴더로 TM 구축", command=self.start_db_build).pack(side="left")

    #     mode_frame = ttk.Frame(build_frame)
    #     mode_frame.pack(side="left", padx=10)

    #     ttk.Radiobutton(mode_frame, text="충돌 우선 해결 모드 (엄격)", variable=self.db_build_mode_var, value="conflict").pack(anchor="w")
    #     ttk.Radiobutton(mode_frame, text="빈칸 채우기 모드 (보강)", variable=self.db_build_mode_var, value="fill_blanks").pack(anchor="w")
    #     ttk.Button(build_frame, text="TM 정리하기 (제외 규칙 적용)", command=self.cleanup_tm_with_rules).pack(side="left", padx=20)
        
    #     search_frame = ttk.Frame(control_frame)
    #     search_frame.pack(side="right", padx=10)
    #     ttk.Label(search_frame, text="KR 검색:").pack(side="left")
    #     self.tm_view_search_var = tk.StringVar()
    #     search_entry = ttk.Entry(search_frame, textvariable=self.tm_view_search_var, width=40)
    #     search_entry.pack(side="left", padx=5)
    #     search_entry.bind("<KeyRelease>", lambda e: self.load_tm_view())
        
    #     list_frame = ttk.LabelFrame(main_frame, text="마스터 번역 메모리 내용")
    #     list_frame.pack(fill="both", expand=True, pady=10)

    #     base_columns = ["KR"] + self.VISIBLE_LANGS
    #     self.tm_view_tree = ttk.Treeview(list_frame, columns=base_columns, show="headings")
    #     for col in base_columns:
    #         self.tm_view_tree.heading(col, text=col)
    #         self.tm_view_tree.column(col, width=150)
    #     self.tm_view_tree.column("KR", width=250)

    #     v_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tm_view_tree.yview)
    #     h_scroll = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tm_view_tree.xview)
    #     self.tm_view_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
    #     self.tm_view_tree.pack(side="left", fill="both", expand=True)
    #     v_scroll.pack(side="right", fill="y")
    #     h_scroll.pack(side="bottom", fill="x")


#### 번역 엔진 관련
    def _execute_translation_thread(self, items_to_translate):
        """(실무자 스레드) 실제 번역 작업을 지시하고 완료 후 결과를 보고합니다."""
        try:
            self.update_status("번역 준비 중...")
            selected_engine = self.api_engine_var.get()
            use_protection = self.protect_tags_var.get()
            llm_prompt = self.get_llm_prompt() if selected_engine == 'llm' else None
            
            total_items = len(items_to_translate)
            for i, trans in enumerate(items_to_translate):
                self.root.after(0, self.update_progress, (i / total_items) * 100, f"번역 중 ({i+1}/{total_items})...")
                
                kr_text = trans['KR']
                
                # 1. EN 번역
                if self.translate_en_var.get() and not trans['translations'].get('EN'):
                    en_result = self.api_client.translate(selected_engine, kr_text, prompt=llm_prompt, target_lang_code='EN-US', use_protection=use_protection)
                    if en_result: trans['translations']['EN'] = en_result
                
                # 2. 다국어 번역
                en_text = trans['translations'].get('EN')
                if en_text:
                    langs_to_translate = [lang for lang in self.MULTI_LANG_GROUP if self.translate_multi_var.get() and not trans['translations'].get(lang)]
                    if self.translate_cn_tw_var.get():
                        if not trans['translations'].get('CN'): langs_to_translate.append('CN')
                        if not trans['translations'].get('TW'): langs_to_translate.append('TW')
                    
                    for lang in langs_to_translate:
                        target_lang_code = LANG_CODES[lang][1]
                        multi_result = self.api_client.translate(selected_engine, en_text, target_lang_code=target_lang_code, source_lang_code='EN', use_protection=use_protection)
                        if multi_result: trans['translations'][lang] = multi_result

                trans['status'] = "[완료]"
                trans['method'] = selected_engine.upper()

            # 3. DB 저장 및 메모리 업데이트
            updated_krs = self.db_manager.update_translation_memory(items_to_translate)
            if updated_krs: self.translation_memory = self.db_manager.get_translation_memory()

            # 4. UI 업데이트는 메인 스레드에 요청
            self.root.after(0, self.on_translation_complete, len(updated_krs))

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: messagebox.showerror("번역 오류", f"번역 스레드에서 오류가 발생했습니다:\n{e}"))
            self.update_status("번역 중 오류 발생")


    def _force_retranslate_thread(self, string_ids):
        try:
            self.update_status("강제 재번역 시작...")
            items_to_retranslate = [trans for trans in self.pending_translations if trans["STRING_ID"] in string_ids]

            if not items_to_retranslate:
                self.update_status("재번역할 항목을 찾을 수 없습니다.")
                return

            # ### 수정됨: 이제 API 클라이언트를 사용합니다 ###
            selected_engine = self.api_engine_var.get()
            use_protection = self.protect_tags_var.get()
            llm_prompt = self.get_llm_prompt() if selected_engine == 'llm' else None
            success_count = 0

            for i, trans in enumerate(items_to_retranslate):
                kr_text = trans["KR"]
                self.update_status(f"재번역 중 ({i+1}/{len(items_to_retranslate)}): {kr_text[:20]}...")

                # EN 번역만 강제 수행
                en_result = self.api_client.translate(selected_engine, kr_text, prompt=llm_prompt, target_lang_code='EN-US', use_protection=use_protection)
                if en_result:
                    trans["translations"]["EN"] = en_result
                    trans["method"] = f"재번역({selected_engine.upper()})"
                    trans["status"] = "[재번역완료]"
                    success_count += 1
                else:
                    trans["method"] = "재번역실패"
                    trans["status"] = "[실패]"

            if success_count > 0:
                self.db_manager.update_translation_memory(items_to_retranslate)
                self.translation_memory = self.db_manager.get_translation_memory()

            self.root.after(0, self.on_force_retranslate_complete, success_count, len(items_to_retranslate))

        except Exception as e:
            self.update_status(f"재번역 오류: {e}")

    def on_force_retranslate_complete(self, success_count, total_count):
        """강제 재번역 완료 후 UI 업데이트"""
        if '재번역완료' in self.filter_vars:
            self.filter_vars['재번역완료'].set(True)
        self.update_translation_table()
        self.update_status(f"재번역 완료: {success_count}/{total_count}개 성공")
        

    def find_similar_translation(self, kr_text, threshold=0.9):
        """유사 번역 찾기"""
        best_match = None
        best_similarity = 0
        
        for saved_kr, translations in self.translation_memory.items():
            similarity = SequenceMatcher(None, kr_text, saved_kr).ratio()
            
            if similarity >= threshold and similarity > best_similarity:
                best_similarity = similarity
                best_match = {
                    "kr": saved_kr,
                    "translations": translations,
                    "similarity": similarity
                }
                
        return best_match
    
    def apply_glossary(self, kr_text):
        """용어집 적용"""
        # 간단한 구현 - 추후 고도화 가능
        if not self.glossary:
            return None
            
        # 용어집에 있는 단어가 포함되어 있는지 확인
        for kr_term, translations in self.glossary.items():
            if kr_term in kr_text:
                # 부분 번역 제공
                return {"EN": f"[{translations.get('EN', '')}...]"}
                
        return None

    def translate_with_enhanced_scenario(self, kr_text, speaker_name):
        """일관성 개선된 시나리오 번역"""
        try:
            if not self.scenario_manager:
                return None
            
            # 1. 향상된 프롬프트 생성 (유사 레퍼런스 기반)
            enhanced_prompt = self.scenario_manager.generate_enhanced_speaker_prompt(
                speaker_name, kr_text, "EN"
            )
            
            # 2. LLM 번역 실행
            translation_result = self.translate_with_llm(kr_text, enhanced_prompt)
            
            if not translation_result:
                return None
            
            # 3. 일관성 검증
            consistency_check = self.scenario_manager.validate_translation_consistency(
                speaker_name, kr_text, translation_result, "EN"
            )
            
            # 4. 일관성이 낮으면 재번역 시도
            if not consistency_check["is_consistent"] and consistency_check["confidence"] < 0.5:
                print(f"일관성 부족으로 재번역 시도: {kr_text[:30]}...")
                
                # 더 강화된 프롬프트로 재시도
                refined_prompt = enhanced_prompt + f"""

    이전 번역이 '{speaker_name}' 캐릭터의 기존 패턴과 일치하지 않습니다.
    다음 점을 개선하여 다시 번역하세요:
    {chr(10).join(['- ' + suggestion for suggestion in consistency_check["suggestions"]])}

    더 일관된 번역:"""
                
                refined_result = self.translate_with_llm(kr_text, refined_prompt)
                if refined_result:
                    translation_result = refined_result
            
            return {
                "translation": translation_result,
                "consistency": consistency_check,
                "method": f"시나리오({speaker_name})" + ("_검증" if not consistency_check["is_consistent"] else "")
            }
            
        except Exception as e:
            print(f"향상된 시나리오 번역 오류: {e}")
            return None

    def get_speaker_for_item_enhanced(self, trans_item, speaker_mapping=None):
        """화자 정보 가져오기 (캐싱 및 오류 처리 강화)"""
        try:
            string_id = trans_item["STRING_ID"]
            
            # 1. 매핑에서 찾기
            if speaker_mapping:
                for speaker, string_ids in speaker_mapping.items():
                    if string_id in string_ids:
                        return speaker
            
            # 2. 원본 파일에서 찾기 (캐싱)
            if not hasattr(self, '_speaker_cache'):
                self._speaker_cache = {}
                try:
                    file_path = self.file_path_var.get()
                    if file_path:
                        df = pd.read_excel(file_path, skiprows=3)
                        if '#화자' in df.columns and 'STRING_ID' in df.columns:
                            for _, row in df.iterrows():
                                if not pd.isna(row['STRING_ID']) and not pd.isna(row['#화자']):
                                    self._speaker_cache[str(row['STRING_ID'])] = str(row['#화자']).strip()
                except Exception as e:
                    print(f"화자 캐시 생성 오류: {e}")
            
            return self._speaker_cache.get(string_id)
            
        except Exception as e:
            print(f"화자 정보 조회 오류: {e}")
            return None

    def on_translation_complete(self, count):
        """번역 완료 후 UI 업데이트 콜백."""
        self.update_translation_table()
        self.update_status(f"번역 완료. {count}개 항목 DB 업데이트됨.")
        self.progress_bar['value'] = 100
        messagebox.showinfo("완료", f"{count}개 항목의 번역 및 저장이 완료되었습니다.")
        


    def translate_with_protection(self, text, target_lang, translator, source_lang="KO"):
        """특수 태그를 보호하면서 번역하는 함수"""
            # 사용자가 특수 태그 보호를 비활성화했으면 일반 번역
        if not self.protect_tags_var.get():
            return translator.translate_text(text, target_lang=target_lang).text
        
        # 복잡한 마크업 옵션이 활성화되어 있고 색상 태그가 있으면
        if self.complex_markup_var.get() and '[#' in text:
            return self.translate_complex_markup(text, target_lang, translator)
        
        try:
            # 1단계: 텍스트 보호
            self.text_protector.reset()
            protected_text, protection_map = self.text_protector.protect_text(text)
            
            print(f"원본: {text}")
            print(f"보호된 텍스트: {protected_text}")
            print(f"보호 맵: {protection_map}")
            
            # 2단계: 보호된 텍스트가 번역할 내용이 있는지 확인
            # 플레이스홀더만 있고 실제 텍스트가 없으면 번역하지 않음
            text_without_placeholders = re.sub(r'<PROTECTED_\d+>', '', protected_text).strip()
            if not text_without_placeholders:
                # 번역할 텍스트가 없으면 원본 반환
                return text
            
            # 3단계: 번역 실행
            translated = translator.translate_text(
                protected_text, 
                target_lang=target_lang,
                source_lang=source_lang if source_lang != "KO" else None
            ).text
            
            print(f"번역된 텍스트: {translated}")
            
            # 4단계: 보호된 요소 복원
            restored_text = self.text_protector.restore_text(translated, protection_map)
            
            print(f"복원된 텍스트: {restored_text}")
            
            return restored_text
            
        except Exception as e:
            print(f"보호된 번역 오류: {e}")
            # 오류 시 기존 방식으로 폴백
            try:
                return translator.translate_text(text, target_lang=target_lang).text
            except:
                return text

    def translate_complex_markup(self, text, target_lang, translator):
        """복잡한 마크업이 있는 텍스트의 고급 번역"""
        try:
            # 색상 태그가 있는 복잡한 텍스트 처리
            translatable_parts = self.text_protector.extract_translatable_parts(text)
            
            if not translatable_parts:
                # 번역 가능한 부분이 없으면 보호된 번역 시도
                return self.translate_with_protection(text, target_lang, translator)
            
            # 각 번역 가능한 부분을 개별 번역
            result_text = text
            for part in reversed(translatable_parts):  # 뒤에서부터 처리
                try:
                    translated_part = translator.translate_text(
                        part['text'], target_lang=target_lang
                    ).text
                    
                    # 원본 텍스트에서 해당 부분을 번역된 텍스트로 교체
                    result_text = (result_text[:part['start']] + 
                                translated_part + 
                                result_text[part['end']:])
                except Exception as e:
                    print(f"부분 번역 오류 ({part['text']}): {e}")
                    
            return result_text
            
        except Exception as e:
            print(f"복잡한 마크업 번역 오류: {e}")
            return self.translate_with_protection(text, target_lang, translator)
