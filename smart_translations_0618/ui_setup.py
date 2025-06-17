import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES

class UISetup:
    def __init__(self, manager):
        self.manager = manager
    
    def setup_ui(self):
        """UI 구성 (좌우 분할 레이아웃으로 개선)"""
        # 메인 프레임을 스크롤 가능하게 설정
        canvas = tk.Canvas(self.manager.root)
        scrollbar = ttk.Scrollbar(self.manager.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 메인 프레임
        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # main_frame 영역을 파일 드랍 대상으로 지정
        main_frame.drop_target_register(DND_FILES)
        main_frame.dnd_bind('<<Drop>>', self.manager.handle_drop)
        
        # === 1. 파일 선택 영역 ===
        self._setup_file_section(main_frame)
        
        # === 2. 번역 설정 & LLM 프롬프트 영역 ===
        self._setup_settings_section(main_frame)
        
        # === 3. 중앙: 탭 컨트롤 ===
        self._setup_tab_control(main_frame)
        
        # === 4. 하단: 실행 버튼 및 상태 ===
        self._setup_bottom_section(main_frame)
        
        # === 5. 진행 상황 ===
        self._setup_progress_section(main_frame)
        
        # 스크롤바 설정
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 마우스 휠 스크롤 지원
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _setup_file_section(self, main_frame):
        """파일 선택 영역 설정"""
        file_frame = ttk.LabelFrame(main_frame, text="📁 파일 선택")
        file_frame.pack(fill="x", pady=5)
        
        file_inner = ttk.Frame(file_frame, padding="5")
        file_inner.pack(fill="x")
        
        ttk.Label(file_inner, text="대상 파일:").grid(row=0, column=0, sticky="w")
        ttk.Entry(file_inner, textvariable=self.manager.file_path_var, width=50).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(file_inner, text="찾기", command=self.manager.select_file).grid(row=0, column=2, padx=2)
        ttk.Button(file_inner, text="로드", command=self.manager.load_data).grid(row=0, column=3, padx=2)
        
        file_inner.grid_columnconfigure(1, weight=1)

    def _setup_settings_section(self, main_frame):
        """번역 설정 & LLM 프롬프트 영역 설정"""
        settings_container = ttk.Frame(main_frame)
        settings_container.pack(fill="x", pady=5)
        
        # 좌측: 번역 설정 영역
        self._setup_left_settings(settings_container)
        
        # 우측: LLM 프롬프트 설정 영역
        self._setup_llm_settings(settings_container)

    def _setup_left_settings(self, settings_container):
        """좌측 번역 설정 영역"""
        left_settings_frame = ttk.LabelFrame(settings_container, text="⚙️ 번역 설정")
        left_settings_frame.pack(side="left", fill="both", expand=False, padx=(0, 5))
        
        settings_inner = ttk.Frame(left_settings_frame, padding="10")
        settings_inner.pack(fill="both", expand=True)
        
        # 번역 엔진 선택
        engine_label_frame = ttk.LabelFrame(settings_inner, text="번역 엔진")
        engine_label_frame.pack(fill="x", pady=2)
        
        engine_frame = ttk.Frame(engine_label_frame, padding="5")
        engine_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Radiobutton(engine_frame, text="Azure", variable=self.manager.api_engine_var, 
                    value="azure", command=self.manager.on_engine_changed).pack(anchor="w")
        ttk.Radiobutton(engine_frame, text="LLM", variable=self.manager.api_engine_var, 
                    value="llm", command=self.manager.on_engine_changed).pack(anchor="w")
        ttk.Radiobutton(engine_frame, text="💎 Gemini", variable=self.manager.api_engine_var,
                    value="gemini",command=self.manager.on_engine_changed).pack(anchor="w")
        
        # 번역 옵션
        options_label_frame = ttk.LabelFrame(settings_inner, text="번역 옵션")
        options_label_frame.pack(fill="x", pady=2)
        
        options_inner = ttk.Frame(options_label_frame, padding="5")
        options_inner.pack(fill="x")
        
        ttk.Checkbutton(options_inner, text="EN", variable=self.manager.translate_en_var).pack(anchor="w")
        ttk.Checkbutton(options_inner, text="다국어 (TH,PT,ES,FR,DE)", variable=self.manager.translate_multi_var).pack(anchor="w")
        ttk.Checkbutton(options_inner, text="CN/TW", variable=self.manager.translate_cn_tw_var).pack(anchor="w")
        
        # 고급 옵션
        advanced_label_frame = ttk.LabelFrame(settings_inner, text="고급 옵션")
        advanced_label_frame.pack(fill="x", pady=2)
        
        advanced_inner = ttk.Frame(advanced_label_frame, padding="5")
        advanced_inner.pack(fill="x")
        
        # manager의 변수들을 사용 (UI에서 새로 정의하지 않음)
        ttk.Checkbutton(advanced_inner, text="특수 태그 보호", variable=self.manager.protect_tags_var).pack(anchor="w")
        ttk.Checkbutton(advanced_inner, text="복잡한 마크업 처리", variable=self.manager.complex_markup_var).pack(anchor="w")
        ttk.Checkbutton(advanced_inner, text="🎭 시나리오 번역 모드", variable=self.manager.scenario_translation_var, 
                    command=self.manager.on_scenario_option_changed).pack(anchor="w")

    def _setup_llm_settings(self, settings_container):
        """우측 LLM 프롬프트 설정 영역"""
        self.manager.llm_settings_frame = ttk.LabelFrame(settings_container, text="🤖 LLM 프롬프트 설정")
        self.manager.llm_settings_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
        
        prompt_inner = ttk.Frame(self.manager.llm_settings_frame, padding="10")
        prompt_inner.pack(fill="both", expand=True)
        
        # 프롬프트 템플릿 버튼들
        template_frame = ttk.Frame(prompt_inner)
        template_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Label(template_frame, text="프롬프트 템플릿:", font=("맑은 고딕", 9, "bold")).pack(side="left")
        
        template_buttons_frame = ttk.Frame(template_frame)
        template_buttons_frame.pack(side="right")
        
        ttk.Button(template_buttons_frame, text="게임", command=lambda: self.manager.set_prompt_template("game")).pack(side="left", padx=1)
        ttk.Button(template_buttons_frame, text="일반", command=lambda: self.manager.set_prompt_template("natural")).pack(side="left", padx=1)
        ttk.Button(template_buttons_frame, text="기술", command=lambda: self.manager.set_prompt_template("technical")).pack(side="left", padx=1)
        ttk.Button(template_buttons_frame, text="캐주얼", command=lambda: self.manager.set_prompt_template("casual")).pack(side="left", padx=1)
        ttk.Button(template_buttons_frame, text="초기화", command=lambda: self.manager.set_prompt_template("default")).pack(side="left", padx=1)
        
        # 프롬프트 입력 영역
        ttk.Label(prompt_inner, text="프롬프트 내용:", font=("맑은 고딕", 9)).pack(anchor="w", pady=(5, 2))
        
        prompt_text_frame = ttk.Frame(prompt_inner)
        prompt_text_frame.pack(fill="both", expand=True)
        
        self.manager.llm_prompt_entry = tk.Text(prompt_text_frame, height=8, wrap="word", font=("맑은 고딕", 9))
        prompt_scrollbar = ttk.Scrollbar(prompt_text_frame, orient="vertical", command=self.manager.llm_prompt_entry.yview)
        self.manager.llm_prompt_entry.configure(yscrollcommand=prompt_scrollbar.set)
        
        self.manager.llm_prompt_entry.pack(side="left", fill="both", expand=True)
        prompt_scrollbar.pack(side="right", fill="y")
        
        # 기본 프롬프트 설정
        default_prompt = """한국어를 영어로 번역해주세요. 10-20대 영미권 사용자가 이해하기 쉽게, 간결하고 자연스럽게 번역하며 특수 태그는 그대로 유지하세요."""
        self.manager.llm_prompt_entry.insert("1.0", default_prompt)

    def _setup_tab_control(self, main_frame):
        """탭 컨트롤 설정"""
        tab_control = ttk.Notebook(main_frame)
        tab_control.pack(fill="both", expand=True, pady=5)
        
        # 탭들 설정
        self.manager.translation_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.translation_tab, text="📝 번역 대상")
        self.manager.setup_translation_tab()
        
        self.manager.scenario_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.scenario_tab, text="🎭 시나리오 번역")
        self.manager.setup_scenario_tab()
        
        self.manager.tm_management_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.tm_management_tab, text="💾 TM 관리")
        self.manager.setup_tm_management_tab()
        
        self.manager.conflict_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.conflict_tab, text="⚠️ 충돌 해결")
        self.manager.setup_conflict_tab()
        
        self.manager.glossary_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.glossary_tab, text="📚 용어집")
        self.manager.setup_glossary_tab()
        
        self.manager.exclusion_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.exclusion_tab, text="🚫 제외 목록")
        self.manager.setup_exclusion_tab()
        
        self.manager.history_tab = ttk.Frame(tab_control)
        tab_control.add(self.manager.history_tab, text="📈 번역 이력")
        self.manager.setup_history_tab()
        
        # 텍스트 분석 탭 추가 (새로운 기능)
        if hasattr(self.manager, 'setup_text_analysis_tab'):
            self.manager.text_analysis_tab = ttk.Frame(tab_control)
            tab_control.add(self.manager.text_analysis_tab, text="📊 텍스트 분석")
            self.manager.setup_text_analysis_tab()
        
        def on_tab_changed(event):
            selected_tab = event.widget.tab('current')['text']
            if selected_tab == "🎭 시나리오 번역":
                self.manager.root.after(100, self.manager.refresh_speaker_list)
        
        tab_control.bind("<<NotebookTabChanged>>", on_tab_changed)

    def _setup_bottom_section(self, main_frame):
        """하단 실행 버튼 및 상태 설정 - 텍스트 분석 기능 포함"""
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="x", pady=5)
        
        # 첫 번째 버튼 행 - 기본 번역 기능
        button_row1 = ttk.Frame(bottom_frame)
        button_row1.pack(fill="x")
        
        ttk.Button(button_row1, text="🔍 번역 분석", command=self.manager.analyze_translations).pack(side="left", padx=2)
        ttk.Button(button_row1, text="🚀 자동 번역", command=self.manager.execute_translation).pack(side="left", padx=2)
        ttk.Button(button_row1, text="💾 결과 저장", command=self.manager.save_results).pack(side="left", padx=2)
        
        # 상태 표시
        self.manager.status_label = ttk.Label(button_row1, text="준비됨", foreground="blue")
        self.manager.status_label.pack(side="right", padx=10)
        
        # 두 번째 버튼 행 - 고급 기능
        button_row2 = ttk.Frame(bottom_frame)
        button_row2.pack(fill="x", pady=2)
        
        ttk.Button(button_row2, text="🔄 선택 재번역", command=self.manager.force_retranslate_selected).pack(side="left", padx=2)
        ttk.Button(button_row2, text="🗑️ 선택 TM삭제", command=self.manager.remove_from_tm).pack(side="left", padx=2)
        
        # 세 번째 버튼 행 - 텍스트 분석 기능 (새로 추가)
        button_row3 = ttk.Frame(bottom_frame)
        button_row3.pack(fill="x", pady=2)
        
        # 구분선
        separator = ttk.Separator(button_row3, orient='vertical')
        separator.pack(side="left", fill="y", padx=5)
        
        ttk.Label(button_row3, text="📊 텍스트 분석:", font=("맑은 고딕", 9, "bold")).pack(side="left", padx=5)
        
        # 텍스트 분석 기능들 (manager에 해당 메서드가 있는 경우에만 표시)
        if hasattr(self.manager, 'batch_analyze_translations'):
            ttk.Button(button_row3, text="📈 일괄 분석", command=self.manager.batch_analyze_translations).pack(side="left", padx=2)
        
        if hasattr(self.manager, 'show_text_processing_stats'):
            ttk.Button(button_row3, text="📊 처리 통계", command=self.manager.show_text_processing_stats).pack(side="left", padx=2)
        
        if hasattr(self.manager, 'reset_processing_stats'):
            ttk.Button(button_row3, text="🔄 통계 초기화", command=self.manager.reset_processing_stats).pack(side="left", padx=2)
        
        # 고급 번역 버튼 (강화된 전처리 포함)
        if hasattr(self.manager, 'execute_translation_with_enhanced_preprocessing'):
            ttk.Button(button_row3, text="🚀⚡ 고급 번역", 
                      command=self.manager.execute_translation_with_enhanced_preprocessing,
                      style="Accent.TButton").pack(side="right", padx=2)

    def _setup_progress_section(self, main_frame):
        """진행 상황 설정"""
        self.manager.progress_var = tk.DoubleVar()
        self.manager.progress_bar = ttk.Progressbar(main_frame, variable=self.manager.progress_var)
        self.manager.progress_bar.pack(fill="x", pady=2)