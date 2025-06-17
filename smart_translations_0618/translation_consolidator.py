import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import os
import json
import time
import re
from collections import defaultdict, Counter
import openpyxl
from exclusion_manager import ExclusionManager # 새로 생성한 ExclusionManager 임포트

# 설정 로드/저장을 위한 간단한 유틸리티 함수
def load_config():
    """설정 파일(config.json)을 로드합니다."""
    try:
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                return json.load(f)
    except (IOError, json.JSONDecodeError):
        pass
    return {}

def save_config(config):
    """설정을 config.json에 저장합니다."""
    try:
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except IOError:
        pass

class TranslationConsolidator:
    def __init__(self, root):
        self.root = root
        self.root.title("번역 통합 관리 도구 v3 (로딩 개선)")
        self.root.geometry("1800x950")

        self.supported_langs = ["EN", "CN", "TW", "JP", "DE", "FR", "TH", "PT", "ES"]
        
        self.config = load_config()
        self.source_db_folder_var = tk.StringVar(value=self.config.get("source_db_folder", ""))
        self.consolidated_db_path_var = tk.StringVar(value=self.config.get("consolidated_db_path", os.path.join(os.getcwd(), "consolidated_translations.db")))
        
        # 제외 패턴 관리자 인스턴스 초기화
        # exclusion_patterns.db는 translation_consolidator.py와 같은 폴더에 생성됩니다.
        self.exclusion_manager = ExclusionManager(self.root, os.path.join(os.path.dirname(os.path.abspath(__file__)), "exclusion_patterns.db"))

        self.consolidated_data = {}
        self.conflict_data = {} # {KR: {lang: Counter({option: count, ...}), ...}}

        self._setup_ui()
        
        # UI를 먼저 그리고, 0.1초 뒤에 데이터베이스 로딩을 시작합니다.
        # 이렇게 하면 프로그램이 즉시 뜨는 것처럼 보입니다.
        self.status_label.config(text="DB 로딩 중...")
        self.root.update_idletasks() # 상태 메시지 즉시 업데이트
        self.root.after(100, self.load_consolidated_db)

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        self._setup_config_frame(main_frame)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True, pady=10)

        self.consolidated_tab = ttk.Frame(notebook)
        notebook.add(self.consolidated_tab, text=" ✅ 확정된 번역 ")
        self._setup_consolidated_tab(self.consolidated_tab) # 이제 이 메서드는 아래에 정의되어 있습니다.

        self.conflict_tab = ttk.Frame(notebook)
        notebook.add(self.conflict_tab, text=" ⚠️ 충돌 해결 ")
        self._setup_conflict_tab(self.conflict_tab) # 이제 이 메서드는 아래에 정의되어 있습니다.
        
        self.status_label = ttk.Label(main_frame, text="준비됨.", anchor="w")
        self.status_label.pack(side="bottom", fill="x")

        # 스타일 설정 (기존 코드에서 이동, 한 번만 설정)
        self.root.style = ttk.Style()
        self.root.style.configure("Accent.TButton", foreground="blue", font=('Helvetica', 10, 'bold'))
        self.root.style.configure("Conflict.TLabel", foreground="red", font=('Helvetica', 9, 'bold'))


    def _setup_config_frame(self, parent):
        config_frame = ttk.LabelFrame(parent, text="설정", padding=10)
        config_frame.pack(fill="x")

        ttk.Label(config_frame, text="소스 DB 폴더:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.source_db_folder_var, width=70).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(config_frame, text="폴더 찾기", command=self.select_source_folder).grid(row=0, column=2, padx=5)

        ttk.Label(config_frame, text="통합 DB 파일:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.consolidated_db_path_var, width=70).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(config_frame, text="파일 선택", command=self.select_consolidated_db).grid(row=1, column=2, padx=5)
        
        update_button = ttk.Button(config_frame, text="🔄 새 소스로 DB 업데이트 및 충돌 감지", command=self.process_source_dbs, style="Accent.TButton")
        update_button.grid(row=0, column=3, rowspan=2, padx=20, ipady=10)

        # "제외 패턴 관리" 버튼 추가
        manage_exclusion_button = ttk.Button(config_frame, text="🗑️ 제외 패턴 관리", command=self.open_exclusion_manager, style="Accent.TButton")
        manage_exclusion_button.grid(row=0, column=4, rowspan=2, padx=10, ipady=10)
        
        config_frame.grid_columnconfigure(1, weight=1) # Entry 위젯이 확장되도록 설정

    def open_exclusion_manager(self):
        """
        제외 패턴 관리 창을 엽니다.
        창이 이미 열려있다면 해당 창으로 포커스를 이동시킵니다.
        """
        if not self.exclusion_manager.root.winfo_exists():
            # 만약 이전 인스턴스가 닫혔다면 새로운 인스턴스 생성 (안전장치)
            self.exclusion_manager = ExclusionManager(self.root, os.path.join(os.path.dirname(os.path.abspath(__file__)), "exclusion_patterns.db"))
        else:
            self.exclusion_manager.root.lift() # 창을 최상단으로 올림
            self.exclusion_manager.root.focus_force() # 포커스 강제 부여

    # --- _setup_consolidated_tab 메서드 정의 시작 ---
    def _setup_consolidated_tab(self, parent):
        search_frame = ttk.Frame(parent)
        search_frame.pack(fill="x", pady=5)
        ttk.Label(search_frame, text="🔍 KR 검색:").pack(side="left", padx=5)
        self.consolidated_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.consolidated_search_var, width=50)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.filter_consolidated_view())
        
        ttk.Button(search_frame, text="내보내기 (Excel)", command=self.export_consolidated_data).pack(side="right", padx=5)


        content_frame = ttk.Frame(parent)
        content_frame.pack(fill="both", expand=True)
        tree_frame = ttk.Frame(content_frame)
        tree_frame.pack(side="left", fill="both", expand=True)

        columns = ("KR",) + tuple(self.supported_langs)
        self.consolidated_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.consolidated_tree.heading("KR", text="KR")
        self.consolidated_tree.column("KR", width=250, stretch=False)
        for lang in self.supported_langs:
            self.consolidated_tree.heading(lang, text=lang)
            self.consolidated_tree.column(lang, width=120, stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.consolidated_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.consolidated_tree.xview)
        self.consolidated_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.consolidated_tree.pack(fill="both", expand=True)
        self.consolidated_tree.bind("<<TreeviewSelect>>", self.on_consolidated_row_selected)

        self.edit_panel = ttk.LabelFrame(content_frame, text="번역 편집", padding=10)
        self.edit_panel.pack(side="right", fill="y", padx=10)
        self.edit_panel.grid_columnconfigure(0, weight=1)
        self.edit_fields = {}
        self.edit_kr_var = tk.StringVar()
        ttk.Label(self.edit_panel, text="KR (수정 불가)").pack(anchor="w")
        ttk.Entry(self.edit_panel, textvariable=self.edit_kr_var, state="readonly", width=40).pack(anchor="w", pady=(0, 10), fill="x")
        for lang in self.supported_langs:
            ttk.Label(self.edit_panel, text=lang).pack(anchor="w")
            entry = ttk.Entry(self.edit_panel, width=40)
            entry.pack(anchor="w", pady=2, fill="x")
            self.edit_fields[lang] = entry
        ttk.Button(self.edit_panel, text="💾 수정 저장", command=self.save_edits).pack(anchor="e", pady=10)
    # --- _setup_consolidated_tab 메서드 정의 끝 ---
    
    # --- _setup_conflict_tab 메서드 정의 시작 ---
    def _setup_conflict_tab(self, parent):
        action_frame = ttk.Frame(parent, padding=5)
        action_frame.pack(fill="x")
        
        ttk.Button(action_frame, text="🚀 가장 많이 쓰인 번역으로 전체 자동 해결", command=self.auto_resolve_all_conflicts).pack(side="left", padx=5)
        ttk.Button(action_frame, text="🔄 새로고침", command=self.refresh_conflict_view).pack(side="left", padx=5)

        content_frame = ttk.Frame(parent)
        content_frame.pack(fill="both", expand=True)

        tree_frame = ttk.Frame(content_frame)
        tree_frame.pack(side="left", fill="both", expand=True)

        columns = ("KR",) + tuple(self.supported_langs)
        self.conflict_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.conflict_tree.heading("KR", text="KR")
        self.conflict_tree.column("KR", width=250, stretch=False)
        for lang in self.supported_langs:
            self.conflict_tree.heading(lang, text=lang)
            self.conflict_tree.column(lang, width=120, stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.conflict_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.conflict_tree.xview)
        self.conflict_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.conflict_tree.pack(fill="both", expand=True)
        self.conflict_tree.bind("<<TreeviewSelect>>", self.on_conflict_row_selected)

        self.conflict_edit_panel = ttk.LabelFrame(content_frame, text="충돌 해결", padding=10)
        self.conflict_edit_panel.pack(side="right", fill="y", padx=10)
        self.conflict_edit_panel.grid_columnconfigure(0, weight=1)

        self.conflict_kr_var = tk.StringVar()
        ttk.Label(self.conflict_edit_panel, text="KR (충돌 항목)").pack(anchor="w")
        ttk.Entry(self.conflict_edit_panel, textvariable=self.conflict_kr_var, state="readonly", width=40).pack(anchor="w", pady=(0, 10), fill="x")

        self.conflict_solution_vars = {}
        self.conflict_candidate_combos = {}

        for lang in self.supported_langs:
            ttk.Label(self.conflict_edit_panel, text=f"{lang} 후보:").pack(anchor="w")
            var = tk.StringVar()
            combo = ttk.Combobox(self.conflict_edit_panel, textvariable=var, state="readonly", width=40)
            combo.pack(anchor="w", pady=2, fill="x")
            self.conflict_solution_vars[lang] = var
            self.conflict_candidate_combos[lang] = combo
            
        ttk.Button(self.conflict_edit_panel, text="✅ 선택된 값으로 충돌 해결 및 저장", command=self.save_conflict_resolution).pack(anchor="e", pady=10)
        
        self._set_conflict_edit_panel_state("disabled")
    # --- _setup_conflict_tab 메서드 정의 끝 ---


    def _set_conflict_edit_panel_state(self, state):
        for widget in self.conflict_edit_panel.winfo_children():
            if isinstance(widget, (ttk.Entry, ttk.Button)):
                widget.config(state=state)
            elif isinstance(widget, ttk.Combobox):
                widget.config(state="readonly" if state == "normal" else "disabled")


    def select_source_folder(self):
        folder = filedialog.askdirectory(title="소스 DB 폴더 선택")
        if folder: 
            self.source_db_folder_var.set(folder)
            self.config["source_db_folder"] = folder

    def select_consolidated_db(self):
        path = filedialog.asksaveasfilename(title="통합 DB 파일 저장 위치 선택", defaultextension=".db", filetypes=[("DB 파일", "*.db")])
        if path:
            self.consolidated_db_path_var.set(path)
            self.config["consolidated_db_path"] = path
            self.load_consolidated_db()

    def filter_consolidated_view(self):
        keyword = self.consolidated_search_var.get().lower()
        self.consolidated_tree.delete(*self.consolidated_tree.get_children())
        for kr, data in self.consolidated_data.items():
            if keyword in kr.lower():
                values = (kr,) + tuple(data.get(lang, "") for lang in self.supported_langs)
                self.consolidated_tree.insert("", "end", values=values)

    def on_consolidated_row_selected(self, event):
        selection = self.consolidated_tree.selection()
        if not selection: return
        item = self.consolidated_tree.item(selection[0])
        kr = item["values"][0]
        self.edit_kr_var.set(kr)
        record = self.consolidated_data.get(kr, {})
        for lang, entry in self.edit_fields.items():
            entry.delete(0, "end"); entry.insert(0, record.get(lang, ""))

    def _create_db_tables_if_not_exist(self, conn):
        cursor = conn.cursor()
        lang_cols = ", ".join([f'"{lang}" TEXT' for lang in self.supported_langs])
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS translations (
                "KR" TEXT PRIMARY KEY, {lang_cols},
                "status" TEXT NOT NULL, "conflict_info" TEXT, "last_updated" TEXT
            )""")
        conn.commit()

    def load_consolidated_db(self):
        self.status_label.config(text="DB 로딩 중...")
        self.root.update_idletasks()

        db_path = self.consolidated_db_path_var.get()
        
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            self._create_db_tables_if_not_exist(conn)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM translations")
            rows = cursor.fetchall()
            cols = [desc[0] for desc in cursor.description]
            
            self.consolidated_data.clear()
            self.conflict_data.clear()
            
            for row in rows:
                record = dict(zip(cols, row))
                kr, status = record["KR"], record["status"]
                
                if status == 'conflict':
                    conflict_info = json.loads(record.get("conflict_info") or '{}')
                    self.conflict_data[kr] = {lang: Counter(candidates) for lang, candidates in conflict_info.items()}
                else:
                    self.consolidated_data[kr] = record

            self.status_label.config(text=f"DB 로드 완료. 확정: {len(self.consolidated_data)}개, 충돌: {len(self.conflict_data)}개. UI 업데이트 중...")
            self.root.update_idletasks()
            
            self.refresh_all_views()
            self.status_label.config(text=f"준비됨. 확정: {len(self.consolidated_data)}개, 충돌: {len(self.conflict_data)}개")
        except sqlite3.Error as e:
            messagebox.showerror("DB 로드 오류", f"데이터베이스 로드 중 오류 발생: {e}\n경로: {db_path}")
            self.status_label.config(text="DB 로드 오류 발생.")
        except Exception as e:
            messagebox.showerror("오류", f"예상치 못한 오류 발생: {e}")
            self.status_label.config(text="오류 발생.")
        finally:
            if conn: conn.close()

    def _should_exclude_text(self, text, exclude_patterns):
        """
        주어진 텍스트가 제외 패턴 목록 중 하나라도 일치하는지 확인합니다.
        """
        if not isinstance(text, str):
            return True
        
        for pattern in exclude_patterns:
            try:
                if re.search(pattern, text):
                    return True
            except re.error:
                print(f"경고: 유효하지 않은 정규식 패턴 '{pattern}'이 감지되어 무시됩니다.")
                continue
        return False

    def process_source_dbs(self):
        source_folder = self.source_db_folder_var.get()
        if not os.path.isdir(source_folder):
            messagebox.showwarning("경고", "유효한 소스 DB 폴더를 선택하세요.")
            return
            
        source_files = [os.path.join(source_folder, f) for f in os.listdir(source_folder) if f.startswith("String") and f.endswith(".db")]
        if not source_files:
            messagebox.showinfo("정보", "선택한 폴더에 처리할 String DB 파일이 없습니다.")
            return

        self.status_label.config(text="소스 DB 처리 및 충돌 감지 중...")
        self.root.update_idletasks()

        # 최신 제외 패턴을 ExclusionManager에서 로드
        if self.exclusion_manager.conn:
            self.exclusion_manager._load_patterns_from_db() 
        
        kr_exclude_patterns = self.exclusion_manager.get_kr_patterns()
        lang_exclude_patterns = self.exclusion_manager.get_lang_patterns()

        raw_data = defaultdict(lambda: {lang: Counter() for lang in self.supported_langs})
        
        for db_file in source_files:
            try:
                conn = sqlite3.connect(db_file)
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE '%String%'")
                tables = [r[0] for r in cursor.fetchall()]
                for table in tables:
                    if not re.match(r'^[a-zA-Z0-9_]+$', table): continue
                    cursor.execute(f'SELECT * FROM "{table}"')
                    cols = [desc[0].upper() for desc in cursor.description]
                    for row in cursor.fetchall():
                        entry = dict(zip(cols, row))
                        kr = entry.get("KR", "").strip()
                        
                        # --- KR 텍스트 예외 처리 적용 ---
                        if not kr or self._should_exclude_text(kr, kr_exclude_patterns):
                            continue
                        
                        for lang in self.supported_langs:
                            if lang in entry and entry[lang]:
                                lang_text = str(entry[lang]).strip()
                                # --- 번역 텍스트 예외 처리 적용 ---
                                if not self._should_exclude_text(lang_text, lang_exclude_patterns):
                                    raw_data[kr][lang][lang_text] += 1
            except Exception as e: 
                print(f"Error processing {db_file}: {e}")
                self.status_label.config(text=f"소스 DB 파일 처리 중 오류 발생: {os.path.basename(db_file)}")
                self.root.update_idletasks()
            finally:
                if conn: conn.close()
        
        db_path = self.consolidated_db_path_var.get()
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            self._create_db_tables_if_not_exist(conn)
            cursor = conn.cursor()
            
            cursor.execute("SELECT KR, status FROM translations")
            existing_records = {row[0]: row[1] for row in cursor.fetchall()}

            for kr, lang_map in raw_data.items():
                is_conflict = any(len(counter.keys()) > 1 for counter in lang_map.values())
                
                if is_conflict:
                    status = "conflict"
                    conflict_info = {lang: dict(counter) for lang, counter in lang_map.items() if counter}
                    
                    record = {lang: counter.most_common(1)[0][0] if counter else "" for lang, counter in lang_map.items()}
                    
                    update_query = f"""INSERT OR REPLACE INTO translations (KR, {', '.join(f'"{lang}"' for lang in self.supported_langs)}, status, conflict_info, last_updated)
                                       VALUES (?, {', '.join(['?'] * len(self.supported_langs))}, ?, ?, ?)"""
                    values = [kr] + [record.get(lang, "") for lang in self.supported_langs] + \
                             [status, json.dumps(conflict_info, ensure_ascii=False), time.strftime("%Y-%m-%d %H:%M:%S")]
                    cursor.execute(update_query, values)
                else:
                    if existing_records.get(kr) == 'conflict':
                        continue
                    
                    status = "consolidated"
                    record = {lang: list(counter.keys())[0] if counter else "" for lang, counter in lang_map.items()}
                    insert_query = f"""INSERT OR IGNORE INTO translations (KR, {', '.join(f'"{lang}"' for lang in self.supported_langs)}, status, conflict_info, last_updated)
                                       VALUES (?, {', '.join(['?'] * len(self.supported_langs))}, ?, ?, ?)"""
                    values = [kr] + [record.get(lang, "") for lang in self.supported_langs] + \
                             [status, None, time.strftime("%Y-%m-%d %H:%M:%S")]
                    cursor.execute(insert_query, values)
            
            conn.commit()
            messagebox.showinfo("업데이트 완료", "DB 업데이트 및 충돌 감지가 완료되었습니다.")
            self.load_consolidated_db()
        except sqlite3.Error as e:
            messagebox.showerror("DB 처리 오류", f"데이터베이스 업데이트 중 오류 발생: {e}")
            self.status_label.config(text="DB 처리 오류 발생.")
        except Exception as e:
            messagebox.showerror("오류", f"예상치 못한 오류 발생: {e}")
            self.status_label.config(text="오류 발생.")
        finally:
            if conn: conn.close()

    def save_edits(self):
        kr = self.edit_kr_var.get()
        if not kr: return messagebox.showwarning("경고", "수정할 항목이 선택되지 않았습니다.")
        lang_updates = {lang: entry.get() for lang, entry in self.edit_fields.items()}
        db_path = self.consolidated_db_path_var.get()
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            set_clause = ", ".join([f'"{lang}" = ?' for lang in self.supported_langs])
            values = [lang_updates.get(lang, "") for lang in self.supported_langs] + [time.strftime("%Y-%m-%d %H:%M:%S"), kr]
            cursor.execute(f"UPDATE translations SET {set_clause}, last_updated = ? WHERE KR = ?", values)
            conn.commit()
            conn.close()
            self.consolidated_data[kr].update(lang_updates)
            self.refresh_consolidated_view()
            messagebox.showinfo("저장 완료", "수정된 내용이 DB에 저장되었습니다.")
        except sqlite3.Error as e:
            messagebox.showerror("저장 오류", f"DB 저장 중 오류 발생: {e}")
        except Exception as e:
            messagebox.showerror("오류", f"예상치 못한 오류 발생: {e}")
        finally:
            if conn: conn.close()

    def refresh_all_views(self):
        self.refresh_consolidated_view()
        self.refresh_conflict_view()

    def refresh_consolidated_view(self):
        self.consolidated_tree.delete(*self.consolidated_tree.get_children())
        for kr, data in sorted(self.consolidated_data.items()):
            values = (kr,) + tuple(data.get(lang, "") for lang in self.supported_langs)
            self.consolidated_tree.insert("", "end", values=values)
    
    def refresh_conflict_view(self):
        self.conflict_tree.delete(*self.conflict_tree.get_children())
        self._set_conflict_edit_panel_state("disabled")
        self.conflict_kr_var.set("")
        for lang_var in self.conflict_solution_vars.values():
            lang_var.set("")
        for combo in self.conflict_candidate_combos.values():
            combo.config(values=[])

        for kr, lang_counters in sorted(self.conflict_data.items()):
            has_conflict_any_lang = any(len(counter) > 1 for counter in lang_counters.values() if counter)
            
            display_values = []
            for lang in self.supported_langs:
                counter = lang_counters.get(lang)
                if counter:
                    if len(counter) > 1:
                        display_text = f"{counter.most_common(1)[0][0]} ({counter.most_common(1)[0][1]}회) ⚠️"
                    else:
                        display_text = list(counter.keys())[0]
                else:
                    display_text = ""
                display_values.append(display_text)
            
            tags = ()
            if has_conflict_any_lang:
                tags = ('conflict',)
            
            self.conflict_tree.insert("", "end", values=(kr,) + tuple(display_values), tags=tags)
        
        self.conflict_tree.tag_configure('conflict', foreground='red', font=('Helvetica', 9, 'bold'))

    def on_conflict_row_selected(self, event):
        selection = self.conflict_tree.selection()
        if not selection: 
            self._set_conflict_edit_panel_state("disabled")
            self.conflict_kr_var.set("")
            for lang_var in self.conflict_solution_vars.values():
                lang_var.set("")
            for combo in self.conflict_candidate_combos.values():
                combo.config(values=[])
            return

        item = self.conflict_tree.item(selection[0])
        kr = item["values"][0]
        
        self._set_conflict_edit_panel_state("normal")
        self.conflict_kr_var.set(kr)

        conflict_record = self.conflict_data.get(kr, {})
        
        for lang in self.supported_langs:
            current_combo = self.conflict_candidate_combos[lang]
            current_var = self.conflict_solution_vars[lang]
            
            counter = conflict_record.get(lang)
            if counter:
                candidates_display = [f"{text} ({count}회)" for text, count in counter.most_common()]
                real_values = [text for text, count in counter.most_common()]
                
                current_combo.config(values=candidates_display)
                if candidates_display:
                    current_combo.set(candidates_display[0])
                
                current_combo.real_values = real_values
                current_combo.config(state="readonly")
            else:
                current_combo.config(values=[])
                current_var.set("")
                current_combo.config(state="disabled")

    def save_conflict_resolution(self):
        kr = self.conflict_kr_var.get()
        if not kr: 
            messagebox.showwarning("경고", "해결할 충돌 항목이 선택되지 않았습니다.")
            return

        if not messagebox.askyesno("확인", f"'{kr}' 항목의 충돌을 선택된 값으로 확정하고 저장하시겠습니까?"):
            return

        final_translations = {}
        for lang in self.supported_langs:
            combo = self.conflict_candidate_combos[lang]
            if combo.cget("state") != "disabled":
                selected_idx = combo.current()
                if selected_idx != -1 and hasattr(combo, 'real_values'):
                    final_translations[lang] = combo.real_values[selected_idx]
                else:
                    final_translations[lang] = self.conflict_solution_vars[lang].get()
            else:
                counter = self.conflict_data[kr].get(lang)
                if counter:
                    final_translations[lang] = list(counter.keys())[0]
                else:
                    final_translations[lang] = ""

        db_path = self.consolidated_db_path_var.get()
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            set_clause = ", ".join([f'"{lang}" = ?' for lang in self.supported_langs])
            values = [final_translations.get(lang, "") for lang in self.supported_langs] + \
                     ["consolidated", None, time.strftime("%Y-%m-%d %H:%M:%S"), kr]
            cursor.execute(f"UPDATE translations SET {set_clause}, status = ?, conflict_info = ?, last_updated = ? WHERE KR = ?", values)
            
            conn.commit()
            messagebox.showinfo("저장 완료", f"'{kr}' 항목의 충돌이 해결되어 DB에 저장되었습니다.")
            
            self.load_consolidated_db()

        except sqlite3.Error as e:
            messagebox.showerror("저장 오류", f"충돌 해결 저장 중 오류 발생: {e}")
        except Exception as e:
            messagebox.showerror("오류", f"예상치 못한 오류 발생: {e}")
        finally:
            if conn: conn.close()

    def auto_resolve_all_conflicts(self):
        if not messagebox.askyesno("경고", "모든 충돌 항목에 대해 '가장 많이 사용된 번역'을 정답으로 간주하여 자동으로 해결합니다. 이 작업은 되돌릴 수 없습니다. 계속하시겠습니까?"):
            return

        resolved_count = 0
        db_path = self.consolidated_db_path_var.get()
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            for kr, lang_counters in self.conflict_data.items():
                final_translations = {}
                for lang, counter in lang_counters.items():
                    if counter:
                        final_translations[lang] = counter.most_common(1)[0][0]
                    else:
                        final_translations[lang] = ""
                
                set_clause = ", ".join([f'"{lang}" = ?' for lang in self.supported_langs])
                values = [final_translations.get(lang, "") for lang in self.supported_langs] + \
                         ["consolidated", None, time.strftime("%Y-%m-%d %H:%M:%S"), kr]
                cursor.execute(f"UPDATE translations SET {set_clause}, status = ?, conflict_info = ?, last_updated = ? WHERE KR = ?", values)
                resolved_count += 1
                
            conn.commit()
            messagebox.showinfo("자동 해결 완료", f"{resolved_count}개의 충돌 항목이 자동으로 해결되었습니다.")
            self.load_consolidated_db()
        except sqlite3.Error as e:
            messagebox.showerror("자동 해결 오류", f"자동 해결 중 DB 오류 발생: {e}")
        except Exception as e:
            messagebox.showerror("자동 해결 오류", f"자동 해결 중 오류 발생: {e}")
        finally:
            if conn: conn.close()

    def export_consolidated_data(self):
        if not self.consolidated_data:
            messagebox.showinfo("정보", "내보낼 데이터가 없습니다.")
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not filepath:
            return

        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Consolidated Translations"

            headers = ["KR"] + self.supported_langs
            sheet.append(headers)

            for kr, data in sorted(self.consolidated_data.items()):
                row = [kr] + [data.get(lang, "") for lang in self.supported_langs]
                sheet.append(row)
            
            workbook.save(filepath)
            
            messagebox.showinfo("내보내기 완료", f"데이터가 '{os.path.basename(filepath)}' 에 성공적으로 내보내졌습니다.")
        except ImportError:
            messagebox.showerror("오류", "openpyxl 라이브러리가 설치되어 있지 않습니다.\n'pip install openpyxl' 명령으로 설치해주세요.")
        except Exception as e:
            messagebox.showerror("내보내기 오류", f"데이터 내보내기 중 오류 발생: {e}")


def run_translation_consolidator(parent=None):
    if parent: root = tk.Toplevel(parent)
    else: root = tk.Tk()
    app = TranslationConsolidator(root)
    def on_closing():
        save_config(app.config)
        # ExclusionManager의 DB 연결도 함께 닫아줍니다.
        if app.exclusion_manager and app.exclusion_manager.conn:
            app.exclusion_manager.conn.close()
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_closing)
    if not parent: root.mainloop()
    return app

if __name__ == '__main__':
    run_translation_consolidator()