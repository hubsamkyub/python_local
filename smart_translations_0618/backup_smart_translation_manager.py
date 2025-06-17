import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
from dotenv import load_dotenv
import sqlite3
import json
import time
from datetime import datetime
import pandas as pd
import re
from difflib import SequenceMatcher
import threading
from collections import defaultdict
import deepl
import requests
import uuid
from collections import defaultdict, Counter
import openpyxl
from datetime import datetime
from typing import List, Dict, Optional, Tuple

# --- 내부 모듈 Import ---
import file_handler
from database_manager import DatabaseManager
from translation_client import TranslationApiClient
from scenario_manager import ScenarioTranslationManager
from utils import TextProtector, ScrollableCheckList
from dialogs.edit_dialogs import InlineEditDialog
from dialogs.preview_dialogs import UpdatePreviewDialog, TranslationReportDialog
from dialogs.selection_dialogs import LanguageSelectionDialog
from config import LANG_CODES # API 키는 이제 config 파일에서만 관리
from ui_setup import UISetup
from tab_setups import TabSetups
from translation_engine import TranslationEngine
from text_processor import TextProcessor
from translation_validator import TranslationValidator

class SmartTranslationManager:
    def __init__(self, root):
        self.root = root
        self.root.title("스마트 번역 자동화 시스템")
        self.root.geometry("1900x900")
        self.root.minsize(1200, 700)

        # --- 핵심 속성 초기화 ---
        self.pending_translations = []
        self.translation_memory = {}
        self.glossary = {}
        self.check_states = {}
        self.MULTI_LANG_GROUP = ["TH", "PT", "FR", "DE", "ES"]
        self.VISIBLE_LANGS = ["EN", "CN", "TW", "TH", "PT", "ES", "FR", "DE"]
        self.translation_db_path = "smart_translations.db"        
        self.quality_history = []
        self.current_batch_quality = {}

        # --- 전문가 매니저 클래스 인스턴스화 ---
        self.db_manager = DatabaseManager(
            db_path = "smart_translations",
            legacy_db_path="unique_texts.db"
        )
        self.text_protector = TextProtector()
        self.api_client = TranslationApiClient(self.text_protector)
        self.scenario_manager = ScenarioTranslationManager(self.db_manager.db_path)
        
        # --- UI 변수 초기화 ---
        self.file_path_var = tk.StringVar()
        self.api_engine_var = tk.StringVar(value="llm")
        self.translate_en_var = tk.BooleanVar(value=True)
        self.translate_multi_var = tk.BooleanVar(value=False)
        self.translate_cn_tw_var = tk.BooleanVar(value=False)
        self.protect_tags_var = tk.BooleanVar(value=True)
        self.complex_markup_var = tk.BooleanVar(value=True)
        self.scenario_translation_var = tk.BooleanVar(value=False)
        self.db_build_mode_var = tk.StringVar(value="conflict")
        self.excel_import_folder_var = tk.StringVar()
        self.excel_import_files = [] # (파일명, 파일경로) 튜플 저장
        self.excel_import_lang_vars = {} # 언어 선택 체크박스 변수 저장
        
        # UI 설정 클래스 초기화 (이 라인을 추가)        
        self.ui_setup = UISetup(self) 
        self.tab_setups = TabSetups(self)        
        self.translation_engine = TranslationEngine(self)
        self.text_processor = TextProcessor(self)
        self.translation_validator = TranslationValidator(self)
        
        # --- UI 구성 및 데이터 로드 ---
        self.setup_ui()
        self.initialize_data()

    def initialize_data(self):
        """프로그램 시작 시 필요한 데이터를 로드하고 UI를 업데이트합니다."""
        ### 수정된 부분: 설정 파일 확인 로직 추가 ###
        missing_files = file_handler.check_config_files()
        if missing_files:
            msg = f"필요한 설정 파일이 없습니다:\n\n" + "\n".join(missing_files) + "\n\n템플릿 파일을 생성하시겠습니까?"
            if messagebox.askyesno("설정 파일 없음", msg):
                file_handler.create_config_templates()
                messagebox.showinfo("생성 완료", "설정 파일 템플릿이 생성되었습니다. API 키 등을 입력 후 프로그램을 다시 시작해주세요.")
                self.root.quit()
                return # 프로그램 종료
            else:
                self.root.quit()
                return # 프로그램 종료
        self.db_manager.init_database()
        self.translation_memory = self.db_manager.get_translation_memory()
        self.load_glossary_and_update_ui()
        self.load_exclusion_rules_and_update_ui()
        self.load_conflicts_to_view()
        self.update_status(f"초기화 완료. TM: {len(self.translation_memory)}개")
 
    def setup_ui(self):
        """UI 구성을 ui_setup 모듈에 위임"""
        self.ui_setup.setup_ui()

    def setup_translation_tab(self):
        """번역 대상 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_translation_tab()

    def setup_scenario_tab(self):
        """시나리오 번역 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_scenario_tab()

    def setup_tm_management_tab(self):
        """TM 관리 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_tm_management_tab()

    def setup_conflict_tab(self):
        """충돌 해결 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_conflict_tab()

    def setup_glossary_tab(self):
        """용어집 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_glossary_tab()

    def setup_exclusion_tab(self):
        """제외 목록 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_exclusion_tab()

    def setup_history_tab(self):
        """번역 이력 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_history_tab()

    def setup_tm_view_edit_tab(self, parent_tab):
        """TM 조회/편집 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_tm_view_edit_tab(parent_tab)

    def setup_excel_import_tab(self, parent_tab):
        """Excel 가져오기 탭 설정을 tab_setups 모듈에 위임"""
        self.tab_setups.setup_excel_import_tab(parent_tab)

    def load_data_from_file(self):
        """선택된 엑셀 파일에서 번역할 데이터를 로드합니다."""
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("경고", "파일을 선택하세요.")
            return

        try:
            self.update_status("엑셀 파일 로드 중...")
            self.translation_memory = self.db_manager.get_translation_memory()
            active_rules = self.db_manager.get_exclusion_rules(only_enabled=True)
            
            df = pd.read_excel(file_path, skiprows=3) # 헤더가 4번째 줄에 있다고 가정
            if "STRING_ID" not in df.columns or "KR" not in df.columns:
                raise ValueError("엑셀 파일에 'STRING_ID'와 'KR' 컬럼이 반드시 필요합니다.")

            self.pending_translations.clear()
            excluded_count = 0
            
            for _, row in df.iterrows():
                if pd.isna(row["KR"]) or not str(row["KR"]).strip():
                    continue

                entry_dict = row.to_dict()
                if self.db_manager._is_entry_excluded(entry_dict, active_rules): # 임시로 직접 호출
                    excluded_count += 1
                    continue
                
                kr_text = str(row["KR"]).strip()
                item = {
                    "STRING_ID": str(row["STRING_ID"]), "KR": kr_text,
                    "status": self.determine_status(kr_text),
                    "translations": {lang: str(row[lang]) for lang in self.VISIBLE_LANGS if lang in df.columns and pd.notna(row[lang])},
                    "method": ""
                }
                self.pending_translations.append(item)
            
            self.update_translation_table()
            self.update_status(f"로드 완료: {len(self.pending_translations)}개 항목 (제외: {excluded_count}개)")

        except Exception as e:
            messagebox.showerror("오류", f"파일 로드 실패: {e}")
            self.update_status(f"파일 로드 실패: {e}")

    def filter_glossary(self):
        """용어집 KR 필터링"""
        search_term = self.glossary_search_var.get().strip()
        self.load_glossary_and_update_ui(kr_filter=search_term)

    def clear_glossary_filter(self):
        """용어집 필터 초기화"""
        self.glossary_search_var.set("")
        self.load_glossary_and_update_ui()
        
    def load_glossary_and_update_ui(self):
        """[지휘자] DB에서 용어집을 가져와 UI에 표시합니다."""
        self.glossary = self.db_manager.get_all_glossary()
        self.glossary_tree.delete(*self.glossary_tree.get_children())
        for record in self.glossary.values():
            values = []
            for col in self.glossary_cols:
                val = record.get(col)
                if col == 'verified':
                    val = "Y" if val == 1 else "N"
                values.append(val or "")
            self.glossary_tree.insert("", "end", values=values)
            
            self.update_status("용어집 전체 로드 완료")

    def load_exclusion_rules_and_update_ui(self):
        """[지휘자] DB에서 제외 규칙을 가져와 UI에 표시합니다."""
        self.exclusion_rule_tree.delete(*self.exclusion_rule_tree.get_children())
        rules = self.db_manager.get_exclusion_rules(only_enabled=False)
        for rule_id, desc, r_type, field, value, is_enabled in rules:
            enabled_text = "활성화" if is_enabled == 1 else "비활성화"
            self.exclusion_rule_tree.insert("", "end", iid=rule_id, values=(desc, r_type, field, value, enabled_text))

    def determine_status(self, kr_text):
        """TM을 기반으로 텍스트의 초기 상태를 결정합니다."""
        if kr_text in self.translation_memory:
            return "[확정]"
        return "[신규]"

    def update_translation_table(self):
        """번역 테이블 업데이트 (필터 개선)"""
        self.translation_tree.delete(*self.translation_tree.get_children())
            
        search_text = self.search_var.get().lower()
        active_filters = {status for status, var in self.filter_vars.items() if var.get()}
        selected_method = self.method_filter_var.get()
        
        # 번역 방법 필터 값 업데이트
        all_methods = sorted(list(set(t["method"] for t in self.pending_translations if t["method"])))
        self.method_filter_combo['values'] = ["전체"] + all_methods
        
        for trans in self.pending_translations:
            # 검색 필터
            if search_text and search_text not in trans["KR"].lower() and search_text not in trans["STRING_ID"].lower():
                continue
                
            # 상태 필터 (대괄호 제거하여 비교)
            status_text = trans["status"].strip("[]")
            if status_text not in active_filters:
                continue

            # 방법 필터
            if selected_method != "전체" and trans["method"] != selected_method:
                continue
                        
            # 값 리스트 생성
            values = [
                "☑",
                trans["STRING_ID"],
                trans["KR"],
                trans["status"],
            ]
            
            # 언어별 번역문 추가
            for lang in self.VISIBLE_LANGS:
                values.append(trans["translations"].get(lang, ""))
                
            values.append(trans["method"])
            
            item_id = self.translation_tree.insert("", "end", values=values)
            self.check_states[item_id] = True
            
        self.update_stats_label()

    def execute_translation(self):
        """[지휘자] 번역 실행을 번역 엔진에 위임"""
        items_to_translate = [
            trans for trans in self.pending_translations
            if self.check_states.get(self.find_item_id_by_string_id(trans["STRING_ID"]), True)
        ]
        self.translation_engine.execute_translation(items_to_translate)
        self.root.after(1000, self.validate_current_batch)

    def on_translation_complete(self, count):
        """번역 완료 후 UI 업데이트 콜백."""
        self.update_translation_table()
        self.update_status(f"번역 완료. {count}개 항목 DB 업데이트됨.")
        self.progress_bar['value'] = 100
        self.validate_current_batch()
        messagebox.showinfo("완료", f"{count}개 항목의 번역 및 저장이 완료되었습니다.")
                
    # --- 기타 유틸리티 및 이벤트 핸들러 ---    
    def update_status(self, message):
        """메인 창의 상태 메시지를 업데이트합니다."""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=message)
            self.root.update_idletasks()

    def update_progress(self, value, text=None):
        """메인 창의 진행 바와 상태 메시지를 업데이트합니다."""
        if hasattr(self, 'progress_bar'):
            self.progress_bar['value'] = value
        if text:
            self.update_status(text)

    def get_llm_prompt(self):
        """LLM 프롬프트 위젯에서 현재 텍스트를 가져옵니다."""
        if hasattr(self, 'llm_prompt_entry'):
            return self.llm_prompt_entry.get("1.0", "end-1c").strip()
        return ""

    def on_engine_changed(self):
        """번역 엔진 선택이 변경되었을 때 호출"""
        selected_engine = self.api_engine_var.get()
        engine_map = {
            "llm": "🤖 LLM 엔진 선택됨 - 우측 프롬프트 영역을 활용하세요",
            "azure": "🔶 Azure 엔진 선택됨"
        }
        self.update_status(engine_map.get(selected_engine, "알 수 없는 엔진"))

    def on_scenario_option_changed(self):
        """시나리오 번역 옵션 변경 시 호출 (프롬프트 창 업데이트 포함)"""
        if self.scenario_translation_var.get():
            # LLM 엔진 강제 선택
            self.api_engine_var.set("llm")
            self.on_engine_changed()
            
            # 시나리오용 프롬프트로 변경
            self.set_scenario_mode_prompt()
            
            self.update_status("✅ 시나리오 번역 모드 활성화 (화자별 맞춤 번역)")
            
            # 시나리오 탭으로 안내
            messagebox.showinfo("시나리오 번역 모드", 
                            "🎭 시나리오 번역이 활성화되었습니다!\n\n"
                            "✨ 번역 시 화자별 맞춤 프롬프트가 자동 생성됩니다.\n"
                            "📝 프롬프트 창은 기본 템플릿이며, 실제로는 화자별로 최적화된 프롬프트를 사용합니다.\n\n"
                            "'🎭 시나리오 번역' 탭에서 레퍼런스 데이터와 화자를 설정하세요.")
        else:
            # 일반 모드로 복원
            self.set_prompt_template("default")
            self.update_status("❌ 시나리오 번역 모드 비활성화")

    def handle_drop(self, event):
        """파일 드래그 앤 드랍 이벤트 처리"""
        try:
            file_path = event.data.strip().replace('{', '').replace('}', '')
            if file_path.lower().endswith((".xlsx", ".xls")):
                self.file_path_var.set(file_path)
                self.load_data()
            else:
                messagebox.showwarning("파일 형식 오류", "엑셀 파일만 드랍할 수 있습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{e}")


    def on_tree_click(self, event):
        """트리 클릭 이벤트 (키보드 조작과 일관성 유지)"""
        region = self.translation_tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.translation_tree.identify_column(event.x)
            item = self.translation_tree.identify_row(event.y)
            
            if column == "#1" and item:  # 선택 컬럼 클릭
                # 현재 선택된 항목들이 여러 개인 경우
                selected_items = self.translation_tree.selection()
                if len(selected_items) > 1 and item in selected_items:
                    # 여러 항목이 선택된 상태에서 체크박스 클릭하면 모든 선택된 항목 토글
                    self.toggle_multiple_items(selected_items)
                else:
                    # 단일 항목만 토글
                    self.toggle_single_item(item)
            else:
                # 체크박스가 아닌 다른 영역 클릭 시에는 해당 항목을 선택
                if item:
                    self.translation_tree.selection_set(item)
                    self.translation_tree.focus(item)

    def on_conflict_row_selected(self, event):
        """충돌 목록에서 항목 선택 시, 해결 패널에 후보들을 표시"""
        selection = self.conflict_tree.selection()
        if not selection: return
        
        kr_text = self.conflict_tree.item(selection[0], "values")[0]
        self.conflict_kr_var.set(kr_text)
        
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT conflict_info FROM translation_memory WHERE kr_text=?", (kr_text,))
        result = cursor.fetchone()
        conn.close()
        
        if not result or not result[0]: return

        conflict_info = json.loads(result[0])
        
        for lang, combo in self.conflict_combos.items():
            if lang in conflict_info:
                candidates = conflict_info[lang] # {'번역1': 3, '번역2': 1} 형태의 dict
                # 빈도순으로 정렬된 후보 목록 생성
                sorted_candidates = sorted(candidates.items(), key=lambda item: item[1], reverse=True)
                display_list = [f"{text} ({count}회)" for text, count in sorted_candidates]
                
                combo.config(values=display_list, state="readonly")
                combo.set(display_list[0]) # 가장 빈도가 높은 것을 기본값으로 설정
                combo.real_values = [text for text, count in sorted_candidates] # 실제 값 저장
            else:
                combo.config(values=[], state="disabled")
                combo.set("")

    def on_inline_edit_complete(self, updated_trans):
        """인라인 편집 완료 후 처리"""
        # TM 업데이트
        self.translation_memory[updated_trans["KR"]] = updated_trans["translations"].copy()
        
        # DB 업데이트
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT OR REPLACE INTO translation_memory 
            (kr_text, translations, source, confidence)
            VALUES (?, ?, ?, ?)
        """, (
            updated_trans["KR"],
            json.dumps(updated_trans["translations"]),
            "직접편집",
            1.0
        ))
        conn.commit()
        conn.close()
        
        # UI 업데이트
        self.update_translation_table()
        self.update_status("편집 내용이 저장되었습니다.")

    def on_speaker_saved(self, speaker):
        """화자 저장 완료 콜백 (안전한 초기화)"""
        try:
            self.ensure_scenario_manager()
            if self.scenario_manager:
                self.scenario_manager.save_speaker(speaker)
                self.refresh_speaker_list()
                self.update_status(f"화자 '{speaker.name}' 저장됨")
        except Exception as e:
            print(f"화자 저장 콜백 오류: {e}")
            messagebox.showerror("오류", f"화자 저장 중 오류: {e}")

    def select_file(self):
        """파일 선택"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def load_data(self):
        """[지휘자] 엑셀 파일 로드를 file_handler에게 지시하고, 받은 데이터로 UI를 처리합니다."""
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("경고", "파일을 선택하세요.")
            return

        try:
            self.update_status("엑셀 파일 로드 중...")

            ### 수정된 부분 ###
            df = file_handler.load_data_from_excel(file_path)

            self.translation_memory = self.db_manager.get_translation_memory()
            active_rules_tuples = self.db_manager.get_exclusion_rules(only_enabled=True)
            active_rules = [{'id': r[0], 'type': r[1], 'field': r[2], 'value': r[3]} for r in active_rules_tuples]

            self.pending_translations.clear()
            excluded_count = 0

            for _, row in df.iterrows():
                if pd.isna(row.get("KR")) or not str(row["KR"]).strip():
                    continue

                entry_dict = row.to_dict()
                if self.db_manager._is_entry_excluded(entry_dict, active_rules):
                    excluded_count += 1
                    continue

                kr_text = str(row["KR"]).strip()
                item = {
                    "STRING_ID": str(row["STRING_ID"]), "KR": kr_text,
                    "status": self.determine_status(kr_text),
                    "translations": {lang: str(row[lang]) for lang in self.VISIBLE_LANGS if lang in df.columns and pd.notna(row.get(lang))},
                    "method": ""
                }
                self.pending_translations.append(item)

            self.update_translation_table()
            self.update_status(f"로드 완료: {len(self.pending_translations)}개 항목 (제외: {excluded_count}개)")

        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("파일 로드 오류", f"{e}")
        except Exception as e:
            messagebox.showerror("오류", f"파일을 처리하는 중 예기치 않은 오류가 발생했습니다: {e}")
        finally:
            self.update_status("준비됨")

    def analyze_translations(self):
        """번역 분석을 번역 엔진에 위임"""
        self.translation_engine.analyze_translations()

    def force_retranslate_selected(self):
        """선택된 항목을 강제로 재번역 - 번역 엔진에 위임"""
        selected_items = self.translation_tree.selection()
        if not selected_items:
            messagebox.showwarning("선택 오류", "재번역할 항목을 선택하세요.")
            return
        
        if not messagebox.askyesno("재번역 확인", 
                                f"선택된 {len(selected_items)}개 항목을 API로 강제 재번역하시겠습니까?\n\n" +
                                "기존 번역 내용은 덮어쓰여지고 TM도 업데이트됩니다."):
            return
        
        # 선택된 항목들의 STRING_ID 수집
        selected_string_ids = []
        for item in selected_items:
            values = self.translation_tree.item(item, "values")
            if len(values) > 1:
                selected_string_ids.append(values[1])  # STRING_ID
        
        self.translation_engine.force_retranslate_selected(selected_string_ids)

    def check_multilang_prerequisites(self):
        """다국어 번역 선행 조건 체크 (개별 함수로 분리)"""
        do_en_trans = self.translate_en_var.get()
        do_multi_trans = self.translate_multi_var.get()
        do_cn_tw_trans = self.translate_cn_tw_var.get()
        
        # EN이 체크되지 않았지만 다국어가 체크된 경우
        if (do_multi_trans or do_cn_tw_trans) and not do_en_trans:
            return False, "EN 번역이 필요합니다"
        
        # 선택된 항목 중 EN이 비어있는 항목 체크
        empty_en_count = 0
        total_selected = 0
        
        for trans in self.pending_translations:
            item_id = self.find_item_id_by_string_id(trans["STRING_ID"])
            if self.check_states.get(item_id, True):
                total_selected += 1
                if not trans["translations"].get("EN"):
                    empty_en_count += 1
        
        if empty_en_count > 0 and (do_multi_trans or do_cn_tw_trans):
            return False, f"{empty_en_count}개 항목에 EN 번역이 없습니다"
        
        return True, "조건 충족"

        
    def save_results(self):
        """[지휘자] 번역 결과를 엑셀에 저장하도록 file_handler에게 지시합니다."""
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("경고", "원본 파일을 먼저 로드하세요.")
            return

        if not messagebox.askyesno("저장 확인", f"번역 결과를 원본 파일에 직접 덮어씁니다.\n파일: {os.path.basename(file_path)}\n\n계속하시겠습니까?"):
            return

        try:
            self.update_status("파일 저장 중...")

            ### 수정된 부분 ###
            updated_rows = file_handler.save_results_to_excel(
                file_path, 
                self.pending_translations, 
                self.VISIBLE_LANGS
            )

            self.update_status(f"결과 저장 완료! 총 {updated_rows}개 행의 번역이 적용되었습니다.")
            messagebox.showinfo("저장 완료", f"'{os.path.basename(file_path)}' 파일에 번역 내용이 성공적으로 적용되었습니다.\n(총 {updated_rows}개 행 업데이트)")

        except PermissionError as e:
            self.update_status("오류: 파일을 다른 프로그램에서 사용 중입니다.")
            messagebox.showerror("저장 실패", f"{e}")
        except Exception as e:
            self.update_status(f"결과 저장 중 오류 발생: {e}")
            messagebox.showerror("저장 오류", f"결과를 저장하는 동안 예기치 않은 오류가 발생했습니다:\n{e}")
        
    def edit_translation_inline(self):
        """선택된 항목을 직접 편집"""
        selected_items = self.translation_tree.selection()
        if not selected_items:
            messagebox.showwarning("선택 오류", "편집할 항목을 선택하세요.")
            return
        
        if len(selected_items) > 1:
            messagebox.showwarning("선택 오류", "한 번에 하나의 항목만 편집할 수 있습니다.")
            return
        
        item = selected_items[0]
        values = self.translation_tree.item(item, "values")
        string_id = values[1]
        
        # pending_translations에서 해당 항목 찾기
        trans_item = next((t for t in self.pending_translations if t["STRING_ID"] == string_id), None)
        if not trans_item:
            messagebox.showerror("오류", "편집할 데이터를 찾을 수 없습니다.")
            return
        
        # 편집 다이얼로그 표시
        InlineEditDialog(self.root, trans_item, self.VISIBLE_LANGS, self.on_inline_edit_complete)

    def remove_from_tm(self):
        """[지휘자] 선택된 항목을 TM에서 삭제하도록 DB 매니저에게 요청합니다."""
        selected_items = self.translation_tree.selection()
        if not selected_items: return messagebox.showwarning("선택 오류", "삭제할 항목을 선택하세요.")
        
        kr_texts = [self.translation_tree.item(item, "values")[2] for item in selected_items]
        
        if messagebox.askyesno("TM 삭제 확인", f"선택된 {len(kr_texts)}개 항목을 TM에서 삭제하시겠습니까?"):
            ### 수정된 부분: DB 매니저에게 삭제 요청 ###
            self.db_manager.remove_from_tm(kr_texts)
            
            # 메모리에서도 삭제
            for kr in kr_texts:
                if kr in self.translation_memory:
                    del self.translation_memory[kr]
            
            self.update_translation_table() # 테이블 새로고침
            self.update_status(f"TM에서 {len(kr_texts)}개 항목 삭제됨")

    def view_tm_entry(self):
        """선택된 항목의 TM 정보 보기"""
        selected_items = self.translation_tree.selection()
        if not selected_items:
            messagebox.showwarning("선택 오류", "확인할 항목을 선택하세요.")
            return
        
        item = selected_items[0]
        values = self.translation_tree.item(item, "values")
        kr_text = values[2]  # KR 컬럼
        
        if kr_text in self.translation_memory:
            tm_data = self.translation_memory[kr_text]
            info_text = f"KR: {kr_text}\n\nTM 저장 내용:\n"
            for lang, trans in tm_data.items():
                info_text += f"{lang}: {trans}\n"
        else:
            info_text = f"KR: {kr_text}\n\n❌ TM에 저장된 내용이 없습니다."
        
        messagebox.showinfo("TM 정보", info_text)

    def _import_gsheet_thread(self, spreadsheet_id, sheet_name):
        """(스레드) 구글 시트 API를 통해 새 구조의 용어집을 가져와 업데이트합니다."""
        try:
            self.update_status("구글 시트에서 데이터 인증 및 다운로드 중...")
            
            from google.oauth2 import service_account
            from googleapiclient.discovery import build

            SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
            SERVICE_ACCOUNT_FILE = 'dulcet-antler-462703-n8-d2fbdb362407.json' 

            creds = service_account.Credentials.from_service_account_file(
                    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
            
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            
            range_name = f"{sheet_name}!A:Z"
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            if not values or len(values) < 2:
                self.update_status("오류: 구글 시트에서 데이터를 찾을 수 없습니다.")
                return
                
            self.update_status("가져온 데이터 처리 및 DB 업데이트 중...")
            
            header = [h.lower() for h in values[0]] # 헤더는 소문자로 통일
            data_rows = values[1:]
            
            master_conn = sqlite3.connect(self.translation_db_path)
            master_cursor = master_conn.cursor()

            imported_count = 0
            for row in data_rows:
                row_data = dict(zip(header, row))
                string_id = row_data.get('string_id')

                if string_id and string_id.strip():
                    # INSERT OR REPLACE 쿼리를 위한 값 리스트 생성
                    # DB 컬럼 순서: string_id, kr, en, cn, tw, th, pt, es, de, fr, jp, engine, contributor, update_at, verified, description
                    db_values = (
                        string_id,
                        row_data.get('kr'), row_data.get('en'), row_data.get('cn'),
                        row_data.get('tw'), row_data.get('th'), row_data.get('pt'),
                        row_data.get('es'), row_data.get('de'), row_data.get('fr'),
                        row_data.get('jp'), row_data.get('engine'), row_data.get('contributor'),
                        row_data.get('update_at'), int(row_data.get('verified', 0) or 0),
                        row_data.get('description', '')
                    )
                    
                    master_cursor.execute("""
                        INSERT OR REPLACE INTO glossary 
                        (string_id, kr, en, cn, tw, th, pt, es, de, fr, jp, engine, contributor, update_at, verified, description)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, db_values)
                    imported_count += 1

            master_conn.commit()
            master_conn.close()

            self.root.after(0, self.load_glossary_and_update_ui)
            self.update_status(f"구글 시트 가져오기 완료! {imported_count}개 항목 처리됨.")
            self.root.after(0, lambda: messagebox.showinfo("완료", f"구글 시트에서 {imported_count}개의 용어를 성공적으로 가져왔습니다."))

        except Exception as e:
            self.update_status(f"구글 시트 가져오기 오류: {e}")
            self.root.after(0, lambda: messagebox.showerror("API 오류", f"구글 시트 처리 중 오류가 발생했습니다:\n{e}"))
            import traceback
            traceback.print_exc()

    def start_db_build(self):
        """DB 구축 시작 버튼의 동작. 선택된 모드에 따라 적절한 스레드를 실행."""
        folder_path = filedialog.askdirectory(title="소스 DB들이 있는 폴더를 선택하세요")
        if not folder_path:
            return
            
        mode = self.db_build_mode_var.get()
        
        if mode == 'conflict':
            target_func = self._run_conflict_build_mode
        elif mode == 'fill_blanks':
            target_func = self._run_fill_blanks_mode
        else:
            messagebox.showerror("오류", "알 수 없는 빌드 모드입니다.")
            return

        threading.Thread(target=target_func, args=(folder_path,), daemon=True).start()


    def _run_conflict_build_mode(self, folder_path):
        """(스레드) '충돌 우선 해결 모드'를 실행합니다. (진행 상황 표시 개선)"""
        try:
            self.update_status("충돌 해결 모드 시작: 소스 DB 분석 중...")
            self.progress_bar['value'] = 0

            conn = sqlite3.connect(self.translation_db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT id, rule_type, field, value FROM exclusion_rules WHERE is_enabled = 1")
            active_rules = [{'id': r[0], 'type': r[1], 'field': r[2], 'value': r[3]} for r in cursor.fetchall()]

            raw_data = defaultdict(lambda: {lang: Counter() for lang in self.VISIBLE_LANGS})
            
            source_db_files = [os.path.join(root, file)
                            for root, _, files in os.walk(folder_path)
                            for file in files
                            if file.lower().startswith('string') and file.lower().endswith('.db')]

            if not source_db_files:
                self.update_status("오류: 폴더에 'String*.db' 파일이 없습니다.")
                conn.close()
                return

            # 1단계: 소스 DB 분석 (진행률 0% -> 50%)
            total_files = len(source_db_files)
            for i, db_path in enumerate(source_db_files):
                self.update_status(f"DB 분석 중 ({i+1}/{total_files}): {os.path.basename(db_path)}")
                self.progress_bar['value'] = ((i + 1) / total_files) * 50
                # (이하 로직은 기존과 동일)
                try:
                    source_conn = sqlite3.connect(db_path)
                    source_cursor = source_conn.cursor()
                    source_cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'String%'")
                    tables = [r[0] for r in source_cursor.fetchall()]
                    for table in tables:
                        source_cursor.execute(f'SELECT * FROM "{table}"')
                        cols = [desc[0] for desc in source_cursor.description]
                        for row in source_cursor.fetchall():
                            entry = dict(zip(cols, row))
                            if self._is_entry_excluded(entry, active_rules):
                                continue
                            kr_text = str(entry.get('KR', '')).strip()
                            if not kr_text: continue
                            for lang in self.VISIBLE_LANGS:
                                if lang in entry and pd.notna(entry[lang]):
                                    raw_data[kr_text][lang][str(entry[lang])] += 1
                    source_conn.close()
                except Exception as e:
                    print(f"소스 DB 처리 오류 {db_path}: {e}")

            # 2단계: 마스터 TM에 충돌 정보 기록 (진행률 50% -> 100%)
            self.update_status("마스터 TM에 충돌 정보 기록 중...")
            processed_count = 0
            total_kr = len(raw_data)
            
            if total_kr > 0:
                for i, (kr_text, lang_map) in enumerate(raw_data.items()):
                    # <<< 시작: 이 부분이 추가되었습니다 >>>
                    self.progress_bar['value'] = 50 + (((i + 1) / total_kr) * 50)
                    # <<< 종료: 이 부분이 추가되었습니다 >>>
                
                    is_conflict = any(len(counter) > 1 for counter in lang_map.values())
                    
                    if is_conflict:
                        status = "conflict"
                        conflict_info = {lang: dict(counter) for lang, counter in lang_map.items() if len(counter) > 1}
                        translations = {lang: counter.most_common(1)[0][0] if counter else "" for lang, counter in lang_map.items()}
                    else:
                        status = "consolidated"
                        conflict_info = None
                        translations = {lang: list(counter.keys())[0] if counter else "" for lang, counter in lang_map.items()}

                    cursor.execute("""
                        INSERT OR REPLACE INTO translation_memory (kr_text, translations, source, status, conflict_info)
                        VALUES (?, ?, ?, ?, ?)
                    """, (kr_text, json.dumps(translations), "DB Build (Conflict Mode)", status, json.dumps(conflict_info) if conflict_info else None))
                    processed_count += 1
            
            conn.commit()
            conn.close()

            # 완료 후 UI 새로고침
            self.root.after(0, self.load_translation_memory)
            self.root.after(0, self.load_tm_view)
            self.root.after(0, self.load_conflicts_to_view)
            self.update_status(f"충돌 해결 모드 완료. {processed_count}개 항목 처리됨.")
            self.root.after(0, lambda: messagebox.showinfo("완료", "충돌 감지가 완료되었습니다.\n[⚠️ 충돌 해결] 탭에서 결과를 확인하고 해결해주세요."))

        except Exception as e:
            self.update_status(f"오류: {e}")
            import traceback
            traceback.print_exc()


    # _run_fill_blanks_mode 함수 코드
    def _run_fill_blanks_mode(self, folder_path):
        """(스레드) '빈칸 채우기 모드'를 실행합니다."""
        try:
            self.update_status("빈칸 채우기 모드 시작: 소스 DB 분석 중...")
            self.progress_bar['value'] = 0
            
            # 1. 소스 DB에서 모든 항목을 수집하여 KR별로 최적의 번역본 하나로 병합
            all_source_entries = self._collect_entries_from_dbs(folder_path)
            self.progress_bar['value'] = 30
            
            # 2. 현재 마스터 TM의 데이터를 메모리로 로드
            self.update_status("기존 마스터 TM 데이터 로딩 중...")
            existing_master_entries = self._load_master_tm_for_merge()
            self.progress_bar['value'] = 40
            
            # 3. 두 데이터를 '비파괴적 업데이트' 방식으로 병합하고, 신규 항목에 ID 부여
            self.update_status("데이터 병합 및 ID 부여 중...")
            merged_entries, new_count, updated_count = self._merge_tm_entries(all_source_entries, existing_master_entries)
            self.progress_bar['value'] = 70

            # 4. 최종 병합 결과를 마스터 DB에 저장
            self.update_status("최종 결과를 마스터 TM에 저장 중...")
            self._save_merged_to_master_tm(merged_entries)
            
            # 5. 작업 완료 및 UI 새로고침
            self.root.after(0, self.load_translation_memory)
            self.root.after(0, self.load_tm_view)
            self.progress_bar['value'] = 100
            self.update_status(f"빈칸 채우기 완료! 신규 {new_count}개, 업데이트 {updated_count}개.")
            self.root.after(0, lambda: messagebox.showinfo("완료", f"빈칸 채우기 작업이 완료되었습니다.\n신규: {new_count}개, 업데이트: {updated_count}개"))

        except Exception as e:
            self.update_status(f"오류: {e}")
            import traceback
            traceback.print_exc()


    def _collect_entries_from_dbs(self, folder_path):
        """(헬퍼) 여러 소스 DB에서 모든 번역 항목을 수집합니다. (동적 테이블 조회 수정)"""
        kr_entries_map = defaultdict(lambda: {lang: [] for lang in self.VISIBLE_LANGS})
        
        source_db_files = [os.path.join(root, file)
                        for root, _, files in os.walk(folder_path)
                        for file in files
                        if file.lower().startswith('string') and file.lower().endswith('.db')]
        
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, rule_type, field, value FROM exclusion_rules WHERE is_enabled = 1")
        active_rules = [{'id': r[0], 'type': r[1], 'field': r[2], 'value': r[3]} for r in cursor.fetchall()]
        conn.close()

        for db_path in source_db_files:
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # <<< 시작: 동적 테이블 조회 로직으로 수정 >>>
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'String%'")
                tables = [r[0] for r in cursor.fetchall()]

                for table in tables:
                    cursor.execute(f'SELECT * FROM "{table}"')
                    columns = [desc[0] for desc in cursor.description]
                    for row in cursor.fetchall():
                        entry = dict(zip(columns, row))
                        if self._is_entry_excluded(entry, active_rules):
                            continue
                        
                        kr_text = str(entry.get('KR', '')).strip()
                        if not kr_text: continue
                        
                        for lang in self.VISIBLE_LANGS:
                            if lang in entry and pd.notna(entry[lang]):
                                kr_entries_map[kr_text][lang].append(str(entry[lang]))
                                
                conn.close()
            except Exception as e:
                print(f"소스 DB 처리 오류 {db_path}: {e}")

        all_merged_entries = []
        for kr, lang_values in kr_entries_map.items():
            merged_entry = {"KR": kr}
            for lang, values in lang_values.items():
                merged_entry[lang] = next((v for v in values if v.strip()), "")
            all_merged_entries.append(merged_entry)
            
        return all_merged_entries


    def _load_master_tm_for_merge(self):
        """(헬퍼) 병합을 위해 기존 마스터 TM을 KR을 키로 하는 딕셔너리로 로드합니다."""
        existing_entries = {}
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT kr_text, translations FROM translation_memory")
        for kr, trans_json in cursor.fetchall():
            entry = json.loads(trans_json)
            entry['KR'] = kr # KR 값 추가
            existing_entries[kr] = entry
        conn.close()
        return existing_entries

    def _merge_tm_entries(self, source_entries, master_entries):
        """(헬퍼) 소스 항목과 마스터 항목을 '비파괴적 업데이트' 방식으로 병합하고 ID를 부여합니다."""
        merged = master_entries.copy()
        new_count = 0
        updated_count = 0
        
        # 다음 ID 계산
        existing_ids = [v.get("STRING_ID", "") for v in master_entries.values()]
        nums = [int(i.replace("utext_", "")) for i in existing_ids if i and i.startswith("utext_")]
        next_num = max(nums + [0]) + 1
        
        for entry in source_entries:
            kr = entry["KR"]
            if kr in merged: # 이미 마스터에 있는 경우
                is_updated = False
                for lang in self.VISIBLE_LANGS:
                    # 마스터의 해당 언어가 비어있고, 소스에 값이 있는 경우에만 채워넣기
                    if not merged[kr].get(lang) and entry.get(lang):
                        merged[kr][lang] = entry[lang]
                        is_updated = True
                if is_updated:
                    updated_count += 1
            else: # 마스터에 없는 신규 항목인 경우
                entry["STRING_ID"] = f"utext_{next_num:05}"
                merged[kr] = entry
                next_num += 1
                new_count += 1
                
        return merged, new_count, updated_count

    def _save_merged_to_master_tm(self, merged_entries):
        """(헬퍼) 병합된 최종 결과를 DB에 다시 씁니다."""
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        
        # INSERT OR REPLACE를 사용하여 효율적으로 저장
        for kr, entry in merged_entries.items():
            # translations 컬럼에 넣을 JSON 데이터 준비
            translations_dict = {lang: entry.get(lang, "") for lang in self.VISIBLE_LANGS}
            translations_dict['STRING_ID'] = entry.get('STRING_ID', '')
            
            cursor.execute("""
                INSERT OR REPLACE INTO translation_memory (kr_text, translations, source, status)
                VALUES (?, ?, ?, ?)
            """, (kr, json.dumps(translations_dict), 'DB Merge', 'consolidated'))
            
        conn.commit()
        conn.close()         


    def search_excel_for_import(self):
        """'파일 검색 실행' 버튼의 동작. 지정된 폴더에서 엑셀 파일을 찾아 체크리스트에 표시."""
        folder = self.excel_import_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "유효한 폴더를 선택하세요.")
            return
        
        self.excel_files_checklist.clear()
        self.excel_import_files.clear()
        
        found_files = []
        for root, _, files in os.walk(folder):
            for file in files:
                file_name_no_ext, ext = os.path.splitext(file)
                if ext.lower() == ".xlsx" and file_name_no_ext.lower().startswith("string"):
                    file_path = os.path.join(root, file)
                    self.excel_import_files.append((file, file_path))
                    found_files.append(file)

        if not found_files:
            messagebox.showinfo("알림", "조건에 맞는 엑셀 파일을 찾지 못했습니다.")
        else:
            for file_name in sorted(found_files):
                self.excel_files_checklist.add_item(file_name, checked=True)
            messagebox.showinfo("알림", f"{len(found_files)}개의 엑셀 파일을 찾았습니다. 목록에서 처리할 파일을 선택하세요.")

    def start_excel_import(self):
        """'업데이트 시작' 버튼의 동작. 선택된 파일과 언어로 TM 업데이트 스레드 실행."""
        checked_files_names = self.excel_files_checklist.get_checked_items()
        if not checked_files_names:
            messagebox.showwarning("선택 오류", "하나 이상의 엑셀 파일을 선택하세요.")
            return
            
        selected_langs = [lang for lang, var in self.excel_import_lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("선택 오류", "하나 이상의 언어를 선택하세요.")
            return
            
        # (파일명, 파일경로) 튜플에서 선택된 파일명의 경로만 추출
        files_to_process = [path for name, path in self.excel_import_files if name in checked_files_names]
        
        # 실제 DB 업데이트 작업을 별도 스레드에서 실행
        threading.Thread(target=self._excel_import_thread, args=(files_to_process, selected_langs), daemon=True).start()

    def _excel_import_thread(self, files_to_process, selected_langs):
        """(스레드) 선택된 엑셀 파일을 읽어 마스터 TM을 업데이트합니다."""
        try:
            self.update_status("TM 업데이트 시작...")
            self.progress_bar['value'] = 0

            # 1. 활성화된 제외 규칙 미리 로드
            conn = sqlite3.connect(self.translation_db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT id, rule_type, field, value FROM exclusion_rules WHERE is_enabled = 1")
            active_rules = [{'id': r[0], 'type': r[1], 'field': r[2], 'value': r[3]} for r in cursor.fetchall()]

            total_files = len(files_to_process)
            processed_rows = 0

            # 2. 선택된 각 엑셀 파일을 순회하며 처리
            for i, file_path in enumerate(files_to_process):
                self.update_status(f"파일 처리 중 ({i+1}/{total_files}): {os.path.basename(file_path)}")
                self.progress_bar['value'] = (i / total_files) * 100
                
                try:
                    xls = pd.ExcelFile(file_path)
                    for sheet_name in xls.sheet_names:
                        if not sheet_name.lower().startswith('string'):
                            continue # 'string'으로 시작하는 시트만 처리
                        
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        if df.empty: continue
                        
                        # 3. 각 행을 순회하며 마스터 TM에 업데이트
                        for _, row in df.iterrows():
                            # 제외 규칙 적용
                            entry_dict = row.to_dict()
                            if self._is_entry_excluded(entry_dict, active_rules):
                                continue
                            
                            kr_text = str(row.get('KR', '')).strip()
                            if not kr_text: continue

                            # 4. 'Select -> Modify -> Replace' 전략으로 안전하게 업데이트
                            cursor.execute("SELECT translations FROM translation_memory WHERE kr_text=?", (kr_text,))
                            master_result = cursor.fetchone()

                            if master_result:
                                # 기존 항목: 선택된 언어만 업데이트
                                current_translations = json.loads(master_result[0])
                                for lang in selected_langs:
                                    if lang in row:
                                        current_translations[lang] = str(row[lang]) if pd.notna(row[lang]) else ""
                            else:
                                # 신규 항목: 모든 언어 정보로 새로 생성
                                current_translations = {}
                                for lang in self.VISIBLE_LANGS: # 모든 지원 언어에 대해
                                    if lang in row:
                                        current_translations[lang] = str(row[lang]) if pd.notna(row[lang]) else ""
                            
                            # DB에 최종본을 덮어쓰기 (없으면 새로 추가됨)
                            cursor.execute("""
                                INSERT OR REPLACE INTO translation_memory (kr_text, translations, source)
                                VALUES (?, ?, ?)
                            """, (kr_text, json.dumps(current_translations), "Excel Update"))
                            processed_rows += 1

                except Exception as e:
                    print(f"엑셀 파일 처리 오류 {file_path}: {e}")
            
            conn.commit()
            conn.close()

            # 5. 작업 완료 후 UI 새로고침
            self.root.after(0, self.load_translation_memory)
            self.root.after(0, self.load_tm_view)
            self.progress_bar['value'] = 100
            self.update_status(f"TM 업데이트 완료! {processed_rows}개 행이 처리되었습니다.")
            self.root.after(0, lambda: messagebox.showinfo("완료", f"선택된 엑셀 파일의 내용으로 마스터 TM이 업데이트되었습니다.\n(총 {processed_rows}개 행 처리)"))

        except Exception as e:
            self.update_status(f"TM 업데이트 중 오류 발생: {e}")

    def update_tm_from_excel(self):
        """'최신 번역(Excel)으로 업데이트' 버튼의 시작점"""
        file_path = filedialog.askopenfilename(
            title="업데이트할 엑셀 파일을 선택하세요",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        dialog = LanguageSelectionDialog(self.root, self.VISIBLE_LANGS)
        self.root.wait_window(dialog.top)
        
        selected_langs = dialog.selected_langs
        
        if selected_langs:
            # 실제 비교 및 업데이트 작업을 별도 스레드에서 실행
            threading.Thread(target=self._update_tm_thread, args=(file_path, selected_langs), daemon=True).start()
        else:
            self.update_status("언어가 선택되지 않아 업데이트를 취소했습니다.")

    def _update_tm_thread(self, excel_path, selected_langs):
        """(스레드) 엑셀->스테이징DB->마스터TM 비교 및 미리보기 표시"""
        try:
            self.update_status("엑셀 파일 읽는 중...")
            # 1. 엑셀을 읽어 스테이징 DB (메모리상) 생성
            df = pd.read_excel(excel_path)
            
            staging_conn = sqlite3.connect(':memory:') # 메모리에 임시 DB 생성
            df.to_sql('staging_tm', staging_conn, if_exists='replace', index=False)
            staging_cursor = staging_conn.cursor()

            self.update_status("마스터 TM과 비교 중...")
            # 2. 마스터 TM과 스테이징 DB 비교
            master_conn = sqlite3.connect(self.translation_db_path)
            master_cursor = master_conn.cursor()

            new_entries = []
            updated_entries = []

            staging_cursor.execute("SELECT * FROM staging_tm")
            for row in staging_cursor.fetchall():
                # staging DB의 컬럼명과 값을 딕셔너리로 매핑
                staging_entry = dict(zip([desc[0] for desc in staging_cursor.description], row))
                kr_text = staging_entry.get('KR')
                if not kr_text: continue

                master_cursor.execute("SELECT translations FROM translation_memory WHERE kr_text=?", (kr_text,))
                master_result = master_cursor.fetchone()

                if not master_result:
                    # 마스터 TM에 없는 경우 -> 신규 항목
                    new_entries.append(staging_entry)
                else:
                    # 마스터 TM에 있는 경우 -> 변경점 확인
                    master_translations = json.loads(master_result[0])
                    changes = {}
                    for lang in selected_langs:
                        master_val = master_translations.get(lang, "")
                        staging_val = str(staging_entry.get(lang, ""))
                        if master_val != staging_val:
                            changes[lang] = {'old': master_val, 'new': staging_val}
                    
                    if changes:
                        updated_entries.append({'kr': kr_text, 'changes': changes})
            
            staging_conn.close()
            master_conn.close()
            
            if not new_entries and not updated_entries:
                self.update_status("마스터 TM과 비교하여 변경된 내용이 없습니다.")
                self.root.after(0, lambda: messagebox.showinfo("업데이트 정보", "새롭거나 변경된 번역 내용이 없습니다."))
                return

            # 3. 비교 결과를 미리보기 창에 표시 (UI 작업은 메인 스레드에서)
            self.root.after(0, self.show_update_preview, new_entries, updated_entries, selected_langs)

        except Exception as e:
            self.update_status(f"업데이트 중 오류 발생: {e}")
            print(f"업데이트 스레드 오류: {e}")

    def show_update_preview(self, new_entries, updated_entries, selected_langs):
        """비교 결과를 담은 미리보기 다이얼로그를 표시"""
        dialog = UpdatePreviewDialog(self.root, new_entries, updated_entries, self.VISIBLE_LANGS)
        self.root.wait_window(dialog.top)

        if dialog.confirmed:
            # 사용자가 '적용'을 눌렀으면 실제 DB 업데이트 실행
            self.apply_tm_updates(dialog.new_to_apply, dialog.updates_to_apply, selected_langs)

    def apply_tm_updates(self, new_entries, updated_entries, selected_langs):
        """미리보기에서 확정된 내용을 실제 마스터 TM DB에 반영"""
        try:
            self.update_status("마스터 TM 업데이트 중...")
            conn = sqlite3.connect(self.translation_db_path)
            cursor = conn.cursor()

            # 신규 항목 추가
            for entry in new_entries:
                kr_text = entry.get('KR')
                translations = {lang: str(entry.get(lang, '')) for lang in self.VISIBLE_LANGS}
                cursor.execute("""
                    INSERT OR REPLACE INTO translation_memory (kr_text, translations, source) 
                    VALUES (?, ?, ?)
                """, (kr_text, json.dumps(translations), "Excel Update"))

            # 기존 항목 업데이트
            for entry in updated_entries:
                kr_text = entry['kr']
                # 먼저 현재 DB의 번역문을 불러온다
                cursor.execute("SELECT translations FROM translation_memory WHERE kr_text=?", (kr_text,))
                current_translations = json.loads(cursor.fetchone()[0])
                # 변경점만 업데이트한다
                for lang, change in entry['changes'].items():
                    current_translations[lang] = change['new']
                
                cursor.execute("UPDATE translation_memory SET translations=? WHERE kr_text=?", 
                            (json.dumps(current_translations), kr_text))

            conn.commit()
            conn.close()
            
            # 메모리와 UI 새로고침
            self.load_translation_memory()
            self.load_tm_view()
            self.update_status("TM 업데이트 완료!")
            messagebox.showinfo("완료", "최신 번역 내용이 마스터 TM에 성공적으로 반영되었습니다.")

        except Exception as e:
            self.update_status(f"TM 반영 중 오류: {e}")
            print(f"TM 반영 오류: {e}")

    def import_tm_from_folder(self):
        """'폴더에서 TM 가져오기' 버튼 클릭 시 실행될 함수"""
        folder_path = filedialog.askdirectory(title="번역 파일들이 있는 폴더(StringDB)를 선택하세요")
        if not folder_path:
            return
        
        # UI가 멈추지 않도록 별도 스레드에서 임포트 작업 실행
        threading.Thread(target=self._import_tm_thread, args=(folder_path,), daemon=True).start()

    def filter_translations(self):
        """번역 필터링"""
        self.update_translation_table()

    def update_stats_label(self):
        """통계 레이블 업데이트"""
        total = len(self.pending_translations)
        신규 = sum(1 for t in self.pending_translations if "[신규]" in t["status"])
        변경 = sum(1 for t in self.pending_translations if "[변경]" in t["status"])
        확인필요 = sum(1 for t in self.pending_translations if "[확인필요]" in t["status"])
        
        stats_text = f"전체: {total} | 신규: {신규} | 변경: {변경} | 확인필요: {확인필요}"
        
        if hasattr(self, 'stats_label'):
            self.stats_label.config(text=stats_text)
        else:
            self.stats_label = ttk.Label(self.stats_frame, text=stats_text)
            self.stats_label.pack()

    def find_item_id_by_string_id(self, string_id):
        """STRING_ID로 Treeview의 item id를 찾습니다. (오류 처리 강화)"""
        try:
            for item_id in self.translation_tree.get_children():
                values = self.translation_tree.item(item_id, "values")
                if len(values) > 1 and values[1] == string_id:
                    return item_id
        except Exception as e:
            print(f"find_item_id_by_string_id 오류: {e}")
        return None

    def toggle_all_selections(self):
        """'전체 선택/해제' 체크박스의 상태에 따라 모든 항목의 체크 상태를 변경합니다."""
        is_checked = self.select_all_var.get()
        new_symbol = "☑" if is_checked else "☐"

        for item_id in self.translation_tree.get_children():
            self.check_states[item_id] = is_checked
            
            current_values = list(self.translation_tree.item(item_id, "values"))
            if len(current_values) > 0:
                current_values[0] = new_symbol
                self.translation_tree.item(item_id, values=current_values)
        
        action = "선택" if is_checked else "해제"
        total_count = len(self.translation_tree.get_children())
        self.update_status(f"전체 {action}: {total_count}개 항목")

    def toggle_selected_checkboxes(self, event):
        """선택된 항목들의 체크박스 토글 (Spacebar/Enter 이벤트)"""
        selected_items = self.translation_tree.selection()
        
        if not selected_items:
            # 아무것도 선택되지 않았으면 현재 포커스된 항목을 토글
            focused_item = self.translation_tree.focus()
            if focused_item:
                selected_items = [focused_item]
            else:
                return "break"  # 이벤트 전파 중단
        
        # 선택된 항목들의 현재 체크 상태 확인
        current_states = [self.check_states.get(item, True) for item in selected_items]
        
        # 토글 방식 결정: 하나라도 체크되어 있으면 모두 해제, 모두 해제되어 있으면 모두 체크
        if any(current_states):
            new_state = False  # 하나라도 체크되어 있으면 모두 해제
            symbol = "☐"
        else:
            new_state = True   # 모두 해제되어 있으면 모두 체크
            symbol = "☑"
        
        # 선택된 모든 항목의 체크 상태 변경
        for item in selected_items:
            self.check_states[item] = new_state
            
            # UI에서 체크박스 심볼 업데이트
            current_values = list(self.translation_tree.item(item, "values"))
            if len(current_values) > 0:
                current_values[0] = symbol
                self.translation_tree.item(item, values=current_values)
        
        # 상태 메시지 업데이트
        action = "선택" if new_state else "해제"
        self.update_status(f"{len(selected_items)}개 항목 {action}됨")
        
        # 전체 선택/해제 체크박스 상태도 업데이트
        self.update_select_all_checkbox()
        
        return "break"

    def toggle_single_item(self, item):
        """단일 항목의 체크 상태 토글"""
        current_state = self.check_states.get(item, True)
        new_state = not current_state
        self.check_states[item] = new_state
        
        values = list(self.translation_tree.item(item, "values"))
        values[0] = "☑" if new_state else "☐"
        self.translation_tree.item(item, values=values)
        
        self.update_select_all_checkbox()

    def toggle_multiple_items(self, selected_items):
        """여러 선택된 항목들의 체크 상태를 일괄 토글"""
        # 선택된 항목들의 현재 상태 확인
        current_states = [self.check_states.get(item, True) for item in selected_items]
        
        # 하나라도 체크되어 있으면 모두 해제, 모두 해제되어 있으면 모두 체크
        if any(current_states):
            new_state = False
            symbol = "☐"
        else:
            new_state = True
            symbol = "☑"
        
        # 모든 선택된 항목 업데이트
        for item in selected_items:
            self.check_states[item] = new_state
            values = list(self.translation_tree.item(item, "values"))
            values[0] = symbol
            self.translation_tree.item(item, values=values)
        
        self.update_select_all_checkbox()

    def select_all_items(self, event):
        """Ctrl+A: 모든 항목 선택"""
        all_items = self.translation_tree.get_children()
        
        # 모든 항목 체크
        for item in all_items:
            self.check_states[item] = True
            current_values = list(self.translation_tree.item(item, "values"))
            if len(current_values) > 0:
                current_values[0] = "☑"
                self.translation_tree.item(item, values=current_values)
        
        # 전체 선택/해제 체크박스도 업데이트
        self.select_all_var.set(True)
        
        self.update_status(f"모든 항목 선택됨 ({len(all_items)}개)")
        return "break"

    def deselect_all_items(self, event):
        """Ctrl+D: 모든 항목 선택 해제"""
        all_items = self.translation_tree.get_children()
        
        # 모든 항목 해제
        for item in all_items:
            self.check_states[item] = False
            current_values = list(self.translation_tree.item(item, "values"))
            if len(current_values) > 0:
                current_values[0] = "☐"
                self.translation_tree.item(item, values=current_values)
        
        # 전체 선택/해제 체크박스도 업데이트
        self.select_all_var.set(False)
        
        self.update_status(f"모든 항목 선택 해제됨 ({len(all_items)}개)")
        return "break"

    def update_select_all_checkbox(self):
        """전체 선택/해제 체크박스 상태를 현재 선택 상태에 맞게 업데이트"""
        all_items = self.translation_tree.get_children()
        if not all_items:
            return
        
        # 모든 항목의 체크 상태 확인
        checked_count = sum(1 for item in all_items if self.check_states.get(item, True))
        
        if checked_count == len(all_items):
            self.select_all_var.set(True)   # 모두 체크됨
        elif checked_count == 0:
            self.select_all_var.set(False)  # 모두 해제됨
        else:
            # 일부만 체크된 경우는 현재 상태 유지 (또는 특별한 상태 표시 가능)
            pass

    def invert_selection(self, event):
        """Ctrl+I: 선택 상태 반전"""
        all_items = self.translation_tree.get_children()
        
        for item in all_items:
            current_state = self.check_states.get(item, True)
            new_state = not current_state
            self.check_states[item] = new_state
            
            symbol = "☑" if new_state else "☐"
            current_values = list(self.translation_tree.item(item, "values"))
            if len(current_values) > 0:
                current_values[0] = symbol
                self.translation_tree.item(item, values=current_values)
        
        self.update_select_all_checkbox()
        self.update_status("선택 상태가 반전되었습니다.")
        return "break"

    def clear_selected_translations(self, event):
        """Delete: 선택된 항목들의 번역 내용 삭제"""
        selected_items = self.translation_tree.selection()
        if not selected_items:
            return "break"
        
        if not messagebox.askyesno("번역 삭제 확인", 
                                f"선택된 {len(selected_items)}개 항목의 번역 내용을 삭제하시겠습니까?"):
            return "break"
        
        # 선택된 항목들의 번역 내용 삭제
        for item in selected_items:
            values = self.translation_tree.item(item, "values")
            string_id = values[1]
            
            # pending_translations에서 해당 항목 찾아서 번역 내용 삭제
            for trans in self.pending_translations:
                if trans["STRING_ID"] == string_id:
                    trans["translations"].clear()
                    trans["method"] = ""
                    trans["status"] = "[신규]"
                    break
        
        self.update_translation_table()
        self.update_status(f"{len(selected_items)}개 항목의 번역 내용이 삭제되었습니다.")
        return "break"

    def edit_selected_item(self, event):
        """F2: 선택된 항목 편집"""
        selected_items = self.translation_tree.selection()
        if len(selected_items) == 1:
            self.edit_translation_inline()
        return "break"

    def show_context_menu(self, event):
        """우클릭 컨텍스트 메뉴 표시"""
        # 클릭한 항목 선택
        item = self.translation_tree.identify_row(event.y)
        if item:
            self.translation_tree.selection_set(item)
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def set_prompt_template(self, template_type):
        """프롬프트 템플릿 설정 (GPT-4o-mini 최적화 및 자연스러움 강화)"""
        templates = {
            "game": """You are a professional game localizer. Translate Korean to natural English for English-speaking gamers aged 10-20.

    RULES:
    • Natural flow > literal accuracy
    • Use standard gaming terminology
    • Keep it concise and impactful
    • Preserve special tags: {}, [#color#], etc.
    • Make it feel native, not translated

    Examples:
    KR: "적을 물리쳤습니다!" → EN: "Enemy defeated!"
    KR: "아이템을 획득했습니다." → EN: "Item acquired!"

    Text to translate:""",
            
            "natural": """Translate Korean to natural, fluent English that sounds like it was originally written in English.

    FOCUS ON:
    • Natural word choice and flow
    • Meaning and nuance over literal translation
    • What native speakers would actually say
    • Appropriate formality level

    Think: "How would an English speaker express this idea?"

    Text:""",
            
            "technical": """Translate Korean technical content to precise, professional English.

    REQUIREMENTS:
    • Use standard technical terminology
    • Maintain accuracy and clarity
    • Professional tone
    • Preserve technical concepts exactly

    Text:""",
            
            "casual": """Translate Korean to casual, conversational English.

    STYLE:
    • Relaxed, friendly tone
    • Natural contractions (it's, don't, etc.)
    • Everyday vocabulary
    • How friends would talk

    Text:""",
            
            "story": """Translate Korean narrative text to engaging English.

    GOALS:
    • Maintain storytelling flow
    • Preserve emotional impact
    • Natural dialogue and descriptions
    • Immersive reading experience

    Text:""",
            
            "default": """Translate Korean to natural English for young English speakers.

    KEY POINTS:
    • Sound natural, not translated
    • Age-appropriate language (teens/young adults)
    • Preserve special formatting
    • Clear and engaging

    Text:"""
        }
        
        prompt = templates.get(template_type, templates["default"])
        
        # 텍스트 위젯 내용 교체
        self.llm_prompt_entry.delete("1.0", "end")
        self.llm_prompt_entry.insert("1.0", prompt)
        
        self.update_status(f"🎯 {template_type.title()} 프롬프트 템플릿 적용됨 (GPT-4o-mini 최적화)")

    def set_scenario_mode_prompt(self):
        """시나리오 모드용 프롬프트 설정"""
        scenario_prompt = """🎭 시나리오 번역 모드 활성화됨

    이 프롬프트는 기본 템플릿입니다.
    실제 번역 시에는 각 화자별로 다음과 같이 최적화된 프롬프트가 자동 생성됩니다:

    1️⃣ 화자 프로필 정보 (성별, 말투, 성격)
    2️⃣ 기존 번역 예시 (일관성 유지)
    3️⃣ 문맥 기반 스타일 지침
    4️⃣ 자연스러운 번역 규칙

    💡 화자가 등록되지 않은 경우에만 이 기본 프롬프트를 사용합니다.

    === 기본 시나리오 번역 프롬프트 ===

    You are localizing game dialogue. Translate Korean to natural English maintaining character consistency.

    NATURAL TRANSLATION RULES:
    • Preserve Korean specificity: "that kid" not "the kid" when 그 아이
    • Time expressions: "a little while ago" for 조금 전에  
    • Emotional nuance: keep hesitation, uncertainty, excitement
    • Character voice: maintain personality in speech patterns

    PRESERVE:
    • Special tags: {}, [#color#], etc.
    • Character personality and speaking style
    • Emotional undertones and atmosphere

    Text to translate:"""
        
        # 프롬프트 창 업데이트
        self.llm_prompt_entry.delete("1.0", "end")
        self.llm_prompt_entry.insert("1.0", scenario_prompt)

    def ensure_scenario_manager(self):
        """시나리오 매니저가 초기화되어 있는지 확인하고, 없으면 생성"""
        try:
            # scenario_manager 속성이 없으면 먼저 None으로 초기화
            if not hasattr(self, 'scenario_manager'):
                self.scenario_manager = None
                
            if not self.scenario_manager:
                try:
                    from scenario_manager import ScenarioTranslationManager
                    self.scenario_manager = ScenarioTranslationManager(self.translation_db_path)
                    print("✅ 시나리오 매니저 초기화 완료")
                except Exception as e:
                    print(f"❌ 시나리오 매니저 초기화 실패: {e}")
                    self.scenario_manager = None
                    
        except Exception as e:
            print(f"ensure_scenario_manager 오류: {e}")
            if not hasattr(self, 'scenario_manager'):
                self.scenario_manager = None

    def refresh_speaker_list(self):
        """화자 리스트 새로고침 (안전한 초기화 포함)"""
        try:
            # 시나리오 매니저 확인
            self.ensure_scenario_manager()
            
            if not self.scenario_manager:
                if hasattr(self, 'speaker_status_label'):
                    self.speaker_status_label.config(text="❌ 시나리오 매니저 초기화 실패", foreground="red")
                return
            
            # 기존 항목 삭제
            if hasattr(self, 'speaker_tree') and self.speaker_tree.winfo_exists():
                for item in self.speaker_tree.get_children():
                    self.speaker_tree.delete(item)
            
            # DB에서 화자 정보 다시 로드
            self.scenario_manager.load_speakers()
            
            # 화자 리스트 업데이트
            speaker_count = 0
            if self.scenario_manager.speakers and hasattr(self, 'speaker_tree'):
                for speaker in self.scenario_manager.speakers.values():
                    self.speaker_tree.insert("", "end", values=(
                        speaker.name,
                        speaker.gender,
                        speaker.tone,
                        speaker.style[:30] + "..." if len(speaker.style) > 30 else speaker.style,
                        speaker.reference_count
                    ))
                    speaker_count += 1
            
            # 상태 라벨 업데이트
            if hasattr(self, 'speaker_status_label'):
                if speaker_count > 0:
                    self.speaker_status_label.config(
                        text=f"✅ {speaker_count}명의 화자 로드 완료", 
                        foreground="green"
                    )
                else:
                    self.speaker_status_label.config(
                        text="ℹ️ 저장된 화자가 없습니다. 레퍼런스 분석을 먼저 진행하세요.", 
                        foreground="orange"
                    )
            
            print(f"화자 리스트 새로고침 완료: {speaker_count}명")
            
        except Exception as e:
            print(f"화자 리스트 새로고침 오류: {e}")
            if hasattr(self, 'speaker_status_label'):
                self.speaker_status_label.config(text=f"❌ 오류: {str(e)[:50]}...", foreground="red")

    def update_speaker_list(self):
        """화자 목록 업데이트 (안전한 초기화 포함)"""
        try:
            if not hasattr(self, 'scenario_manager') or not self.scenario_manager:
                self.ensure_scenario_manager()
                
            if not self.scenario_manager:
                return
                
            # 기존 항목 삭제
            if hasattr(self, 'speaker_tree') and self.speaker_tree.winfo_exists():
                for item in self.speaker_tree.get_children():
                    self.speaker_tree.delete(item)
                
                # 화자들 추가
                speaker_count = 0
                for speaker in self.scenario_manager.speakers.values():
                    self.speaker_tree.insert("", "end", values=(
                        speaker.name,
                        speaker.gender,
                        speaker.tone,
                        speaker.style[:30] + "..." if len(speaker.style) > 30 else speaker.style,
                        speaker.reference_count
                    ))
                    speaker_count += 1
                
                # 상태 업데이트
                if hasattr(self, 'speaker_status_label'):
                    self.speaker_status_label.config(
                        text=f"✅ {speaker_count}명의 화자 정보", 
                        foreground="green"
                    )
                    
        except Exception as e:
            print(f"화자 목록 업데이트 오류: {e}")

    def add_speaker(self):
        """새 화자 추가 (안전한 초기화)"""
        try:
            self.ensure_scenario_manager()
            if not self.scenario_manager:
                messagebox.showerror("오류", "시나리오 매니저가 초기화되지 않았습니다.")
                return
                
            from dialogs.speaker_dialog import SpeakerEditDialog
            SpeakerEditDialog(self.root, None, self.on_speaker_saved)
        except Exception as e:
            messagebox.showerror("오류", f"화자 추가 중 오류: {e}")

    def edit_speaker(self):
        """선택된 화자 편집 (안전한 초기화)"""
        try:
            if not hasattr(self, 'speaker_tree'):
                return
                
            selected = self.speaker_tree.selection()
            if not selected:
                messagebox.showwarning("선택 없음", "편집할 화자를 선택하세요.")
                return
            
            self.ensure_scenario_manager()
            if not self.scenario_manager:
                messagebox.showerror("오류", "시나리오 매니저가 초기화되지 않았습니다.")
                return
            
            speaker_name = self.speaker_tree.item(selected[0])['values'][0]
            if speaker_name in self.scenario_manager.speakers:
                speaker = self.scenario_manager.speakers[speaker_name]
                from dialogs.speaker_dialog import SpeakerEditDialog
                SpeakerEditDialog(self.root, speaker, self.on_speaker_saved)
        except Exception as e:
            messagebox.showerror("오류", f"화자 편집 중 오류: {e}")

    def delete_speaker(self):
        """선택된 화자 삭제 (안전한 초기화)"""
        try:
            if not hasattr(self, 'speaker_tree'):
                return
                
            selected = self.speaker_tree.selection()
            if not selected:
                messagebox.showwarning("선택 없음", "삭제할 화자를 선택하세요.")
                return
            
            self.ensure_scenario_manager()
            if not self.scenario_manager:
                messagebox.showerror("오류", "시나리오 매니저가 초기화되지 않았습니다.")
                return
            
            speaker_name = self.speaker_tree.item(selected[0])['values'][0]
            if messagebox.askyesno("삭제 확인", f"'{speaker_name}' 화자를 삭제하시겠습니까?"):
                # DB에서 삭제
                import sqlite3
                conn = sqlite3.connect(self.translation_db_path)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM speakers WHERE name = ?", (speaker_name,))
                conn.commit()
                conn.close()
                
                # 메모리에서 삭제
                if speaker_name in self.scenario_manager.speakers:
                    del self.scenario_manager.speakers[speaker_name]
                
                self.refresh_speaker_list()
                self.update_status(f"화자 '{speaker_name}' 삭제됨")
        except Exception as e:
            messagebox.showerror("오류", f"화자 삭제 중 오류: {e}")

    def select_reference_file(self):
        """레퍼런스 엑셀 파일 선택 (간소화)"""
        file_path = filedialog.askopenfilename(
            title="레퍼런스 데이터 파일 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.ref_file_var.set(file_path)
            self.update_status(f"📁 레퍼런스 파일: {os.path.basename(file_path)}")

    def load_reference_from_gsheet(self):
        """구글 시트에서 시나리오 레퍼런스 데이터 로드"""
        try:
            # 1. 구글 시트 URL/ID 입력받기
            sheet_url_or_id = simpledialog.askstring(
                "구글 시트 정보 입력", 
                "시나리오 레퍼런스 구글 시트의 전체 URL 또는 스프레드시트 ID를 입력하세요:",
                parent=self.root
            )
            if not sheet_url_or_id:
                return

            # URL에서 스프레드시트 ID 추출
            match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_url_or_id)
            spreadsheet_id = match.group(1) if match else sheet_url_or_id

            # 2. 시트(탭) 이름 입력받기
            sheet_name = simpledialog.askstring(
                "시트 이름 입력", 
                "레퍼런스 데이터가 있는 시트(탭)의 정확한 이름을 입력하세요 (예: Sheet1):",
                parent=self.root
            )
            if not sheet_name:
                return

            # 3. 분석할 언어 미리 선택받기
            available_langs = ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']
            dialog = LanguageSelectionDialog(self.root, available_langs, "분석할 언어 선택")
            self.root.wait_window(dialog.top)
            
            target_language = dialog.selected_lang if hasattr(dialog, 'selected_lang') else None
            if not target_language:
                messagebox.showwarning("선택 취소", "언어를 선택하지 않아 작업을 취소합니다.")
                return

            # 4. 확인 메시지
            if not messagebox.askyesno(
                "구글 시트 로드 확인", 
                f"다음 설정으로 구글 시트에서 레퍼런스 데이터를 가져와 분석합니다:\n\n"
                f"• 스프레드시트 ID: {spreadsheet_id}\n"
                f"• 시트 이름: {sheet_name}\n"
                f"• 분석 언어: {target_language}\n\n"
                f"계속하시겠습니까?"
            ):
                return

            # 5. 별도 스레드에서 구글 시트 로드 및 분석 실행
            threading.Thread(
                target=self._load_reference_gsheet_thread, 
                args=(spreadsheet_id, sheet_name, target_language), 
                daemon=True
            ).start()

        except Exception as e:
            messagebox.showerror("오류", f"구글 시트 로드 중 오류 발생: {e}")
            import traceback
            traceback.print_exc()

    def _load_reference_gsheet_thread(self, spreadsheet_id, sheet_name, target_language):
        """(스레드) 구글 시트에서 시나리오 레퍼런스 데이터 로드 및 분석"""
        try:
            self.update_status("구글 시트에서 레퍼런스 데이터 가져오는 중...")
            
            # 1. 구글 시트 API 인증 및 데이터 가져오기
            from google.oauth2 import service_account
            from googleapiclient.discovery import build

            creds = service_account.Credentials.from_service_account_file(
                'credentials.json', 
                scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
            )
            service = build('sheets', 'v4', credentials=creds)
            
            # 전체 시트 데이터 가져오기
            range_name = sheet_name
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, 
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            if not values or len(values) < 2:
                self.update_status("오류: 구글 시트에서 데이터를 찾을 수 없습니다.")
                self.root.after(0, lambda: messagebox.showerror(
                    "데이터 없음", 
                    "구글 시트에 데이터가 없거나 형식이 올바르지 않습니다."
                ))
                return

            self.update_status("구글 시트 데이터 구조 분석 중...")
            
            # 2. 헤더 위치 자동 감지 (3~6행 범위에서 찾기)
            best_skiprows = None
            best_score = 0
            best_header = None
            
            for skiprows in range(min(6, len(values))):
                if skiprows >= len(values):
                    continue
                    
                try:
                    header = [str(h).strip() for h in values[skiprows]]
                    
                    # 필수 컬럼 점수 계산
                    score = 0
                    required_columns = ['STRING_ID', '#화자', 'KR']
                    for req_col in required_columns:
                        if req_col in header:
                            score += 10
                    
                    # 언어 컬럼 확인
                    lang_columns = ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']
                    for lang_col in lang_columns:
                        if lang_col in header:
                            score += 2
                    
                    # 대상 언어가 있는지 특별히 확인
                    if target_language in header:
                        score += 5
                    
                    if score > best_score:
                        best_score = score
                        best_skiprows = skiprows
                        best_header = header
                        
                except Exception as e:
                    print(f"헤더 분석 오류 (행 {skiprows}): {e}")
                    continue
            
            if best_skiprows is None or best_score < 15:  # 최소 점수 조건
                self.update_status("오류: 적절한 헤더를 찾을 수 없습니다.")
                self.root.after(0, lambda: messagebox.showerror(
                    "헤더 감지 실패", 
                    f"구글 시트에서 적절한 헤더를 찾을 수 없습니다.\n\n"
                    f"필요한 컬럼: STRING_ID, #화자, KR, {target_language}\n"
                    f"감지된 최고 점수: {best_score}/25"
                ))
                return

            self.update_status(f"헤더 감지 완료 (행 {best_skiprows + 1}, 점수: {best_score})")
            
            # 3. 데이터를 pandas DataFrame으로 변환
            import pandas as pd
            
            if best_skiprows + 1 >= len(values):
                raise ValueError("헤더 다음에 데이터가 없습니다.")
            
            data_rows = values[best_skiprows + 1:]  # 헤더 다음 행부터
            
            # 빈 행 제거 및 데이터 정리
            clean_data = []
            for row in data_rows:
                # 행이 완전히 비어있지 않은 경우만 포함
                if any(cell.strip() if isinstance(cell, str) else str(cell).strip() for cell in row if cell):
                    # 헤더 길이에 맞춰 행 길이 조정
                    while len(row) < len(best_header):
                        row.append('')
                    clean_data.append(row[:len(best_header)])  # 헤더 길이로 자르기
            
            if not clean_data:
                raise ValueError("헤더 다음에 유효한 데이터가 없습니다.")
            
            # DataFrame 생성
            df = pd.DataFrame(clean_data, columns=best_header)
            
            self.update_status(f"데이터 변환 완료: {len(df)}행 × {len(df.columns)}열")
            
            # 4. 시나리오 매니저 초기화 및 분석 실행
            if not self.scenario_manager:
                from scenario_manager import ScenarioTranslationManager
                self.scenario_manager = ScenarioTranslationManager(self.translation_db_path)
            
            self.update_status(f"{target_language} 언어로 시나리오 분석 실행 중...")
            
            # analyze_dataframe 메서드 사용
            analysis_result = self.scenario_manager.analyze_dataframe(df, target_language)
            
            # 5. 결과 처리
            if analysis_result:
                def update_ui_success():
                    self.update_speaker_list()
                    source_info = {
                        'type': 'gsheet',
                        'path': f"{spreadsheet_id}/{sheet_name}"
                    }
                    self.auto_save_reference_analysis(analysis_result, source_info, target_language)
                    # 성공 메시지 생성
                    total_speakers = len(analysis_result)
                    total_sentences = sum(data['total_sentences'] for data in analysis_result.values())
                    
                    summary = f"✅ 구글 시트 분석 완료!\n\n"
                    summary += f"📊 분석 결과:\n"
                    summary += f"• 화자 수: {total_speakers}명\n"
                    summary += f"• 총 문장 수: {total_sentences}개\n"
                    summary += f"• 분석 언어: {target_language}\n\n"
                    summary += f"📋 화자별 상세:\n"
                    
                    for speaker, data in analysis_result.items():
                        lang_count = data['languages'].get(target_language, {}).get('count', 0)
                        summary += f"• {speaker}: {lang_count}개 {target_language} 문장\n"
                    
                    messagebox.showinfo("구글 시트 분석 완료", summary)
                    self.update_status("구글 시트 레퍼런스 분석 완료")
                    
                    # 레퍼런스 파일 경로도 업데이트 (옵션)
                    self.ref_file_var.set(f"구글시트: {sheet_name}")
                
                self.root.after(0, update_ui_success)
            else:
                def update_ui_failure():
                    messagebox.showerror("분석 실패", "구글 시트 데이터 분석에 실패했습니다.")
                    self.update_status("구글 시트 분석 실패")
                
                self.root.after(0, update_ui_failure)
                
        except Exception as e:
            error_msg = f"구글 시트 로드 중 오류: {e}"
            self.update_status(error_msg)
            
            def show_error():
                messagebox.showerror("구글 시트 오류", f"구글 시트에서 레퍼런스 데이터를 가져오는 중 오류가 발생했습니다:\n\n{e}")
            
            self.root.after(0, show_error)
            
            import traceback
            traceback.print_exc()

    def prepare_scenario_translation(self):
        """시나리오 번역 준비 - 화자 정보 매핑 및 검증"""
        try:
            if self.scenario_manager:
                available_datasets = self.scenario_manager.get_available_datasets()
                if available_datasets:
                    # 가장 최근 사용된 데이터셋 자동 제안
                    latest_dataset = available_datasets[0]  # last_used DESC로 정렬됨
                    
                    if messagebox.askyesno("레퍼런스 데이터셋 발견", 
                                        f"저장된 레퍼런스 데이터셋을 발견했습니다:\n\n"
                                        f"• 이름: {latest_dataset['name']}\n"
                                        f"• 화자: {latest_dataset['total_speakers']}명\n"
                                        f"• 문장: {latest_dataset['total_sentences']}개\n"
                                        f"• 언어: {latest_dataset['target_language']}\n\n"
                                        f"이 데이터셋을 사용하시겠습니까?"):
                        
                        # 데이터셋 로드
                        loaded_data = self.scenario_manager.load_reference_dataset(latest_dataset['name'])
                        if loaded_data:
                            self.update_status(f"레퍼런스 데이터셋 '{latest_dataset['name']}' 로드됨")
                            messagebox.showinfo("데이터셋 로드 완료", 
                                            f"저장된 레퍼런스 데이터가 활성화되었습니다.\n"
                                            f"기존 화자 정보와 번역 패턴을 사용합니다.")
            
            file_path = self.file_path_var.get()
            
            if not file_path:
                raise ValueError("번역 대상 파일이 선택되지 않았습니다.")
            
            # 현재 파일에서 화자 정보 추출
            df = pd.read_excel(file_path, skiprows=3)
            
            if '#화자' not in df.columns:
                raise ValueError("번역 파일에 '#화자' 컬럼이 없습니다.")
            
            # 화자별 문장 수 집계
            speaker_stats = {}
            unknown_speakers = set()
            
            for _, row in df.iterrows():
                if pd.isna(row.get('STRING_ID')) or pd.isna(row.get('#화자')):
                    continue
                    
                speaker = str(row['#화자']).strip()
                string_id = str(row['STRING_ID'])
                
                if speaker not in speaker_stats:
                    speaker_stats[speaker] = []
                speaker_stats[speaker].append(string_id)
                
                # 등록되지 않은 화자 체크
                if speaker not in self.scenario_manager.speakers:
                    unknown_speakers.add(speaker)
            
            # 알 수 없는 화자가 있으면 사용자에게 알림
            if unknown_speakers:
                self.handle_unknown_speakers(unknown_speakers, speaker_stats)
            
            return speaker_stats
            
        except Exception as e:
            messagebox.showerror("준비 오류", f"시나리오 번역 준비 중 오류 발생: {e}")
            return None

    def handle_unknown_speakers(self, unknown_speakers, speaker_stats):
        """알 수 없는 화자 처리"""
        unknown_list = list(unknown_speakers)
        
        message = f"등록되지 않은 화자가 {len(unknown_list)}명 발견되었습니다:\n\n"
        for speaker in unknown_list[:5]:  # 최대 5명만 표시
            count = len(speaker_stats.get(speaker, []))
            message += f"• {speaker}: {count}개 문장\n"
        
        if len(unknown_list) > 5:
            message += f"... 외 {len(unknown_list) - 5}명\n"
        
        message += "\n어떻게 처리하시겠습니까?"
        
        # 선택 다이얼로그
        choice = messagebox.askyesnocancel(
            "알 수 없는 화자 발견", 
            message + "\n\n예: 자동으로 기본 프로필 생성\n아니오: 수동으로 화자 추가\n취소: 번역 중단"
        )
        
        if choice is True:  # 자동 생성
            self.auto_create_speaker_profiles(unknown_speakers)
        elif choice is False:  # 수동 추가
            self.manual_add_speakers(unknown_speakers)
        else:  # 취소
            raise ValueError("사용자가 번역을 취소했습니다.")

    def auto_create_speaker_profiles(self, unknown_speakers):
        """알 수 없는 화자들의 기본 프로필 자동 생성"""
        from scenario_manager import SpeakerProfile
        
        for speaker_name in unknown_speakers:
            # 기본 프로필 생성 (이름 기반 간단 추론)
            default_profile = SpeakerProfile(
                name=speaker_name,
                gender=self.infer_gender_from_name(speaker_name),
                tone="보통",
                style="일반적인 말투, 자동 생성된 프로필",
                examples=[]
            )
            
            self.scenario_manager.save_speaker(default_profile)
            
        self.update_speaker_list()
        self.update_status(f"{len(unknown_speakers)}개 화자 프로필이 자동 생성되었습니다.")

    def infer_gender_from_name(self, name):
        """이름으로부터 성별 간단 추론"""
        name_lower = name.lower()
        
        # 간단한 키워드 기반 추론
        male_keywords = ['king', 'prince', 'duke', 'sir', 'mr', 'lord', 'knight', '왕', '왕자', '공작']
        female_keywords = ['queen', 'princess', 'duchess', 'lady', 'ms', 'mrs', 'miss', '여왕', '공주', '공작부인']
        
        for keyword in male_keywords:
            if keyword in name_lower:
                return "남성"
        
        for keyword in female_keywords:
            if keyword in name_lower:
                return "여성"
        
        return "중성"

    def manual_add_speakers(self, unknown_speakers):
        """수동으로 화자 추가 안내"""
        messagebox.showinfo(
            "수동 추가 안내", 
            f"화자 관리 탭에서 다음 화자들을 추가해주세요:\n\n" + 
            "\n".join(f"• {speaker}" for speaker in list(unknown_speakers)[:10]) +
            "\n\n추가 완료 후 번역을 다시 실행하세요."
        )
        raise ValueError("화자를 수동으로 추가한 후 다시 시도하세요.")

    def analyze_reference_data_smart(self):
        """스마트 레퍼런스 데이터 분석 (헤더 자동 감지)"""
        ref_file = self.ref_file_var.get()
        if not ref_file:
            messagebox.showwarning("파일 없음", "레퍼런스 파일을 먼저 선택하세요.")
            return
        
        try:
            self.update_status("헤더 위치 자동 감지 중...")
            
            # 3~6행 범위에서 헤더 찾기
            best_skiprows = None
            best_score = 0
            best_columns = None
            
            for skiprows in range(3, 7):  # 3~6행 시도 (0부터 시작하므로 4~7행째)
                try:
                    df = pd.read_excel(ref_file, skiprows=skiprows, nrows=5)
                    
                    # 필수 컬럼 점수 계산
                    score = 0
                    columns = list(df.columns)
                    
                    # 필수 컬럼 확인
                    required_columns = ['STRING_ID', '#화자', 'KR']
                    for req_col in required_columns:
                        if req_col in columns:
                            score += 10
                    
                    # 언어 컬럼 확인
                    lang_columns = ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']
                    for lang_col in lang_columns:
                        if lang_col in columns:
                            score += 2
                    
                    # 데이터 유효성 확인 (빈 값이 너무 많으면 감점)
                    if len(df) > 0:
                        empty_ratio = df.isnull().sum().sum() / (len(df) * len(df.columns))
                        if empty_ratio < 0.5:  # 빈 값이 50% 미만이면 가점
                            score += 5
                    
                    print(f"skiprows={skiprows}, score={score}, columns={columns}")
                    
                    if score > best_score:
                        best_score = score
                        best_skiprows = skiprows
                        best_columns = columns
                        
                except Exception as e:
                    print(f"skiprows={skiprows} 실패: {e}")
                    continue
            
            if best_skiprows is None or best_score < 10:
                messagebox.showerror("헤더 감지 실패", 
                                "적절한 헤더 행을 찾을 수 없습니다.\n\n" +
                                "'수동 매핑' 버튼을 사용해주세요.")
                return
            
            self.update_status(f"헤더 감지 완료: {best_skiprows+1}행 (점수: {best_score})")
            
            # 사용할 언어 선택
            available_langs = [col for col in best_columns if col in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']]
            
            if not available_langs:
                messagebox.showerror("언어 컬럼 없음", "번역 가능한 언어 컬럼을 찾을 수 없습니다.")
                return
            
            # 언어 선택 다이얼로그
            if len(available_langs) == 1:
                target_lang = available_langs[0]
            else:
                dialog = LanguageSelectionDialog(self.root, available_langs, "분석할 언어 선택")
                self.root.wait_window(dialog.top)
                target_lang = dialog.selected_lang if hasattr(dialog, 'selected_lang') else None
                
                if not target_lang:
                    return
            
            # 실제 분석 실행
            if not self.scenario_manager:
                from scenario_manager import ScenarioTranslationManager
                self.scenario_manager = ScenarioTranslationManager(self.translation_db_path)
            
            self.update_status(f"{target_lang} 언어로 분석 중...")
            analysis_result = self.scenario_manager.analyze_reference_data(ref_file, target_lang, best_skiprows)
            
            if analysis_result:
                self.update_speaker_list()
                summary = f"✅ 자동 분석 완료! (헤더: {best_skiprows+1}행)\n\n"
                total_speakers = len(analysis_result)
                total_sentences = sum(data['total_sentences'] for data in analysis_result.values())
                
                summary += f"화자 수: {total_speakers}명\n"
                summary += f"총 문장 수: {total_sentences}개\n\n"
                summary += "화자별 상세:\n"
                
                for speaker, data in analysis_result.items():
                    lang_count = data['languages'].get(target_lang, {}).get('count', 0)
                    summary += f"• {speaker}: {lang_count}개 {target_lang} 문장\n"
                    
                messagebox.showinfo("분석 완료", summary)
                source_info = {
                    'type': 'file',
                    'path': ref_file
                }
                self.auto_save_reference_analysis(analysis_result, source_info, target_lang)
                self.update_status("자동 분석 완료")
            else:
                messagebox.showerror("분석 실패", "데이터 분석에 실패했습니다.")
                
        except Exception as e:
            messagebox.showerror("오류", f"자동 분석 중 오류: {e}")
            import traceback
            traceback.print_exc()

    def debug_file_structure_detailed(self):
        """레퍼런스 파일 구조 상세 디버깅"""
        ref_file = self.ref_file_var.get()
        if not ref_file:
            messagebox.showwarning("파일 없음", "레퍼런스 파일을 먼저 선택하세요.")
            return
        
        try:
            debug_info = "=== 파일 구조 상세 분석 ===\n\n"
            
            # 여러 skiprows 옵션으로 시도
            for skip_rows in range(6):  # 0~5까지 시도
                try:
                    df = pd.read_excel(ref_file, skiprows=skip_rows, nrows=10)
                    
                    debug_info += f"📊 skiprows={skip_rows} 결과:\n"
                    debug_info += f"   데이터프레임 크기: {df.shape}\n"
                    debug_info += f"   컬럼 개수: {len(df.columns)}\n\n"
                    
                    # 컬럼명 상세 분석
                    debug_info += f"   📋 컬럼명 목록:\n"
                    for i, col in enumerate(df.columns):
                        debug_info += f"      [{i}] '{col}' (타입: {type(col).__name__})\n"
                    
                    # 필수 컬럼 검사
                    required_columns = ['KR', '#화자']
                    found_required = []
                    for req_col in required_columns:
                        matches = [col for col in df.columns if str(col).strip() == req_col]
                        if matches:
                            found_required.append(req_col)
                    
                    debug_info += f"\n   ✅ 발견된 필수 컬럼: {found_required}\n"
                    
                    # 언어 컬럼 검사
                    lang_columns = []
                    for col in df.columns:
                        col_str = str(col).strip().upper()
                        if col_str in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']:
                            lang_columns.append(col_str)
                    
                    debug_info += f"   🌍 발견된 언어 컬럼: {lang_columns}\n"
                    
                    # 샘플 데이터
                    if len(df) > 0:
                        debug_info += f"\n   📝 첫 번째 행 데이터:\n"
                        for col in df.columns:
                            value = df[col].iloc[0] if len(df) > 0 else "N/A"
                            debug_info += f"      {col}: '{value}'\n"
                    
                    debug_info += "\n" + "="*50 + "\n\n"
                    
                except Exception as e:
                    debug_info += f"❌ skiprows={skip_rows} 실패: {e}\n\n"
                    continue
            
            # 결과를 새 창에 표시
            self.show_debug_result(debug_info)
            
        except Exception as e:
            messagebox.showerror("디버깅 오류", f"파일 분석 중 오류: {e}")

    def show_latest_debug_log(self):
        """최신 디버그 로그 표시"""
        try:
            import os
            import glob
            
            log_dir = "debug_logs"
            if not os.path.exists(log_dir):
                messagebox.showinfo("로그 없음", "디버그 로그 폴더가 없습니다.")
                return
            
            # 가장 최신 로그 파일 찾기
            log_files = glob.glob(os.path.join(log_dir, "analysis_debug_*.txt"))
            if not log_files:
                messagebox.showinfo("로그 없음", "디버그 로그 파일이 없습니다.")
                return
            
            latest_log = max(log_files, key=os.path.getctime)
            
            # 로그 내용 읽기
            with open(latest_log, 'r', encoding='utf-8') as f:
                log_content = f.read()
            
            # 로그 뷰어 창 생성
            log_window = tk.Toplevel(self.root)
            log_window.title(f"📜 디버그 로그 - {os.path.basename(latest_log)}")
            log_window.geometry("1000x700")
            log_window.transient(self.root)
            
            main_frame = ttk.Frame(log_window, padding="10")
            main_frame.pack(fill="both", expand=True)
            
            # 텍스트 위젯
            text_frame = ttk.Frame(main_frame)
            text_frame.pack(fill="both", expand=True)
            
            text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 9))
            scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            text_widget.insert("1.0", log_content)
            text_widget.config(state="disabled")
            
            # 버튼
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill="x", pady=10)
            
            ttk.Button(button_frame, text="📂 로그 폴더 열기", 
                    command=lambda: os.startfile(log_dir)).pack(side="left")
            ttk.Button(button_frame, text="📋 클립보드 복사", 
                    command=lambda: self.copy_to_clipboard(log_content)).pack(side="left", padx=5)
            ttk.Button(button_frame, text="닫기", command=log_window.destroy).pack(side="right")
            
        except Exception as e:
            messagebox.showerror("오류", f"디버그 로그 표시 중 오류: {e}")

    def auto_save_reference_analysis(self, analysis_result, source_info, target_language):
        """분석 완료 후 자동 저장 제안"""
        if not analysis_result:
            return
        
        # 자동 이름 생성
        from datetime import datetime
        auto_name = f"REF_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # 사용자에게 저장 여부 확인
        save_dialog = tk.Toplevel(self.root)
        save_dialog.title("레퍼런스 데이터 저장")
        save_dialog.geometry("500x300")
        save_dialog.transient(self.root)
        save_dialog.grab_set()
        
        main_frame = ttk.Frame(save_dialog, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        ttk.Label(main_frame, text="분석 완료된 레퍼런스 데이터를 저장하시겠습니까?", 
                font=("맑은 고딕", 10, "bold")).pack(pady=10)
        
        info_text = f"""• 화자 수: {len(analysis_result)}명
    • 총 문장 수: {sum(data['total_sentences'] for data in analysis_result.values())}개
    • 분석 언어: {target_language}
    • 소스: {source_info['path']}

    저장하면 다음 번역 시 레퍼런스를 다시 선택하지 않아도 됩니다."""
        
        ttk.Label(main_frame, text=info_text, justify="left").pack(pady=10)
        
        # 데이터셋 이름 입력
        name_frame = ttk.Frame(main_frame)
        name_frame.pack(fill="x", pady=10)
        
        ttk.Label(name_frame, text="데이터셋 이름:").pack(side="left")
        name_var = tk.StringVar(value=auto_name)
        ttk.Entry(name_frame, textvariable=name_var, width=30).pack(side="left", padx=5, fill="x", expand=True)
        
        result = {'save': False, 'name': ''}
        
        def save_and_close():
            result['save'] = True
            result['name'] = name_var.get().strip()
            save_dialog.destroy()
        
        def cancel_and_close():
            save_dialog.destroy()
        
        # 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="저장", command=save_and_close).pack(side="right", padx=5)
        ttk.Button(button_frame, text="저장하지 않음", command=cancel_and_close).pack(side="right")
        
        # 다이얼로그 대기
        self.root.wait_window(save_dialog)
        
        # 저장 실행
        if result['save'] and result['name']:
            # 덮어쓰기 확인 콜백 함수 정의
            def confirm_overwrite(message):
                return messagebox.askyesno("기존 데이터셋 발견", message)
            
            success = self.scenario_manager.save_reference_dataset(
                result['name'], source_info, analysis_result, target_language, confirm_overwrite
            )
            if success:
                self.update_status(f"레퍼런스 데이터셋 '{result['name']}' 저장 완료")
                messagebox.showinfo("저장 완료", f"레퍼런스 데이터셋이 성공적으로 저장되었습니다.\n이름: {result['name']}")
            else:
                messagebox.showerror("저장 실패", "레퍼런스 데이터셋 저장에 실패했습니다.")

    def show_reference_dataset_manager(self):
        """레퍼런스 데이터셋 관리 창"""
        if not self.scenario_manager:
            messagebox.showwarning("시나리오 매니저 없음", "시나리오 번역 기능이 초기화되지 않았습니다.")
            return
        
        manager_window = tk.Toplevel(self.root)
        manager_window.title("레퍼런스 데이터셋 관리")
        manager_window.geometry("800x600")
        manager_window.transient(self.root)
        
        main_frame = ttk.Frame(manager_window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 상단 버튼
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=5)
        
        ttk.Button(top_frame, text="🔄 새로고침", command=lambda: refresh_list()).pack(side="left")
        ttk.Button(top_frame, text="🗑️ 삭제", command=lambda: delete_selected()).pack(side="left", padx=5)
        ttk.Button(top_frame, text="📋 불러오기", command=lambda: load_selected()).pack(side="left", padx=5)
        
        # 데이터셋 목록
        columns = ("이름", "설명", "소스", "언어", "화자수", "문장수", "생성일", "최근사용")
        tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        tree.pack(fill="both", expand=True, pady=10)
        
        def refresh_list():
            # 기존 항목 삭제
            for item in tree.get_children():
                tree.delete(item)
            
            # 데이터셋 목록 로드
            datasets = self.scenario_manager.get_available_datasets()
            for dataset in datasets:
                tree.insert("", "end", values=(
                    dataset['name'],
                    dataset['description'][:30] + "..." if len(dataset['description']) > 30 else dataset['description'],
                    dataset['source_type'],
                    dataset['target_language'],
                    dataset['total_speakers'],
                    dataset['total_sentences'],
                    dataset['created_at'][:10],  # 날짜만
                    dataset['last_used'][:10]    # 날짜만
                ))
        
        def delete_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("선택 없음", "삭제할 데이터셋을 선택하세요.")
                return
            
            dataset_name = tree.item(selected[0])['values'][0]
            if messagebox.askyesno("삭제 확인", f"'{dataset_name}' 데이터셋을 삭제하시겠습니까?"):
                # 삭제 실행
                try:
                    conn = sqlite3.connect(self.scenario_manager.db_path)
                    cursor = conn.cursor()
                    cursor.execute("SELECT id FROM reference_datasets WHERE name = ?", (dataset_name,))
                    dataset_id = cursor.fetchone()[0]
                    cursor.execute("DELETE FROM reference_translations WHERE dataset_id = ?", (dataset_id,))
                    cursor.execute("DELETE FROM reference_datasets WHERE id = ?", (dataset_id,))
                    conn.commit()
                    conn.close()
                    refresh_list()
                    messagebox.showinfo("삭제 완료", f"'{dataset_name}' 데이터셋이 삭제되었습니다.")
                except Exception as e:
                    messagebox.showerror("삭제 실패", f"데이터셋 삭제 중 오류: {e}")
        
        def load_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("선택 없음", "불러올 데이터셋을 선택하세요.")
                return
            
            dataset_name = tree.item(selected[0])['values'][0]
            loaded_data = self.scenario_manager.load_reference_dataset(dataset_name)
            
            if loaded_data:
                # 시나리오 번역 활성화 및 설정
                self.scenario_translation_var.set(True)
                self.on_scenario_option_changed()
                
                messagebox.showinfo("불러오기 완료", 
                                f"'{dataset_name}' 데이터셋이 활성화되었습니다.\n"
                                f"시나리오 번역에서 사용할 수 있습니다.")
                manager_window.destroy()
            else:
                messagebox.showerror("불러오기 실패", "데이터셋을 불러오는데 실패했습니다.")
        # 초기 목록 로드
        refresh_list()
        
        ttk.Button(main_frame, text="닫기", command=manager_window.destroy).pack(pady=10)

    def debug_file_structure(self):
        """레퍼런스 파일 구조 디버깅"""
        ref_file = self.ref_file_var.get()
        if not ref_file:
            messagebox.showwarning("파일 없음", "레퍼런스 파일을 먼저 선택하세요.")
            return
        
        try:
            # 여러 skiprows 옵션으로 시도
            for skip_rows in [0, 1, 2, 3, 4]:
                try:
                    df = pd.read_excel(ref_file, skiprows=skip_rows, nrows=5)  # 상위 5행만 읽기
                    
                    debug_info = f"=== skiprows={skip_rows} ===\n"
                    debug_info += f"컬럼명: {list(df.columns)}\n"
                    debug_info += f"첫 번째 행: {df.iloc[0].to_dict() if len(df) > 0 else 'Empty'}\n"
                    debug_info += f"데이터 타입: {df.dtypes.to_dict()}\n\n"
                    
                    print(debug_info)
                    
                    # 언어 컬럼 찾기
                    possible_lang_cols = []
                    for col in df.columns:
                        col_str = str(col).strip().upper()
                        if col_str in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']:
                            possible_lang_cols.append(col)
                    
                    if possible_lang_cols:
                        debug_info += f"발견된 언어 컬럼: {possible_lang_cols}\n"
                        break
                        
                except Exception as e:
                    print(f"skiprows={skip_rows} 실패: {e}")
                    continue
            
            # 결과를 사용자에게 표시
            result_dialog = tk.Toplevel(self.root)
            result_dialog.title("파일 구조 분석 결과")
            result_dialog.geometry("800x600")
            
            text_widget = tk.Text(result_dialog, wrap="word", font=("Consolas", 10))
            text_widget.pack(fill="both", expand=True, padx=10, pady=10)
            text_widget.insert("1.0", debug_info)
            
            ttk.Button(result_dialog, text="닫기", command=result_dialog.destroy).pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("디버깅 오류", f"파일 분석 중 오류: {e}")

    def show_debug_result(self, debug_info):
        """디버깅 결과를 새 창에 표시"""
        result_dialog = tk.Toplevel(self.root)
        result_dialog.title("📊 파일 구조 상세 분석 결과")
        result_dialog.geometry("900x700")
        result_dialog.transient(self.root)
        
        main_frame = ttk.Frame(result_dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 텍스트 위젯 (스크롤 가능)
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True)
        
        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        text_widget.insert("1.0", debug_info)
        text_widget.config(state="disabled")
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="🔧 수동 컬럼 매핑", command=lambda: self.manual_column_mapping_dialog()).pack(side="left")
        ttk.Button(button_frame, text="📋 클립보드 복사", command=lambda: self.copy_to_clipboard(debug_info)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="닫기", command=result_dialog.destroy).pack(side="right")

    def copy_to_clipboard(self, text):
        """텍스트를 클립보드에 복사"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("복사 완료", "분석 결과가 클립보드에 복사되었습니다.")
        except Exception as e:
            messagebox.showerror("복사 실패", f"클립보드 복사 중 오류: {e}")

    def manual_column_mapping_dialog(self):
        """간단한 수동 컬럼 매핑 다이얼로그"""
        ref_file = self.ref_file_var.get()
        if not ref_file:
            messagebox.showwarning("파일 없음", "레퍼런스 파일이 선택되지 않았습니다.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("🔧 수동 헤더 설정")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill="both", expand=True)
        
        ttk.Label(main_frame, text="헤더가 있는 행 번호를 지정하세요.", font=("맑은 고딕", 10, "bold")).pack(pady=10)
        
        # skiprows 설정
        skiprows_frame = ttk.LabelFrame(main_frame, text="헤더 행 설정")
        skiprows_frame.pack(fill="x", pady=10)
        
        info_frame = ttk.Frame(skiprows_frame)
        info_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(info_frame, text="헤더 행 번호:").grid(row=0, column=0, sticky="w")
        skiprows_var = tk.IntVar(value=4)  # 기본값 4행
        skiprows_spin = ttk.Spinbox(info_frame, from_=1, to=10, textvariable=skiprows_var, width=10)
        skiprows_spin.grid(row=0, column=1, padx=5)
        ttk.Label(info_frame, text="행").grid(row=0, column=2, sticky="w")
        
        ttk.Label(info_frame, text="(예: 4행에 헤더가 있으면 4 입력)").grid(row=1, column=0, columnspan=3, sticky="w", pady=2)
        
        # 미리보기
        preview_frame = ttk.LabelFrame(main_frame, text="미리보기")
        preview_frame.pack(fill="both", expand=True, pady=5)
        
        preview_text = tk.Text(preview_frame, height=8, wrap="word", font=("Consolas", 9))
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=preview_text.yview)
        preview_text.configure(yscrollcommand=preview_scroll.set)
        
        preview_text.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        preview_scroll.pack(side="right", fill="y", pady=5)
        
        def update_preview():
            try:
                skiprows = skiprows_var.get() - 1  # 사용자는 1부터, pandas는 0부터
                df = pd.read_excel(ref_file, skiprows=skiprows, nrows=3)
                
                preview_info = f"헤더 행: {skiprows_var.get()}\n"
                preview_info += f"컬럼 개수: {len(df.columns)}\n\n"
                preview_info += "발견된 컬럼:\n"
                
                for i, col in enumerate(df.columns):
                    preview_info += f"{i+1:2d}. {col}\n"
                
                # 필수 컬럼 체크
                required = ['STRING_ID', '#화자', 'KR']
                found_required = [col for col in required if col in df.columns]
                missing_required = [col for col in required if col not in df.columns]
                
                preview_info += f"\n✅ 발견된 필수 컬럼: {found_required}\n"
                if missing_required:
                    preview_info += f"❌ 누락된 필수 컬럼: {missing_required}\n"
                
                # 언어 컬럼 체크
                lang_cols = [col for col in df.columns if col in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']]
                preview_info += f"🌍 언어 컬럼: {lang_cols}\n"
                
                preview_text.delete("1.0", "end")
                preview_text.insert("1.0", preview_info)
                
            except Exception as e:
                preview_text.delete("1.0", "end")
                preview_text.insert("1.0", f"미리보기 오류: {e}")
        
        # 초기 미리보기
        update_preview()
        
        # skiprows 변경 시 미리보기 업데이트
        skiprows_spin.configure(command=update_preview)
        
        # 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="🔄 미리보기 새로고침", command=update_preview).pack(side="left")
        
        def apply_and_analyze():
            try:
                skiprows = skiprows_var.get() - 1
                df = pd.read_excel(ref_file, skiprows=skiprows, nrows=1)
                
                # 필수 컬럼 확인
                required = ['STRING_ID', '#화자', 'KR']
                missing = [col for col in required if col not in df.columns]
                
                if missing:
                    messagebox.showerror("필수 컬럼 없음", f"다음 필수 컬럼이 없습니다: {missing}")
                    return
                
                dialog.destroy()
                
                # 분석 실행
                self.analyze_with_skiprows(ref_file, skiprows)
                
            except Exception as e:
                messagebox.showerror("오류", f"분석 실행 오류: {e}")
        
        ttk.Button(button_frame, text="취소", command=dialog.destroy).pack(side="right", padx=5)
        ttk.Button(button_frame, text="✅ 분석 실행", command=apply_and_analyze).pack(side="right")

    def analyze_with_skiprows(self, file_path, skiprows):
        """지정된 skiprows로 분석 실행"""
        try:
            # 언어 선택
            df = pd.read_excel(file_path, skiprows=skiprows, nrows=1)
            available_langs = [col for col in df.columns if col in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']]
            
            if not available_langs:
                messagebox.showerror("언어 없음", "번역 가능한 언어 컬럼이 없습니다.")
                return
            
            if len(available_langs) == 1:
                target_lang = available_langs[0]
            else:
                dialog = LanguageSelectionDialog(self.root, available_langs, "분석할 언어 선택")
                self.root.wait_window(dialog.top)
                target_lang = dialog.selected_lang if hasattr(dialog, 'selected_lang') else None
                
                if not target_lang:
                    return
            
            # 분석 실행
            if not self.scenario_manager:
                from scenario_manager import ScenarioTranslationManager
                self.scenario_manager = ScenarioTranslationManager(self.translation_db_path)
            
            self.update_status(f"{target_lang} 언어로 분석 중...")
            analysis_result = self.scenario_manager.analyze_reference_data(file_path, target_lang, skiprows)
            
            if analysis_result:
                self.update_speaker_list()
                summary = f"✅ 수동 설정 분석 완료!\n\n"
                summary += f"화자 수: {len(analysis_result)}명\n"
                summary += f"총 문장 수: {sum(data['total_sentences'] for data in analysis_result.values())}개\n\n"
                
                for speaker, data in analysis_result.items():
                    lang_count = data['languages'].get(target_lang, {}).get('count', 0)
                    summary += f"• {speaker}: {lang_count}개 {target_lang} 문장\n"
                    
                messagebox.showinfo("분석 완료", summary)
                
                source_info = {
                    'type': 'file', 
                    'path': file_path
                }
                self.auto_save_reference_analysis(analysis_result, source_info, target_lang)
                
                self.update_status("수동 설정 분석 완료")
            else:
                messagebox.showerror("분석 실패", "데이터 분석에 실패했습니다.")
                
        except Exception as e:
            messagebox.showerror("오류", f"분석 중 오류: {e}")
            import traceback
            traceback.print_exc()

    def find_language_columns(self, columns):
        """컬럼명에서 언어 코드 찾기 (유연한 매칭)"""
        available_langs = []
        
        for col in columns:
            col_str = str(col).strip().upper()
            
            # 정확한 매칭
            if col_str in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']:
                available_langs.append(col_str)
            # 부분 매칭 (예: "English", "EN_US" 등)
            elif any(lang in col_str for lang in ['EN', 'ENGLISH']):
                available_langs.append('EN')
            elif any(lang in col_str for lang in ['CN', 'CHINESE', 'ZH_CN']):
                available_langs.append('CN')
            elif any(lang in col_str for lang in ['TW', 'ZH_TW', 'TRADITIONAL']):
                available_langs.append('TW')
            # 필요에 따라 다른 언어도 추가
        
        return list(set(available_langs))

    def show_manual_column_mapping(self, file_path):
        """수동 컬럼 매핑 다이얼로그"""
        try:
            # 첫 5행 정도 읽어서 사용자에게 보여주기
            df = pd.read_excel(file_path, nrows=5)
            
            dialog = tk.Toplevel(self.root)
            dialog.title("수동 컬럼 매핑")
            dialog.geometry("600x400")
            dialog.transient(self.root)
            dialog.grab_set()
            
            ttk.Label(dialog, text="파일의 컬럼 구조를 확인하고 언어 컬럼을 수동으로 지정하세요.").pack(pady=10)
            
            # 컬럼 정보 표시
            info_frame = ttk.LabelFrame(dialog, text="발견된 컬럼")
            info_frame.pack(fill="both", expand=True, padx=10, pady=5)
            
            text_widget = tk.Text(info_frame, height=10, wrap="word")
            text_widget.pack(fill="both", expand=True, padx=5, pady=5)
            
            column_info = f"컬럼 목록: {list(df.columns)}\n\n"
            column_info += "첫 번째 행 데이터:\n"
            for col in df.columns:
                column_info += f"{col}: {df[col].iloc[0] if len(df) > 0 else 'N/A'}\n"
            
            text_widget.insert("1.0", column_info)
            text_widget.config(state="disabled")
            
            # 수동 매핑 프레임
            mapping_frame = ttk.LabelFrame(dialog, text="언어 컬럼 매핑")
            mapping_frame.pack(fill="x", padx=10, pady=5)
            
            ttk.Label(mapping_frame, text="영어 컬럼:").grid(row=0, column=0, sticky="w", padx=5)
            en_combo = ttk.Combobox(mapping_frame, values=list(df.columns), state="readonly")
            en_combo.grid(row=0, column=1, sticky="ew", padx=5)
            
            mapping_frame.grid_columnconfigure(1, weight=1)
            
            # 버튼
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", padx=10, pady=5)
            
            def apply_mapping():
                en_col = en_combo.get()
                if en_col:
                    # 임시로 EN 컬럼 매핑 저장
                    self.manual_column_mapping = {'EN': en_col}
                    dialog.destroy()
                else:
                    messagebox.showwarning("선택 오류", "최소한 영어 컬럼은 선택해야 합니다.")
            
            ttk.Button(button_frame, text="적용", command=apply_mapping).pack(side="right", padx=5)
            ttk.Button(button_frame, text="취소", command=dialog.destroy).pack(side="right")
            
        except Exception as e:
            messagebox.showerror("오류", f"수동 매핑 다이얼로그 오류: {e}")

    def get_speaker_for_item(self, trans_item, speaker_mapping):
        """번역 항목의 화자 정보 가져오기 (최적화된 버전)"""
        if not speaker_mapping:
            return None
            
        string_id = trans_item["STRING_ID"]
        
        # speaker_mapping에서 해당 STRING_ID가 포함된 화자 찾기
        for speaker, string_ids in speaker_mapping.items():
            if string_id in string_ids:
                return speaker
        
        return None

    def get_speaker_for_translation(self, trans_item):
        """번역 대상 항목의 화자 정보 가져오기"""
        try:
            # 현재 번역 파일에서 해당 STRING_ID의 화자 정보 찾기
            file_path = self.file_path_var.get()
            if not file_path:
                return None
                
            # 캐시된 데이터에서 찾기 (성능 개선)
            if not hasattr(self, '_speaker_cache'):
                self._speaker_cache = {}
                df = pd.read_excel(file_path, skiprows=3)
                if '#화자' in df.columns and 'STRING_ID' in df.columns:
                    for _, row in df.iterrows():
                        if not pd.isna(row['STRING_ID']) and not pd.isna(row['#화자']):
                            self._speaker_cache[str(row['STRING_ID'])] = str(row['#화자']).strip()
            
            return self._speaker_cache.get(trans_item['STRING_ID'])
            
        except Exception as e:
            print(f"화자 정보 조회 오류: {e}")
            return None

    def show_translation_report(self, report_data):
        """번역 작업이 끝난 후 결과 요약 창을 표시합니다."""
        TranslationReportDialog(self.root, "번역 결과 요약", report_data)

    def setup_keyboard_shortcuts(self):
        """키보드 단축키 설정"""
        self.root.bind('<Control-o>', lambda e: self.select_file())          # Ctrl+O: 파일 열기
        self.root.bind('<Control-l>', lambda e: self.load_data())            # Ctrl+L: 데이터 로드  
        self.root.bind('<Control-t>', lambda e: self.analyze_translations()) # Ctrl+T: 번역 분석
        self.root.bind('<Control-r>', lambda e: self.execute_translation())  # Ctrl+R: 번역 실행
        self.root.bind('<Control-s>', lambda e: self.save_results())         # Ctrl+S: 결과 저장
        self.root.bind('<F5>', lambda e: self.update_translation_table())

    def setup_compact_tabs(self):
        """탭 제목을 더 컴팩트하게 수정"""
        # 이미 위의 setup_ui에서 적용됨
        pass

    def setup_progress_bar(self):
        """진행바 설정 개선"""
        style = ttk.Style()
        style.configure("Compact.Horizontal.TProgressbar", thickness=15)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.main_frame, 
            variable=self.progress_var, 
            style="Compact.Horizontal.TProgressbar"
        )

    def setup_status_help(self):
        """상태바 하단에 도움말 텍스트 추가"""
        help_frame = ttk.Frame(self.main_frame)
        help_frame.pack(fill="x", pady=2)
        
        help_text = "💡 단축키: Ctrl+O(파일열기) | Ctrl+L(로드) | Ctrl+T(분석) | Ctrl+R(번역) | Ctrl+S(저장) | F5(새로고침)"
        help_label = ttk.Label(help_frame, text=help_text, font=("맑은 고딕", 8), foreground="gray")
        help_label.pack()

    def setup_responsive_layout(self):
        """반응형 레이아웃 설정"""
        # 최소 크기 설정
        self.root.minsize(1200, 700)
        
        # 윈도우 크기 변경 시 컴포넌트 크기 자동 조정
        def on_window_resize(event):
            if event.widget == self.root:
                # 윈도우가 너무 작아지면 스크롤바 표시
                if self.root.winfo_width() < 1200:
                    # 컴포넌트들을 더 작게 조정
                    pass
        
        self.root.bind('<Configure>', on_window_resize)

    def setup_theme_support(self):
        """테마 지원 설정"""
        style = ttk.Style()
        
        # 사용 가능한 테마 확인
        available_themes = style.theme_names()
        
        # 시스템에 따른 기본 테마 선택
        if 'vista' in available_themes:
            style.theme_use('vista')  # Windows 10/11
        elif 'aqua' in available_themes:
            style.theme_use('aqua')   # macOS
        else:
            style.theme_use('clam')

    def setup_memory_monitor(self):
        """메모리 사용량 모니터링 (개발자용)"""
        import psutil
        import threading
        
        def update_memory_info():
            while True:
                try:
                    process = psutil.Process()
                    memory_mb = process.memory_info().rss / 1024 / 1024
                    
                    if hasattr(self, 'memory_label'):
                        self.root.after(0, lambda: self.memory_label.config(
                            text=f"메모리: {memory_mb:.1f}MB"
                        ))
                    
                    time.sleep(5)  # 5초마다 업데이트
                except:
                    break
        
        # 메모리 표시 라벨 (상태바에 추가)
        self.memory_label = ttk.Label(self.status_frame, text="메모리: 0MB", font=("맑은 고딕", 8))
        self.memory_label.pack(side="left", padx=10)
        
        # 백그라운드 스레드로 모니터링 시작
        memory_thread = threading.Thread(target=update_memory_info, daemon=True)
        memory_thread.start()

    def setup_enhanced_ui(self):
        """향상된 UI 설정 통합 메서드"""
        self.setup_ui()                    # 메인 UI
        self.setup_keyboard_shortcuts()    # 키보드 단축키
        self.setup_status_help()          # 도움말
        self.setup_responsive_layout()     # 반응형 레이아웃
        self.setup_theme_support()

    def setup_translation_tab(self):
        """번역 대상 탭 설정 (동적 컬럼 생성)"""
        # 검색 및 필터 (기존과 동일)
        filter_frame = ttk.Frame(self.translation_tab)
        filter_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(filter_frame, text="검색:").pack(side="left")
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self.filter_translations())
        
        ttk.Label(filter_frame, text="상태 필터:").pack(side="left", padx=20)
        self.filter_vars = {
            "신규": tk.BooleanVar(value=True), 
            "변경": tk.BooleanVar(value=True),
            "확인필요": tk.BooleanVar(value=True), 
            "확정": tk.BooleanVar(value=True),      # False → True로 변경
            "완료": tk.BooleanVar(value=True),
            "재번역완료": tk.BooleanVar(value=True)  # ← 새로 추가
        }
        for status, var in self.filter_vars.items():
            ttk.Checkbutton(filter_frame, text=status, variable=var,
                        command=self.filter_translations).pack(side="left", padx=5)
        
        # <<< 시작: 번역 방법 필터 추가 >>>
        ttk.Label(filter_frame, text="방법 필터:").pack(side="left", padx=(20, 5))
        self.method_filter_var = tk.StringVar(value="전체")
        self.method_filter_combo = ttk.Combobox(filter_frame, textvariable=self.method_filter_var, state="readonly", width=15)
        self.method_filter_combo.pack(side="left")
        self.method_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_translations())
        # <<< 종료: 번역 방법 필터 추가 >>>
        
        self.stats_frame = ttk.Frame(filter_frame)
        self.stats_frame.pack(side="right", padx=10)
        self.update_stats_label()  
            
        # <<< 시작: 전체 선택/해제 체크박스 추가 >>>
        select_all_frame = ttk.Frame(self.translation_tab)
        select_all_frame.pack(fill="x", padx=5, pady=(5,0))
        self.select_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(select_all_frame, text="전체 선택/해제", variable=self.select_all_var, command=self.toggle_all_selections).pack(side="left")
        # <<< 종료: 전체 선택/해제 체크박스 추가 >>>
        
        # <<< 시작: 동적 컬럼 리스트 생성 >>>
        base_columns = ["선택", "STRING_ID", "KR", "상태"]
        lang_columns = self.VISIBLE_LANGS  # __init__에서 정의한 언어 순서 사용
        end_columns = ["번역방법"]
        columns = base_columns + lang_columns + end_columns
        # <<< 종료: 동적 컬럼 리스트 생성 >>>

        tree_frame = ttk.Frame(self.translation_tab)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.translation_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings")
        
        # 컬럼 설정
        self.translation_tree.column("#0", width=0, stretch=False)
        self.translation_tree.heading("#0", text="")

        # 고정 컬럼 설정
        self.translation_tree.column("선택", width=40, anchor="center")
        self.translation_tree.column("STRING_ID", width=150)
        self.translation_tree.column("KR", width=250)
        self.translation_tree.column("상태", width=80)
        self.translation_tree.column("번역방법", width=100)

        # <<< 시작: 동적 언어 컬럼 설정 >>>
        for lang_col in lang_columns:
            self.translation_tree.column(lang_col, width=150)
        # <<< 종료: 동적 언어 컬럼 설정 >>>

        for col in columns:
            self.translation_tree.heading(col, text=col)
        
        # 스크롤바 (기존과 동일)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.translation_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.translation_tree.xview)
        self.translation_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.translation_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # 우클릭 컨텍스트 메뉴 생성
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="🔄 이 항목 재번역 (API 강제)", command=self.force_retranslate_selected)
        self.context_menu.add_command(label="✏️ 직접 편집", command=self.edit_translation_inline)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🗑️ TM에서 삭제", command=self.remove_from_tm)
        self.context_menu.add_command(label="📋 TM 항목 보기", command=self.view_tm_entry)
        
        # 트리뷰에 우클릭 이벤트 바인딩
        self.translation_tree.bind("<Button-3>", self.show_context_menu)  # 우클릭
        self.translation_tree.bind("<Control-Button-1>", self.show_context_menu)  # Ctrl+클릭 (Mac 호환)
        self.translation_tree.bind("<Button-1>", self.on_tree_click)  # 기존 클릭 이벤트
                
        # 키보드 이벤트 바인딩
        self.translation_tree.bind("<KeyPress-space>", self.toggle_selected_checkboxes)
        self.translation_tree.bind("<Return>", self.toggle_selected_checkboxes)  # Enter키도 동일하게
        self.translation_tree.bind("<Control-a>", self.select_all_items)  # Ctrl+A로 전체 선택
        self.translation_tree.bind("<Control-d>", self.deselect_all_items)  # Ctrl+D로 전체 해제
        self.translation_tree.bind("<Control-i>", self.invert_selection)  # Ctrl+I: 선택 반전
        self.translation_tree.bind("<Delete>", self.clear_selected_translations)  # Delete: 선택된 항목의 번역 내용 삭제
        self.translation_tree.bind("<F2>", self.edit_selected_item)  # F2: 선택된 항목 편집

        # 포커스 설정으로 키보드 이벤트가 작동하도록
        self.translation_tree.focus_set()
        
        # 키보드 단축키 안내 추가
        help_frame = ttk.Frame(self.translation_tab)
        help_frame.pack(fill="x", padx=5, pady=2)
        
        help_text = "💡 키보드 단축키: Space/Enter=체크토글 | Ctrl+A=전체선택 | Ctrl+D=전체해제 | Ctrl+I=선택반전 | F2=편집 | Del=번역삭제"
        help_label = ttk.Label(help_frame, text=help_text, font=("맑은 고딕", 8), foreground="gray")
        help_label.pack(side="right")        
        
        self.check_states = {}

 #디버그 관련
    def debug_tm_status(self):
        """TM 상태 디버깅"""
        status_info = f"""TM 상태 정보:
        
    메모리상 TM 항목 수: {len(self.translation_memory)}
    DB TM 항목 수: {self.db_manager.get_db_tm_count()}

    최근 5개 TM 항목:
    {list(self.translation_memory.keys())[:5]}

    분석 대상 첫 5개 KR:
    {[t["KR"] for t in self.pending_translations[:5]]}
    """
        messagebox.showinfo("TM 상태", status_info)


    def cleanup_tm_with_rules(self):
        """'TM 정리하기' 버튼의 동작. 확인 후 정리 스레드를 실행합니다."""
        if messagebox.askyesno("TM 정리 확인", "현재 '제외 규칙'을 기준으로 마스터 TM 전체를 검사하여, 규칙에 위배되는 모든 항목을 영구적으로 삭제합니다.\n\n이 작업은 되돌릴 수 없습니다. 계속하시겠습니까?"):
            threading.Thread(target=self._cleanup_tm_thread, daemon=True).start()


#추가한 함수 TM 관련
    def auto_resolve_all_conflicts(self):
        """모든 충돌 항목을 가장 빈도가 높은 번역으로 자동 해결"""
        if not messagebox.askyesno("경고", "모든 충돌 항목을 자동으로 해결하시겠습니까? 이 작업은 되돌릴 수 없습니다."):
            return
            
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT kr_text, translations, conflict_info FROM translation_memory WHERE status = 'conflict'")
        conflicts = cursor.fetchall()
        
        resolved_count = 0
        for kr_text, trans_json, conflict_json in conflicts:
            current_translations = json.loads(trans_json)
            conflict_info = json.loads(conflict_json)
            
            for lang, candidates in conflict_info.items():
                most_common = max(candidates, key=candidates.get)
                current_translations[lang] = most_common
                
            cursor.execute("""
                UPDATE translation_memory 
                SET translations=?, status='consolidated', conflict_info=NULL 
                WHERE kr_text=?
            """, (json.dumps(current_translations), kr_text))
            resolved_count += 1
            
        conn.commit()
        conn.close()
        
        messagebox.showinfo("자동 해결 완료", f"{resolved_count}개의 충돌 항목이 자동으로 해결되었습니다.")
        self.load_conflicts_to_view()
        self.load_tm_view()


    def load_conflicts_to_view(self):
        """DB에서 충돌 상태인 항목을 불러와 충돌 해결 탭에 표시"""
        self.conflict_tree.delete(*self.conflict_tree.get_children())
        
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT kr_text, translations, conflict_info FROM translation_memory WHERE status = 'conflict'")
        
        conflicts = cursor.fetchall()
        conn.close()

        for kr, trans_json, conflict_json in conflicts:
            translations = json.loads(trans_json)
            conflict_info = json.loads(conflict_json) if conflict_json else {}
            
            display_values = [kr]
            tags = [''] * (len(self.VISIBLE_LANGS) + 1)
            
            for i, lang in enumerate(self.VISIBLE_LANGS):
                display_values.append(translations.get(lang, ""))
                if lang in conflict_info:
                    tags[i+1] = 'conflict' # KR 다음부터 언어 컬럼이므로 i+1
                    
            self.conflict_tree.insert("", "end", values=display_values, tags=tags)
    
    def resolve_selected_conflict(self):
        """사용자가 선택한 값으로 충돌을 해결하고 DB에 반영"""
        kr_text = self.conflict_kr_var.get()
        if not kr_text:
            messagebox.showwarning("선택 오류", "해결할 항목을 목록에서 선택하세요.")
            return

        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT translations FROM translation_memory WHERE kr_text=?", (kr_text,))
        current_translations = json.loads(cursor.fetchone()[0])

        # 콤보박스에서 선택된 값으로 번역문 업데이트
        for lang, combo in self.conflict_combos.items():
            if combo.cget("state") != "disabled":
                selected_index = combo.current()
                if selected_index != -1:
                    current_translations[lang] = combo.real_values[selected_index]
        
        # DB 업데이트: status를 'consolidated'로 변경하고, conflict_info를 null로, 확정된 번역문을 저장
        cursor.execute("""
            UPDATE translation_memory 
            SET translations=?, status='consolidated', conflict_info=NULL 
            WHERE kr_text=?
        """, (json.dumps(current_translations), kr_text))
        
        conn.commit()
        conn.close()

        messagebox.showinfo("해결 완료", f"'{kr_text}' 항목의 충돌이 해결되었습니다.")
        self.load_conflicts_to_view() # 충돌 목록 새로고침
        self.load_tm_view() # 마스터 TM 뷰도 새로고침

        
    def load_translation_memory(self):
        """번역 메모리 로드"""
        # 기존 unique_texts.db에서 로드
        if os.path.exists(self.db_manager.legacy_db_path):
            try:
                conn = sqlite3.connect(self.db_manager.legacy_db_path)
                cursor = conn.cursor()
                
                cursor.execute("SELECT KR, EN, CN, TW, JP, DE, FR, TH, PT, ES FROM unique_texts")
                rows = cursor.fetchall()
                
                for row in rows:
                    kr = row[0]
                    translations = {
                        "EN": row[1] or "",
                        "CN": row[2] or "",
                        "TW": row[3] or "",
                        "JP": row[4] or "",
                        "DE": row[5] or "",
                        "FR": row[6] or "",
                        "TH": row[7] or "",
                        "PT": row[8] or "",
                        "ES": row[9] or ""
                    }
                    self.translation_memory[kr] = translations
                
                conn.close()
                self.update_status(f"번역 메모리 로드 완료: {len(self.translation_memory)}개")
            except Exception as e:
                self.update_status(f"번역 메모리 로드 오류: {str(e)}")
        
        # smart_translations.db에서도 로드
        try:
            conn = sqlite3.connect(self.translation_db_path)
            cursor = conn.cursor()
            
            cursor.execute("SELECT kr_text, translations FROM translation_memory")
            rows = cursor.fetchall()
            
            for kr, trans_json in rows:
                translations = json.loads(trans_json)
                if kr in self.translation_memory:
                    # 병합
                    for lang, text in translations.items():
                        if text and not self.translation_memory[kr].get(lang):
                            self.translation_memory[kr][lang] = text
                else:
                    self.translation_memory[kr] = translations
            
            conn.close()
        except Exception as e:
            print(f"스마트 번역 DB 로드 오류: {str(e)}")
  

    def load_tm_view(self):
        """DB의 TM을 읽어 관리 탭의 테이블에 표시"""
        self.tm_view_tree.delete(*self.tm_view_tree.get_children())
        search_term = self.tm_view_search_var.get()
        
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()

        if search_term:
            cursor.execute("SELECT kr_text, translations FROM translation_memory WHERE kr_text LIKE ?", (f"%{search_term}%",))
        else:
            cursor.execute("SELECT kr_text, translations FROM translation_memory")
        
        for kr, trans_json in cursor.fetchall():
            translations = json.loads(trans_json)
            values = [kr] + [translations.get(lang, "") for lang in self.VISIBLE_LANGS]
            self.tm_view_tree.insert("", "end", values=values)

        conn.close()

#용어집 관련 추가
    def sync_glossary_from_gsheet(self):
        """'구글 시트와 동기화' 버튼의 동작. (시트 이름 glossary로 고정)"""
        if not messagebox.askyesno("동기화 확인", "마스터 구글 시트의 최신 내용으로 로컬 용어집 DB를 동기화합니다.\n\n로컬 DB의 내용은 구글 시트 기준으로 덮어쓰거나 삭제될 수 있습니다. 계속하시겠습니까?", parent=self.root):
            return

        # 구글 시트 URL이나 ID를 입력받습니다.
        sheet_url_or_id = simpledialog.askstring("구글 시트 정보 입력", "구글 시트의 전체 URL 또는 스프레드시트 ID를 입력하세요:", initialvalue="1Ff7jFAmpgMLDFQ2S0pnT2sNkKfE__xW2jRnA5AQOZqY", parent=self.root)
        if not sheet_url_or_id: return

        # 시트 이름을 "glossary"로 고정
        sheet_name = "glossary"

        # 정규식을 이용해 URL에서 스프레드시트 ID만 추출
        match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_url_or_id)
        spreadsheet_id = match.group(1) if match else sheet_url_or_id

        threading.Thread(target=self._sync_gsheet_thread, args=(spreadsheet_id, sheet_name), daemon=True).start()
    
    def _sync_gsheet_thread(self, spreadsheet_id, sheet_name):
        """(스레드) 구글 시트와 로컬 DB 동기화 (STRING_ID 제거)"""
        try:
            self.update_status("동기화 시작: 구글 시트 데이터 가져오는 중...")
            
            from google.oauth2 import service_account
            from googleapiclient.discovery import build

            creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
            service = build('sheets', 'v4', credentials=creds)
            range_name = sheet_name
            
            result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            master_values = result.get('values', [])
            
            if not master_values or len(master_values) < 2:
                self.update_status("오류: 마스터 구글 시트에서 데이터를 찾을 수 없습니다.")
                return

            master_header = [h.lower() for h in master_values[0]]
            
            # STRING_ID가 있으면 무시하고 kr을 기준으로 처리
            master_data = {}
            for row in master_values[1:]:
                row_dict = dict(zip(master_header, row))
                kr = row_dict.get('kr')
                if kr and kr.strip():  # kr을 키로 사용
                    master_data[kr.strip()] = row_dict

            # 로컬 DB 데이터 로드 (kr 기준)
            self.update_status("로컬 캐시 DB와 데이터 비교 중...")
            conn = sqlite3.connect(self.translation_db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM glossary")
            local_data = {row['kr']: dict(row) for row in cursor.fetchall() if row['kr']}

            # 변경점 계산
            master_keys = set(master_data.keys())
            local_keys = set(local_data.keys())
            
            new_keys = master_keys - local_keys
            deleted_keys = local_keys - master_keys
            common_keys = master_keys.intersection(local_keys)
            
            updated_keys = set()
            for key in common_keys:
                # 간단한 비교
                master_row_str = str(sorted(master_data[key].items()))
                local_row_str = str(sorted(local_data[key].items()))
                if master_row_str != local_row_str:
                    updated_keys.add(key)
            
            if not any([new_keys, deleted_keys, updated_keys]):
                self.update_status("동기화 완료: 변경된 내용이 없습니다.")
                self.root.after(0, lambda: messagebox.showinfo("완료", "용어집이 이미 최신 상태입니다.", parent=self.root))
                conn.close()
                return
                    
            # 로컬 DB에 변경사항 적용
            self.update_status("로컬 캐시 DB에 변경사항 적용 중...")
            
            # 삭제
            if deleted_keys:
                cursor.executemany("DELETE FROM glossary WHERE kr = ?", [(key,) for key in deleted_keys])
            
            # 신규 및 변경 (INSERT OR REPLACE로 처리)
            keys_to_upsert = new_keys.union(updated_keys)
            upsert_data = []
            # STRING_ID 제외한 컬럼들
            db_cols = ["kr", "en", "cn", "tw", "th", "pt", "es", "de", "fr", "jp", "engine", "contributor", "update_at", "verified", "description"]
            
            for key in keys_to_upsert:
                row_data = master_data[key]
                db_values = tuple(row_data.get(col, None) for col in db_cols)
                upsert_data.append(db_values)

            if upsert_data:
                cursor.executemany(f"INSERT OR REPLACE INTO glossary ({', '.join(db_cols)}) VALUES ({','.join(['?']*len(db_cols))})", upsert_data)

            conn.commit()
            conn.close()

            # 완료 및 UI 새로고침
            self.root.after(0, self.load_glossary_and_update_ui)
            self.update_status("동기화 완료!")
            summary = f"동기화가 완료되었습니다.\n\n- 신규: {len(new_keys)}개\n- 변경: {len(updated_keys)}개\n- 삭제: {len(deleted_keys)}개"
            self.root.after(0, lambda: messagebox.showinfo("동기화 완료", summary, parent=self.root))
            
        except Exception as e:
            self.update_status(f"동기화 오류: {e}")
            self.root.after(0, lambda: messagebox.showerror("오류", f"동기화 중 오류가 발생했습니다:\n{e}", parent=self.root))
            import traceback
            traceback.print_exc()
    
    def add_exclusion_rule(self):
        """새로운 제외 규칙을 DB에 추가"""
        desc = self.rule_desc_var.get().strip()
        rule_type = self.rule_type_var.get()
        field = self.rule_field_var.get()
        value = self.rule_value_var.get().strip()
        
        if not all([desc, rule_type, field, value]):
            messagebox.showwarning("입력 오류", "모든 필드를 입력해야 합니다.")
            return

        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO exclusion_rules (description, rule_type, field, value) VALUES (?, ?, ?, ?)",
                    (desc, rule_type, field, value))
        conn.commit()
        conn.close()
        
        # 입력 필드 초기화
        self.rule_desc_var.set("")
        self.rule_value_var.set("")
        
        self.load_exclusion_rules()

    def load_exclusion_rules(self):
        """DB에서 제외 규칙을 불러와 테이블에 표시"""
        self.exclusion_rule_tree.delete(*self.exclusion_rule_tree.get_children())
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, description, rule_type, field, value, is_enabled FROM exclusion_rules ORDER BY id")
        
        for row in cursor.fetchall():
            # is_enabled 값이 1이면 '활성화', 0이면 '비활성화'로 표시
            enabled_text = "활성화" if row[5] == 1 else "비활성화"
            # treeview에는 id를 내부 id로, 값에는 표시될 내용만 전달
            self.exclusion_rule_tree.insert("", "end", iid=row[0], values=(row[1], row[2], row[3], row[4], enabled_text))
        
        conn.close()

    def add_exclusion_rule(self):
        """새로운 제외 규칙을 DB에 추가"""
        desc = self.rule_desc_var.get().strip()
        rule_type = self.rule_type_var.get()
        field = self.rule_field_var.get()
        value = self.rule_value_var.get().strip()
        
        if not all([desc, rule_type, field, value]):
            messagebox.showwarning("입력 오류", "모든 필드를 입력해야 합니다.")
            return

        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO exclusion_rules (description, rule_type, field, value) VALUES (?, ?, ?, ?)",
                    (desc, rule_type, field, value))
        conn.commit()
        conn.close()
        
        # 입력 필드 초기화
        self.rule_desc_var.set("")
        self.rule_value_var.set("")
        
        self.load_exclusion_rules()

    def delete_exclusion_rule(self):
        """선택된 제외 규칙을 DB에서 삭제"""
        selected_items = self.exclusion_rule_tree.selection()
        if not selected_items: return

        if messagebox.askyesno("삭제 확인", f"{len(selected_items)}개의 규칙을 삭제하시겠습니까?"):
            conn = sqlite3.connect(self.translation_db_path)
            cursor = conn.cursor()
            for item_id in selected_items:
                cursor.execute("DELETE FROM exclusion_rules WHERE id = ?", (item_id,))
            conn.commit()
            conn.close()
            self.load_exclusion_rules()

    def toggle_exclusion_rule(self):
        """선택된 규칙의 활성화 상태를 변경"""
        selected_items = self.exclusion_rule_tree.selection()
        if not selected_items: return

        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        for item_id in selected_items:
            # 현재 상태를 확인하기 위해 is_enabled 값을 DB에서 직접 읽어옴
            cursor.execute("SELECT is_enabled FROM exclusion_rules WHERE id = ?", (item_id,))
            current_status = cursor.fetchone()[0]
            # 상태 토글 (1 -> 0, 0 -> 1)
            new_status = 1 - current_status
            cursor.execute("UPDATE exclusion_rules SET is_enabled = ? WHERE id = ?", (new_status, item_id))
        conn.commit()
        conn.close()
        self.load_exclusion_rules()

    def reset_default_rules(self):
        """모든 규칙을 삭제하고 기본 규칙 세트로 초기화"""
        if not messagebox.askyesno("초기화 확인", "모든 규칙을 삭제하고 기본값으로 되돌리시겠습니까? 이 작업은 되돌릴 수 없습니다."):
            return
            
        default_rules = [
            ('#으로 시작하는 KR 제외', 'startswith', 'KR', '#', 1),
            ('cs_로 시작하는 STRING_ID 제외', 'startswith', 'STRING_ID', 'cs_', 1),
            ('KR의 길이가 20 이상인 항목 제외', 'length', 'KR', '20', 0), # 기본 비활성화 예시
            ('\\n\\n으로 시작하는 언어 제외', 'startswith', 'KR', '\\n\\n', 1),
            ('[@...] 형식 제외', 'regex', 'KR', r'^\[@.\]\w*$', 1)
        ]
        
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM exclusion_rules") # 모든 규칙 삭제
        cursor.executemany("INSERT INTO exclusion_rules (description, rule_type, field, value, is_enabled) VALUES (?, ?, ?, ?, ?)", default_rules)
        conn.commit()
        conn.close()
        self.load_exclusion_rules()

    def _is_entry_excluded(self, entry, rules):
        """하나의 데이터 행(entry)이 제외 규칙에 해당하는지 검사"""
        for rule in rules:
            field_to_check = rule.get('field')
            if field_to_check not in entry:
                continue
            
            field_value = str(entry[field_to_check])
            rule_type = rule.get('type')
            rule_value = rule.get('value')
            
            try:
                if rule_type == "startswith" and field_value.startswith(rule_value): return True
                elif rule_type == "endswith" and field_value.endswith(rule_value): return True
                elif rule_type == "contains" and rule_value in field_value: return True
                elif rule_type == "equals" and field_value == rule_value: return True
                elif rule_type == "length" and len(field_value) > int(rule_value): return True
                elif rule_type == "regex" and re.search(rule_value, field_value): return True
            except Exception as e:
                print(f"규칙 적용 오류 (규칙 ID: {rule.get('id')}): {e}")
                continue # 규칙에 오류가 있으면 건너뜀
                
        return False
   
    def search_history(self):
        """번역 이력 검색"""
        search_text = self.history_search_var.get().strip()
        
        # 트리 초기화
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
            
        conn = sqlite3.connect(self.translation_db_path)
        cursor = conn.cursor()
        
        if search_text:
            cursor.execute("""
                SELECT created_at, string_id, kr_text, translation_method, status
                FROM translation_history
                WHERE kr_text LIKE ? OR string_id LIKE ?
                ORDER BY created_at DESC
                LIMIT 100
            """, (f"%{search_text}%", f"%{search_text}%"))
        else:
            cursor.execute("""
                SELECT created_at, string_id, kr_text, translation_method, status
                FROM translation_history
                ORDER BY created_at DESC
                LIMIT 100
            """)
            
        for row in cursor.fetchall():
            self.history_tree.insert("", "end", values=row)
            
        conn.close()

    def find_similar_translation(self, kr_text, threshold=0.9):
        """유사 번역 찾기를 번역 엔진에 위임"""
        return self.translation_engine.find_similar_translation(kr_text, threshold)

    def apply_glossary(self, kr_text):
        """용어집 적용을 번역 엔진에 위임"""
        return self.translation_engine.apply_glossary(kr_text)

#새로운 함수
    def analyze_text_quality(self, kr_text):
        """텍스트 품질 분석 (새 기능)"""
        result = self.text_processor.preprocess_for_translation(kr_text)
        
        analysis = {
            'length': result['metadata']['length'],
            'complexity': result['metadata']['complexity'],
            'estimated_difficulty': result['metadata']['estimated_difficulty'],
            'glossary_matches': len(result['glossary_matches']),
            'warnings': result['warnings'],
            'has_markup': result['metadata']['has_markup'],
            'has_variables': result['metadata']['has_variables']
        }
        
        return analysis

    def batch_analyze_translations(self):
        """현재 번역 대상들 일괄 분석 (새 기능)"""
        if not self.pending_translations:
            messagebox.showinfo("알림", "분석할 번역 대상이 없습니다.")
            return
        
        kr_texts = [trans['KR'] for trans in self.pending_translations]
        
        # 일괄 전처리
        results, batch_stats = self.text_processor.batch_preprocess(kr_texts)
        
        # 분석 결과 다이얼로그 표시
        analysis_summary = f"""📊 텍스트 분석 결과

전체 항목: {batch_stats['total']}개
성공 처리: {batch_stats['successful']}개
실패: {batch_stats['failed']}개

용어집 매칭: {batch_stats['glossary_hits']}개 항목
복잡한 텍스트: {batch_stats['complex_texts']}개 항목

💡 복잡한 텍스트가 많으면 번역 시간이 더 오래 걸릴 수 있습니다."""

        messagebox.showinfo("텍스트 분석 완료", analysis_summary)
        self.update_status(f"텍스트 분석 완료: {batch_stats['successful']}/{batch_stats['total']} 성공")

    def show_text_processing_stats(self):
        """텍스트 처리 통계 표시 (새 기능)"""
        stats = self.text_processor.get_processing_stats()
        
        stats_text = f"""📈 텍스트 처리 통계

처리된 텍스트: {stats['processed_count']}개
용어집 적용: {stats['glossary_applied']}개
태그 보호: {stats['tags_protected']}개
오류 감지: {stats['errors_detected']}개"""

        messagebox.showinfo("처리 통계", stats_text)
        

    def reset_processing_stats(self):
        """처리 통계 초기화"""
        self.text_processor.reset_stats()
        self.update_status("텍스트 처리 통계가 초기화되었습니다.")


# 6. 기존 번역 관련 함수에서 전처리 강화 옵션 추가

    def execute_translation_with_enhanced_preprocessing(self):
        """강화된 전처리와 함께 번역 실행"""
        # 먼저 일괄 분석 수행
        self.batch_analyze_translations()
        
        # 사용자 확인 후 번역 실행
        if messagebox.askyesno("번역 실행", "분석이 완료되었습니다. 번역을 시작하시겠습니까?"):
            self.execute_translation()

    #품질 관련
    def validate_current_batch(self):
        """현재 배치의 번역 품질 검증"""
        if not self.pending_translations:
            return
        
        self.update_status("번역 품질 검증 중...")
        
        # 번역된 항목들만 선별
        translated_items = [
            item for item in self.pending_translations
            if any(item["translations"].get(lang) for lang in self.VISIBLE_LANGS)
        ]
        
        if not translated_items:
            self.update_status("검증할 번역 항목이 없습니다.")
            return
        
        # 별도 스레드에서 검증 실행
        threading.Thread(
            target=self._batch_validation_thread, 
            args=(translated_items,), 
            daemon=True
        ).start()

    def _batch_validation_thread(self, translated_items):
        """(스레드) 배치 번역 품질 검증"""
        try:
            validation_pairs = []
            
            for item in translated_items:
                kr_text = item["KR"]
                
                # 각 언어별로 검증
                for lang in self.VISIBLE_LANGS:
                    translated_text = item["translations"].get(lang)
                    if translated_text and translated_text.strip():
                        validation_pairs.append((kr_text, translated_text))
            
            if not validation_pairs:
                self.root.after(0, lambda: self.update_status("검증할 번역이 없습니다."))
                return
            
            # 일괄 검증 실행
            self.root.after(0, lambda: self.update_status(f"{len(validation_pairs)}개 번역 품질 검증 중..."))
            
            batch_results, batch_stats = self.translation_validator.batch_validate_translations(
                validation_pairs, target_lang='EN'
            )
            
            # 결과 저장
            self.current_batch_quality = {
                'results': batch_results,
                'stats': batch_stats,
                'timestamp': datetime.now().isoformat(),
                'total_items': len(translated_items)
            }
            
            self.quality_history.append(self.current_batch_quality)
            
            # UI 업데이트
            self.root.after(0, lambda: self._update_ui_with_quality_results(batch_stats))
            
        except Exception as e:
            self.root.after(0, lambda: self.update_status(f"품질 검증 오류: {e}"))

    def _update_ui_with_quality_results(self, batch_stats):
        """품질 검증 결과로 UI 업데이트"""
        avg_score = batch_stats.get('average_score', 0.0)
        total = batch_stats.get('total', 0)
        
        # 품질 레벨별 개수
        excellent = batch_stats.get('excellent', 0)
        good = batch_stats.get('good', 0)
        acceptable = batch_stats.get('acceptable', 0)
        poor = batch_stats.get('poor', 0)
        
        # 상태 메시지 업데이트
        quality_summary = f"품질 검증 완료 - 평균: {avg_score:.2f} | 우수: {excellent} | 양호: {good} | 수용: {acceptable} | 부족: {poor}"
        self.update_status(quality_summary)
        
        # 품질이 낮은 항목이 많으면 경고
        if poor > total * 0.3:  # 30% 이상이 품질 부족
            self.root.after(
                500, 
                lambda: messagebox.showwarning(
                    "품질 경고", 
                    f"번역 품질이 낮은 항목이 {poor}개 발견되었습니다.\n\n"
                    f"전체 번역 품질 보고서를 확인하시겠습니까?",
                )
            )
            # 자동으로 품질 보고서 제안
            self.root.after(1000, self.offer_quality_report)

    def offer_quality_report(self):
        """품질 보고서 생성 제안"""
        if messagebox.askyesno(
            "품질 보고서", 
            "상세한 번역 품질 분석 보고서를 생성하시겠습니까?\n\n"
            "보고서에는 다음이 포함됩니다:\n"
            "• 품질 점수 분석\n"
            "• 공통 이슈 및 개선점\n"
            "• 재번역 권장 항목"
        ):
            self.generate_quality_report()

    def generate_quality_report(self):
        """품질 분석 보고서 생성 및 표시"""
        if not self.current_batch_quality.get('results'):
            messagebox.showinfo("보고서 없음", "분석할 품질 데이터가 없습니다.")
            return
        
        try:
            validation_results = self.current_batch_quality['results']
            
            # 보고서 생성
            report = self.translation_validator.create_quality_report(
                validation_results, 
                title=f"번역 품질 분석 보고서 - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            )
            
            # 보고서 다이얼로그 표시
            self.show_quality_report_dialog(report)
            
        except Exception as e:
            messagebox.showerror("보고서 생성 오류", f"품질 보고서 생성 중 오류: {e}")

    def show_quality_report_dialog(self, report):
        """품질 보고서 다이얼로그 표시"""
        dialog = tk.Toplevel(self.root)
        dialog.title("📊 번역 품질 분석 보고서")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 탭 구성
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True)
        
        # 요약 탭
        summary_frame = ttk.Frame(notebook)
        notebook.add(summary_frame, text="📈 요약")
        
        summary_text = self._format_summary_report(report)
        summary_widget = tk.Text(summary_frame, wrap="word", font=("맑은 고딕", 10))
        summary_widget.pack(fill="both", expand=True, padx=5, pady=5)
        summary_widget.insert("1.0", summary_text)
        summary_widget.config(state="disabled")
        
        # 상세 분석 탭
        details_frame = ttk.Frame(notebook)
        notebook.add(details_frame, text="🔍 상세 분석")
        
        details_text = self._format_details_report(report)
        details_widget = tk.Text(details_frame, wrap="word", font=("Consolas", 9))
        details_widget.pack(fill="both", expand=True, padx=5, pady=5)
        details_widget.insert("1.0", details_text)
        details_widget.config(state="disabled")
        
        # 권장사항 탭
        recommendations_frame = ttk.Frame(notebook)
        notebook.add(recommendations_frame, text="💡 권장사항")
        
        rec_text = self._format_recommendations_report(report)
        rec_widget = tk.Text(recommendations_frame, wrap="word", font=("맑은 고딕", 10))
        rec_widget.pack(fill="both", expand=True, padx=5, pady=5)
        rec_widget.insert("1.0", rec_text)
        rec_widget.config(state="disabled")
        
        # 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="📋 클립보드 복사", 
                command=lambda: self._copy_report_to_clipboard(report)).pack(side="left")
        ttk.Button(button_frame, text="💾 보고서 저장", 
                command=lambda: self._save_report_to_file(report)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="닫기", command=dialog.destroy).pack(side="right")

    def _format_summary_report(self, report):
        """요약 보고서 포맷팅"""
        summary = report.get('summary', {})
        
        text = f"""📊 번역 품질 분석 요약

    📈 전체 통계:
    - 분석 번역 수: {summary.get('total_translations', 0):,}개
    - 평균 품질 점수: {summary.get('average_score', 0.0):.3f}/1.000
    - 최고 점수: {summary.get('highest_score', 0.0):.3f}
    - 최저 점수: {summary.get('lowest_score', 0.0):.3f}

    🎯 품질 분포:"""
        
        distribution = summary.get('quality_distribution', {})
        for level, count in distribution.items():
            percentage = (count / summary.get('total_translations', 1)) * 100
            text += f"\n• {level}: {count}개 ({percentage:.1f}%)"
        
        return text

    def _format_details_report(self, report):
        """상세 보고서 포맷팅"""
        details = report.get('detailed_analysis', {})
        
        text = "🔍 상세 분석 결과\n\n"
        
        # 공통 이슈
        common_issues = details.get('common_issues', {})
        if common_issues:
            text += "⚠️ 자주 발생하는 이슈:\n"
            for issue_type, count in common_issues.items():
                text += f"• {issue_type}: {count}회\n"
            text += "\n"
        
        # 점수 세부 분석
        score_breakdown = details.get('score_breakdown', {})
        if score_breakdown:
            text += "📊 항목별 평균 점수:\n"
            for metric, avg_score in score_breakdown.items():
                text += f"• {metric}: {avg_score:.3f}\n"
            text += "\n"
        
        # 개선 영역
        improvement_areas = details.get('improvement_areas', [])
        if improvement_areas:
            text += "🎯 주요 개선 영역:\n"
            for area in improvement_areas[:5]:
                text += f"• {area['area']}: 현재 {area['current_score']:.3f} → 개선 필요 {area['improvement_needed']:.3f}\n"
        
        return text

    def _format_recommendations_report(self, report):
        """권장사항 보고서 포맷팅"""
        recommendations = report.get('recommendations', [])
        
        text = "💡 개선 권장사항\n\n"
        
        if recommendations:
            for i, rec in enumerate(recommendations, 1):
                text += f"{i}. {rec}\n\n"
        else:
            text += "현재 번역 품질이 양호하여 특별한 개선사항이 없습니다."
        
        # 일반적인 품질 개선 팁 추가
        text += "\n🎯 일반적인 품질 개선 팁:\n\n"
        text += "• 용어집을 정기적으로 업데이트하여 일관성을 향상시키세요.\n"
        text += "• 특수 태그 보호 설정을 확인하여 마크업이 올바르게 보존되는지 확인하세요.\n"
        text += "• 번역 엔진별 성능을 비교하여 최적의 엔진을 선택하세요.\n"
        text += "• 복잡한 텍스트는 문장 단위로 분할하여 번역 정확도를 높이세요.\n"
        
        return text

    def _copy_report_to_clipboard(self, report):
        """보고서를 클립보드에 복사"""
        try:
            full_report = f"{self._format_summary_report(report)}\n\n"
            full_report += f"{self._format_details_report(report)}\n\n"
            full_report += f"{self._format_recommendations_report(report)}"
            
            self.root.clipboard_clear()
            self.root.clipboard_append(full_report)
            messagebox.showinfo("복사 완료", "보고서가 클립보드에 복사되었습니다.")
        except Exception as e:
            messagebox.showerror("복사 실패", f"클립보드 복사 중 오류: {e}")

    def _save_report_to_file(self, report):
        """보고서를 파일로 저장"""
        try:
            filename = f"quality_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialvalue=filename
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"{self._format_summary_report(report)}\n\n")
                    f.write(f"{self._format_details_report(report)}\n\n")
                    f.write(f"{self._format_recommendations_report(report)}")
                
                messagebox.showinfo("저장 완료", f"보고서가 저장되었습니다:\n{file_path}")
        except Exception as e:
            messagebox.showerror("저장 실패", f"보고서 저장 중 오류: {e}")

    def show_validation_statistics(self):
        """검증 통계 표시"""
        stats = self.translation_validator.get_validation_stats()
        processor_stats = self.text_processor.get_processing_stats()
        
        stats_text = f"""📊 번역 처리 및 검증 통계

    🔧 텍스트 전처리:
    - 처리된 텍스트: {processor_stats['processed_count']}개
    - 용어집 적용: {processor_stats['glossary_applied']}개
    - 태그 보호: {processor_stats['tags_protected']}개
    - HTML 처리: {processor_stats['html_processed']}개
    - 오류 감지: {processor_stats['errors_detected']}개

    ✅ 품질 검증:
    - 총 검증 수행: {stats['total_validations']}개
    - 통과: {stats['passed_validations']}개
    - 실패: {stats['failed_validations']}개
    - 자동 수정 적용: {stats['auto_fixes_applied']}개
    - 재번역 제안: {stats['retranslation_suggested']}개"""
        
        messagebox.showinfo("처리 통계", stats_text)

    def reset_quality_tracking(self):
        """품질 추적 데이터 초기화"""
        if messagebox.askyesno("초기화 확인", "모든 품질 추적 데이터를 초기화하시겠습니까?"):
            self.quality_history.clear()
            self.current_batch_quality.clear()
            self.translation_validator.reset_stats()
            self.text_processor.reset_stats()
            self.update_status("품질 추적 데이터가 초기화되었습니다.")


if __name__ == "__main__":
    # TkinterDnD는 메인 창이 Tk()일 때 가장 잘 동작합니다.
    root = TkinterDnD.Tk()
    app = SmartTranslationManager(root)
    root.mainloop()