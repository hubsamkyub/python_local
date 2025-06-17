# 파일명: database_manager.py

import sqlite3
import os
import json
import re
from collections import defaultdict, Counter
import pandas as pd

class DatabaseManager:
    def __init__(self, db_path, legacy_db_path=None):
        """
        DatabaseManager를 초기화합니다.
        - db_path: 메인 데이터베이스 파일 경로
        - legacy_db_path: 이전 버전의 DB 경로 (선택 사항)
        """
        self.db_path = db_path
        self.legacy_db_path = legacy_db_path

    def _get_connection(self):
        """DB 연결 객체를 반환합니다."""
        return sqlite3.connect(self.db_path)

    def init_database(self):
        """데이터베이스와 모든 테이블을 초기화합니다."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            # translation_memory 테이블
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS translation_memory (
                    kr_text TEXT PRIMARY KEY, translations TEXT, usage_count INTEGER DEFAULT 1,
                    last_used TIMESTAMP DEFAULT CURRENT_TIMESTAMP, source TEXT, confidence REAL DEFAULT 1.0,
                    status TEXT DEFAULT 'consolidated', conflict_info TEXT
                )
            """)
            # glossary 테이블
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS glossary (
                    string_id TEXT PRIMARY KEY, kr TEXT, en TEXT, cn TEXT, tw TEXT, th TEXT, pt TEXT, es TEXT,
                    de TEXT, fr TEXT, jp TEXT, engine TEXT, contributor TEXT, update_at TEXT,
                    verified INTEGER, description TEXT
                )
            """)
            # translation_history 테이블
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS translation_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, string_id TEXT, kr_text TEXT, translations TEXT,
                    translation_method TEXT, status TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            # exclusion_rules 테이블
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS exclusion_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, description TEXT, rule_type TEXT NOT NULL,
                    field TEXT NOT NULL, value TEXT NOT NULL, is_enabled INTEGER NOT NULL DEFAULT 1,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

    # --- 데이터 조회 (Read) 메서드 ---

    def get_tm_entries(self, search_term=""):
        """번역 메모리에서 항목들을 조회하여 리스트로 반환합니다."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if search_term:
                cursor.execute("SELECT kr_text, translations FROM translation_memory WHERE kr_text LIKE ?", (f"%{search_term}%",))
            else:
                cursor.execute("SELECT kr_text, translations FROM translation_memory")
            return cursor.fetchall()

    def get_all_glossary(self):
        """DB에서 전체 용어집을 로드하여 딕셔너리로 반환합니다."""
        glossary_data = {}
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM glossary")
            for row in cursor.fetchall():
                if row["kr"]: # kr 값이 있는 경우만 추가
                    glossary_data[row["kr"]] = dict(row)
        return glossary_data

    def get_exclusion_rules(self, only_enabled=True):
        """DB에서 제외 규칙을 불러와 리스트로 반환합니다."""
        sql = "SELECT id, description, rule_type, field, value, is_enabled FROM exclusion_rules"
        if only_enabled:
            sql += " WHERE is_enabled = 1"
        sql += " ORDER BY id"
        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(sql)
            return cursor.fetchall()

    def get_translation_memory(self):
        """두 DB(레거시 포함)에서 데이터를 읽어와 번역 메모리(TM)를 구축하고 반환합니다."""
        translation_memory = {}
        # 1. 레거시 DB 로드
        if self.legacy_db_path and os.path.exists(self.legacy_db_path):
            try:
                with sqlite3.connect(self.legacy_db_path) as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT KR, EN, CN, TW, JP, DE, FR, TH, PT, ES FROM unique_texts")
                    for row in cursor.fetchall():
                        kr, translations = row[0], {
                            "EN": row[1] or "", "CN": row[2] or "", "TW": row[3] or "", "JP": row[4] or "",
                            "DE": row[5] or "", "FR": row[6] or "", "TH": row[7] or "", "PT": row[8] or "", "ES": row[9] or ""
                        }
                        if kr: translation_memory[kr] = translations
            except Exception as e:
                print(f"레거시 DB 로드 오류: {e}")

        # 2. 메인 DB 로드 및 병합
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT kr_text, translations FROM translation_memory")
                for kr, trans_json in cursor.fetchall():
                    if not kr or not trans_json: continue
                    translations = json.loads(trans_json)
                    if kr in translation_memory:
                        for lang, text in translations.items():
                            if text and not translation_memory[kr].get(lang):
                                translation_memory[kr][lang] = text
                    else:
                        translation_memory[kr] = translations
        except Exception as e:
            print(f"메인 DB 로드 오류: {e}")
            
        return translation_memory

    def get_db_tm_count(self):
        """DB의 TM 항목 수 확인"""
        try:
            conn = sqlite3.connect(self.translation_db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM translation_memory")
            count = cursor.fetchone()[0]
            conn.close()
            return count
        except:
            return 0


    # --- 데이터 수정 (Create, Update, Delete) 메서드 ---
    
    def update_translation_memory(self, pending_translations):
        """번역 작업이 완료된 항목들을 DB와 이력에 저장하고, 업데이트된 KR 목록을 반환합니다."""
        updated_krs = []
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for trans in pending_translations:
                if trans.get("translations"):
                    cursor.execute("""
                        INSERT OR REPLACE INTO translation_memory (kr_text, translations, source, confidence) 
                        VALUES (?, ?, ?, ?)
                    """, (trans["KR"], json.dumps(trans["translations"]), trans["method"], 1.0 if trans["method"] == "완전일치" else 0.8))
                    
                    cursor.execute("""
                        INSERT INTO translation_history (string_id, kr_text, translations, translation_method, status) 
                        VALUES (?, ?, ?, ?, ?)
                    """, (trans["STRING_ID"], trans["KR"], json.dumps(trans["translations"]), trans["method"], trans["status"]))
                    
                    updated_krs.append(trans["KR"])
            conn.commit()
        return updated_krs

    def add_exclusion_rule(self, desc, rule_type, field, value):
        """새로운 제외 규칙을 DB에 추가합니다."""
        with self._get_connection() as conn:
            conn.execute("INSERT INTO exclusion_rules (description, rule_type, field, value) VALUES (?, ?, ?, ?)",
                         (desc, rule_type, field, value))

    def delete_exclusion_rules(self, item_ids):
        """여러 제외 규칙을 DB에서 삭제합니다."""
        with self._get_connection() as conn:
            conn.executemany("DELETE FROM exclusion_rules WHERE id = ?", [(item_id,) for item_id in item_ids])

    def toggle_exclusion_rule(self, item_id):
        """선택된 규칙의 활성화/비활성화 상태를 변경합니다."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT is_enabled FROM exclusion_rules WHERE id = ?", (item_id,))
            current_status = cursor.fetchone()[0]
            new_status = 1 - current_status
            cursor.execute("UPDATE exclusion_rules SET is_enabled = ? WHERE id = ?", (new_status, item_id))

    def reset_default_exclusion_rules(self):
        """모든 규칙을 삭제하고 기본 규칙 세트로 초기화합니다."""
        default_rules = [
            ('#으로 시작하는 KR 제외', 'startswith', 'KR', '#', 1),
            ('cs_로 시작하는 STRING_ID 제외', 'startswith', 'STRING_ID', 'cs_', 1),
            ('KR의 길이가 20 이상인 항목 제외', 'length', 'KR', '20', 0),
            ('\\n\\n으로 시작하는 언어 제외', 'startswith', 'KR', '\\n\\n', 1),
            ('[@...] 형식 제외', 'regex', 'KR', r'^\[@.\]\w*$', 1)
        ]
        with self._get_connection() as conn:
            conn.execute("DELETE FROM exclusion_rules")
            conn.executemany("INSERT INTO exclusion_rules (description, rule_type, field, value, is_enabled) VALUES (?, ?, ?, ?, ?)", default_rules)

    def cleanup_tm_with_rules_thread(self, status_callback):
        """(스레드) 제외 규칙을 사용하여 마스터 TM을 정리하고, 처리 결과를 반환합니다."""
        try:
            status_callback("TM 정리 시작: 규칙 및 데이터 로드 중...")
            
            with self._get_connection() as conn:
                # 1. 활성화된 규칙 로드
                rules_cursor = conn.cursor()
                rules_cursor.execute("SELECT id, rule_type, field, value FROM exclusion_rules WHERE is_enabled = 1")
                active_rules = [{'id': r[0], 'type': r[1], 'field': r[2], 'value': r[3]} for r in rules_cursor.fetchall()]

                # 2. 마스터 TM 로드
                tm_cursor = conn.cursor()
                tm_cursor.execute("SELECT kr_text, translations FROM translation_memory")
                all_tm_entries = tm_cursor.fetchall()
            
            status_callback(f"TM 항목 검사 중... (총 {len(all_tm_entries)}개)")
            
            # 3. 삭제할 항목 식별
            kr_to_delete = []
            for kr_text, trans_json in all_tm_entries:
                entry_dict = {"KR": kr_text, **json.loads(trans_json)}
                if self._is_entry_excluded(entry_dict, active_rules):
                    kr_to_delete.append(kr_text)
            
            # 4. 식별된 항목 삭제
            if kr_to_delete:
                status_callback(f"삭제 작업 진행 중... ({len(kr_to_delete)}개 항목)")
                with self._get_connection() as conn:
                    conn.executemany("DELETE FROM translation_memory WHERE kr_text = ?", [(kr,) for kr in kr_to_delete])
            
            status_callback(f"TM 정리 완료! 총 {len(kr_to_delete)}개 항목이 삭제되었습니다.")
            return len(kr_to_delete) # 처리된 개수 반환
            
        except Exception as e:
            status_callback(f"TM 정리 중 오류 발생: {e}")
            return -1 # 오류 발생을 알림

    def _is_entry_excluded(self, entry, rules):
        """하나의 데이터 행(entry)이 제외 규칙에 해당하는지 검사합니다."""
        for rule in rules:
            field_value = str(entry.get(rule['field'], ''))
            rule_type = rule['type']
            rule_value = rule['value']
            try:
                if rule_type == "startswith" and field_value.startswith(rule_value): return True
                if rule_type == "endswith" and field_value.endswith(rule_value): return True
                if rule_type == "contains" and rule_value in field_value: return True
                if rule_type == "equals" and field_value == rule_value: return True
                if rule_type == "length" and len(field_value) > int(rule_value): return True
                if rule_type == "regex" and re.search(rule_value, field_value): return True
            except Exception as e:
                print(f"규칙 적용 오류 (규칙 ID: {rule['id']}): {e}")
                continue
        return False