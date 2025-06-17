import sqlite3
import json
import os
from collections import defaultdict

class DatabaseManager:
    def __init__(self, db_path, legacy_db_path=None):
        self.db_path = f"{db_path}.db"
        self.legacy_db_path = legacy_db_path
        self.init_database()

    def init_database(self):
        """데이터베이스와 필요한 모든 테이블을 초기화합니다. (CREATE IF NOT EXISTS 사용)"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # 1. translation_memory 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS translation_memory (
                kr_text TEXT PRIMARY KEY, translations TEXT, source TEXT,
                status TEXT, conflict_info TEXT, last_updated TEXT
            )""")
            # 2. glossary 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS glossary (
                kr TEXT PRIMARY KEY,
                category TEXT, 
                en TEXT, cn TEXT, tw TEXT, th TEXT, pt TEXT, es TEXT, 
                de TEXT, fr TEXT, jp TEXT, engine TEXT, contributor TEXT, 
                update_at TEXT, verified INTEGER, description TEXT, string_id TEXT
            )""")
            # 3. exclusion_rules 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS exclusion_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT, description TEXT, rule_type TEXT,
                field TEXT, value TEXT, is_enabled INTEGER DEFAULT 1
            )""")
            # 4. speakers 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS speakers (
                name TEXT PRIMARY KEY, gender TEXT, tone TEXT, style TEXT, reference_count INTEGER
            )""")
            # 5. reference_datasets 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS reference_datasets (
                id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, description TEXT,
                source_type TEXT, source_path TEXT, target_language TEXT, 
                total_speakers INTEGER, total_sentences INTEGER,
                created_at TEXT, last_used TEXT
            )""")
            # 6. reference_translations 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS reference_translations (
                id INTEGER PRIMARY KEY AUTOINCREMENT, dataset_id INTEGER, speaker_name TEXT,
                kr_text TEXT, translated_text TEXT,
                FOREIGN KEY (dataset_id) REFERENCES reference_datasets (id)
            )""")
            # 7. translation_history 테이블
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS translation_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, string_id TEXT,
                kr_text TEXT, translation_method TEXT, status TEXT, details TEXT
            )""")
            conn.commit()
            print(f"✅ 데이터베이스 '{self.db_path}' 초기화 완료.")

    # === [수정된 부분 시작] ===
    # 각 'get' 함수에 try-except 구문을 추가하여 "no such table" 오류를 처리합니다.
    
    def get_translation_memory(self):
        """번역 메모리(TM)를 로드합니다. 테이블이 없으면 빈 dict를 반환합니다."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT kr_text, translations FROM translation_memory")
                tm = {row[0]: json.loads(row[1]) for row in cursor.fetchall()}
                return tm
        except sqlite3.OperationalError as e:
            if "no such table" in str(e):
                print("⚠️ 경고: translation_memory 테이블이 없어 빈 TM을 반환합니다.")
                return {}
            else:
                raise e

    def get_all_glossary(self):
        """용어집 전체를 로드합니다. 테이블이 없으면 빈 dict를 반환합니다."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM glossary")
                # kr을 키로 사용하는 딕셔너리 반환
                return {row['kr']: dict(row) for row in cursor.fetchall() if row['kr']}
        except sqlite3.OperationalError as e:
            if "no such table" in str(e):
                print("⚠️ 경고: glossary 테이블이 없어 빈 용어집을 반환합니다.")
                return {}
            else:
                raise e

    def get_exclusion_rules(self, only_enabled=True):
        """제외 규칙을 로드합니다. 테이블이 없으면 빈 list를 반환합니다."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                query = "SELECT id, description, rule_type, field, value, is_enabled FROM exclusion_rules"
                if only_enabled:
                    query += " WHERE is_enabled = 1"
                cursor.execute(query)
                return cursor.fetchall()
        except sqlite3.OperationalError as e:
            if "no such table" in str(e):
                print("⚠️ 경고: exclusion_rules 테이블이 없어 빈 규칙 목록을 반환합니다.")
                return []
            else:
                raise e

    def get_conflicts(self):
        """충돌 항목들을 로드합니다. 테이블이 없으면 빈 list를 반환합니다."""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT kr_text, translations, conflict_info FROM translation_memory WHERE status = 'conflict'")
                return cursor.fetchall()
        except sqlite3.OperationalError as e:
            if "no such table" in str(e):
                print("⚠️ 경고: translation_memory 테이블이 없어 충돌 항목을 조회할 수 없습니다.")
                return []
            else:
                raise e

    # === [수정된 부분 끝] ===

    def update_translation_memory(self, translated_items):
        """번역된 항목들로 TM을 업데이트/삽입합니다."""
        if not translated_items:
            return []
        
        updated_krs = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            for item in translated_items:
                kr = item.get("KR")
                if not kr: continue
                
                translations_json = json.dumps(item.get("translations", {}))
                source = item.get("method", "N/A")
                status = item.get("status", "[완료]").strip("[]")
                
                cursor.execute("""
                    INSERT INTO translation_memory (kr_text, translations, source, status, last_updated)
                    VALUES (?, ?, ?, ?, datetime('now', 'localtime'))
                    ON CONFLICT(kr_text) DO UPDATE SET
                        translations = excluded.translations,
                        source = excluded.source,
                        status = excluded.status,
                        last_updated = datetime('now', 'localtime')
                """, (kr, translations_json, source, status))
                updated_krs.append(kr)
            conn.commit()
        return updated_krs
        
    def _get_connection(self):
        """DB 연결 객체를 반환합니다."""
        return sqlite3.connect(self.db_path)

    def get_tm_entries(self, search_term=""):
        """번역 메모리에서 항목들을 조회하여 리스트로 반환합니다."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if search_term:
                cursor.execute("SELECT kr_text, translations FROM translation_memory WHERE kr_text LIKE ?", (f"%{search_term}%",))
            else:
                cursor.execute("SELECT kr_text, translations FROM translation_memory")
            return cursor.fetchall()

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