import os
import sqlite3
import logging
import traceback

from utils.string_db_utils import search_all_string_dbs, load_or_build_string_db
from utils.cache_utils import hash_paths, update_excel_cache

class StringDBOperations:
    """문자열 DB 검색과 관리를 담당하는 클래스"""
    
    def __init__(self, folder_path, db_path, progress_update=None):
        """
        StringDBOperations 초기화
        
        Args:
            folder_path: 기본 폴더 경로
            db_path: DB 파일 경로
            progress_update: 진행 상태 업데이트 함수 (선택 사항)
        """
        self.folder_path = folder_path
        self.db_path = db_path
        self.progress_update = progress_update
    
    def set_progress_callback(self, callback):
        """진행 상태 업데이트 콜백 함수 설정"""
        self.progress_update = callback
    
    def _log_progress(self, message, success=True):
        """진행 상태 로깅 (콜백이 설정된 경우만)"""
        if self.progress_update:
            self.progress_update(message, success)
    
    def search_string_db(self, keyword, columns, match_exact=False, match_case=False, match_word=False, use_regex=False):
        """
        String DB에서 검색 실행
        
        Args:
            keyword: 검색 키워드
            columns: 검색할 컬럼 목록
            match_exact: 정확히 일치 여부
            match_case: 대소문자 구분 여부
            match_word: 단어 단위 검색 여부
            use_regex: 정규식 사용 여부
            
        Returns:
            검색 결과 목록
        """
        
        # 폴더 경로 확인
        folder = self.folder_path
        db_folder = self.db_path

        cache_id = hash_paths(folder, db_folder)
        db_dir = os.path.join(".cache", cache_id, "string_dbs")

        if not os.path.exists(db_dir):
            self._log_progress(f"검색용 DB 폴더가 존재하지 않습니다: {db_dir}", success=False)
            return []

        # 검색 실행
        results = search_all_string_dbs(
            keyword=keyword,
            columns=columns,
            db_dir=db_dir,
            match_exact=match_exact,
            match_case=match_case,
            match_word=match_word,
            use_regex=use_regex
        )
        
        return results
    
    def query_unique_string_db(self, kr_texts, langs=None):
        """
        고유 텍스트 DB에서 KR 값에 매칭되는 다국어 데이터 조회
        
        Args:
            kr_texts: KR 텍스트 목록
            langs: 검색할 언어 목록 (기본값: None - 모든 언어)
            
        Returns:
            매칭된 다국어 데이터 딕셔너리
        """
        # 고유 텍스트 DB 경로
        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unique_texts.db")
        print(f"DEBUG[query]: DB 연결 시도 - {db_path}")
        
        # DB 파일 존재 확인
        if not os.path.exists(db_path):
            # 프로그램 실행 경로에서 찾아보기
            alternative_path = "unique_texts.db"
            if os.path.exists(alternative_path):
                db_path = alternative_path
                print(f"DEBUG[query]: 대체 경로에서 DB 파일 찾음 - {db_path}")
            else:
                self._log_progress("고유 텍스트 DB 파일을 찾을 수 없습니다.", success=False)
                return None
        
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 테이블 목록 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            print(f"DEBUG[query]: DB 테이블 목록: {tables}")
            
            # 테이블 이름 선택
            table_name = None
            for possible_name in ["unique_strings", "unique_texts"]:
                if any(table[0] == possible_name for table in tables):
                    table_name = possible_name
                    print(f"DEBUG[query]: 테이블 찾음: {table_name}")
                    break
            
            if not table_name:
                if tables:
                    table_name = tables[0][0]
                    print(f"DEBUG[query]: 대체 테이블 사용: {table_name}")
                else:
                    print(f"DEBUG[query]: 테이블이 없음")
                    self._log_progress("DB에 테이블이 없습니다.", success=False)
                    conn.close()
                    return None
            
            # 테이블 구조 확인 (큰따옴표 사용 - 특히 SQLite에서 중요)
            query = f'PRAGMA table_info("{table_name}")'
            print(f"DEBUG[query]: 테이블 구조 쿼리: {query}")
            cursor.execute(query)
            columns_info = cursor.fetchall()
            print(f"DEBUG[query]: 테이블 컬럼 정보: {columns_info}")
            
            if not columns_info:
                # 다른 방법으로 컬럼 정보 가져오기 시도
                try:
                    print(f"DEBUG[query]: 다른 방법으로 컬럼 정보 가져오기 시도")
                    cursor.execute(f"SELECT * FROM {table_name} LIMIT 1")
                    columns = [desc[0] for desc in cursor.description]
                    print(f"DEBUG[query]: 대체 방법으로 찾은 컬럼: {columns}")
                except Exception as col_error:
                    print(f"DEBUG[query]: 컬럼 정보 가져오기 실패: {str(col_error)}")
                    self._log_progress("DB 컬럼 정보를 가져올 수 없습니다.", success=False)
                    columns = []
            else:
                columns = [col[1] for col in columns_info]
            
            print(f"DEBUG[query]: 컬럼 목록: {columns}")
            
            # 요청한 언어 컬럼이 테이블에 있는지 확인
            valid_langs = []
            if "KR" in columns:
                valid_langs.append("KR")
            
            # 선택한 언어 목록
            lang_selection = langs or ["ALL"]
            
            for lang in lang_selection:
                if lang == "ALL":
                    # ALL이 선택된 경우, 언어 컬럼 자동 감지
                    for possible_lang in ["EN", "CN", "TW", "TH", "PT", "ES", "DE", "FR", "JP"]:
                        if possible_lang in columns:
                            valid_langs.append(possible_lang)
                            print(f"DEBUG[query]: 언어 컬럼 찾음: {possible_lang}")
                elif lang in columns:
                    valid_langs.append(lang)
                    print(f"DEBUG[query]: 요청 언어 컬럼 찾음: {lang}")
                else:
                    print(f"DEBUG[query]: {lang} 컬럼이 테이블에 없음")
            
            print(f"DEBUG[query]: 유효한 언어 컬럼: {valid_langs}")
            
            if "KR" not in valid_langs:
                print(f"DEBUG[query]: KR 컬럼이 테이블에 없음")
                self._log_progress("DB에 KR 컬럼이 없습니다.", success=False)
                conn.close()
                return None
            
            if len(valid_langs) <= 1:  # KR만 있으면 의미 없음
                print(f"DEBUG[query]: KR 외에 다른 유효한 언어 컬럼이 없음")
                self._log_progress("DB에서 선택한 언어 컬럼을 찾을 수 없습니다.", success=False)
                conn.close()
                return None
                
            # 쿼리 구성 및 실행 (KR 텍스트 값으로 검색)
            replacement_data = {}
            
            # 중복 제거
            unique_kr_texts = list(set(kr_texts))
            placeholders = ", ".join(["?"] * len(unique_kr_texts))
            
            print(f"DEBUG[query]: 검색할 KR 텍스트: {unique_kr_texts}")
            
            # 선택한 컬럼만 조회
            select_cols = ", ".join([f'"{col}"' for col in valid_langs])
            query = f'SELECT {select_cols} FROM "{table_name}" WHERE "KR" IN ({placeholders})'
            print(f"DEBUG[query]: 실행할 쿼리: {query}")
            print(f"DEBUG[query]: 쿼리 파라미터: {unique_kr_texts}")
            
            cursor.execute(query, unique_kr_texts)
            rows = cursor.fetchall()
            print(f"DEBUG[query]: 쿼리 결과 행 수: {len(rows)}")
            
            # 결과 매핑
            field_names = valid_langs
            
            for row in rows:
                kr_text = row[field_names.index("KR")]
                print(f"DEBUG[query]: 매칭된 KR 텍스트: {kr_text}")
                data = {}
                
                for lang in valid_langs:
                    if lang != "KR" and lang in field_names:
                        idx = field_names.index(lang)
                        if idx < len(row):
                            data[lang] = row[idx]
                            print(f"DEBUG[query]: {lang} 값: {row[idx]}")
                
                # KR 텍스트 기준으로 저장
                replacement_data[kr_text] = data
            
            print(f"DEBUG[query]: 최종 치환 데이터: {replacement_data}")
            conn.close()
            return replacement_data
            
        except Exception as e:
            print(f"DEBUG[query]: DB 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            self._log_progress(f"DB 조회 중 오류 발생: {str(e)}", success=False)
            if 'conn' in locals():
                conn.close()
            return None
    
    def update_string_db_and_refresh_search(self, folder_path=None, db_path=None, keyword=None):
        """
        String DB 최신화 후 검색 결과 갱신
        
        Args:
            folder_path: 폴더 경로 (기본값: None - 인스턴스 변수 사용)
            db_path: DB 경로 (기본값: None - 인스턴스 변수 사용)
            keyword: 검색 키워드 (기본값: None - 검색 갱신 안함)
        
        Returns:
            성공 여부
        """
        folder_path = folder_path or self.folder_path
        db_path = db_path or self.db_path
            
        try:
            # 1. 캐시 경로 계산
            cache_id = hash_paths(folder_path, db_path)
            cache_dir = os.path.join(".cache", cache_id)
            excel_cache_path = os.path.join(cache_dir, "excel_cache.json")
            db_folder = os.path.join(cache_dir, "string_dbs")
            
            # 2. 엑셀 캐시 업데이트 (데이터 변경되었으므로)
            self._log_progress("엑셀 캐시 업데이트 중...", success=True)
            updated_cache = update_excel_cache(folder_path, db_path)
            
            # 3. String DB 업데이트
            self._log_progress("String DB 업데이트 중...", success=True)
            load_or_build_string_db(excel_cache_path, db_folder, folder_path)
            
            # 4. 검색 결과가 필요한 경우 재검색 (여기서는 결과를 반환하지만 UI 업데이트는 호출자에서 수행)
            if keyword:
                self._log_progress(f"'{keyword}' 검색 결과 갱신 중.", success=True)
                # 검색에 필요한 정보는 호출자에서 처리 (검색 모드 전환 등)
            
            return True, updated_cache
        
        except Exception as e:
            error_msg = f"DB 최신화 중 오류 발생: {str(e)}"
            print(f"ERROR: {error_msg}")
            self._log_progress(error_msg, success=False)
            traceback.print_exc()
            return False, None

    def get_db_columns(self, table_name):
        """
        DB에서 테이블의 컬럼 목록을 가져옵니다.
        
        Args:
            table_name: 테이블 이름
            
        Returns:
            컬럼 목록
        """
        # 통합된 경로 사용
        db_file = os.path.join(self.folder_path, f"{table_name}.db")
        
        if not os.path.exists(db_file):
            return []
        
        try:
            conn = sqlite3.connect(db_file)
            cur = conn.cursor()
            cur.execute(f"PRAGMA table_info({table_name})")
            columns = [row[1] for row in cur.fetchall()]
            conn.close()
            return columns
        except Exception as e:
            print(f"컬럼 정보 가져오기 실패: {str(e)}")
            return []