import os
import json
import hashlib
import pandas as pd
import sqlite3
import tkinter as tk
import time
from functools import lru_cache

# 공통 유틸리티 모듈 임포트
from utils.common_utils import FileUtils, PathUtils, HashUtils, DBUtils, ExcelUtils, logger
from utils.excel_utils import ExcelFileManager

# 기존 함수 유지하되 내부는 common_utils 호출로 대체
def load_cached_data(cache_path):
    """캐시 데이터 로드"""
    return FileUtils.load_cached_data(cache_path)

def save_cache(path, data):
    """캐시 데이터 저장"""
    return FileUtils.save_cache(path, data)

def get_file_mtime(path):
    """파일 수정 시간 가져오기"""
    return PathUtils.get_file_mtime(path)

def truncate(value, max_len=100):
    """긴 문자열 자르기"""
    return FileUtils.truncate(value, max_len)

def hash_paths(*paths):
    """경로 해시 생성"""
    return HashUtils.hash_paths(*paths)

def get_columns_from_db(db_path, table_name):
    """DB 테이블의 컬럼 정보 가져오기"""
    return DBUtils.get_columns_from_db(db_path, table_name)

def find_header_row(file_path, sheet_name, db_columns, xls=None):
    """엑셀 시트의 헤더 행 찾기"""
    return ExcelFileManager.find_header_row(file_path, sheet_name, db_columns, xls)

@lru_cache(maxsize=100)
def get_columns_from_db_cached(db_path, table_name):
    # 파일이 존재하는 경우에만 연결 시도
    if not os.path.exists(db_path):
        return []
        
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        # 테이블 이름을 따옴표로 감싸서 이스케이프 처리
        escaped_table_name = f'"{table_name}"'
        cur.execute(f"PRAGMA table_info({escaped_table_name})")
        columns = [row[1] for row in cur.fetchall()]
        conn.close()
        return columns
    except Exception as e:
        print(f"[DB 컬럼 조회 오류] {db_path}/{table_name}: {e}")
        return []

def analyze_excel_file(file_path, db_folder=None, with_debug=False):
    """엑셀 파일 분석"""
    return ExcelFileManager.analyze_excel_file(file_path, db_folder)

def analyze_excel_file_with_debug(file_path, db_folder_path, file_log):
    """디버그 정보를 포함한 엑셀 파일 분석"""
    result = ExcelFileManager.analyze_excel_file(file_path, db_folder_path)
    
    # 디버그 정보 추가 (파일 로그 업데이트)
    if file_log and isinstance(file_log, dict) and "steps" in file_log:
        # 디버그 정보 추가
        file_log["steps"]["excel_analysis"] = time.time()
        
    return result

def update_excel_cache(folder_path, cache_path=None, progress_callback=None):
    """엑셀 캐시 업데이트"""
    return ExcelFileManager.update_excel_cache(folder_path, cache_path, progress_callback)

def save_debug_log(path, data):
    """디버그 로그를 JSON 파일로 저장"""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logger.error(f"디버그 로그 저장 오류: {path} - {e}")

def save_debug_csv(path, file_logs):
    """디버그 정보를 CSV 파일로 저장"""
    import csv
    
    try:
        with open(path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            # 헤더 작성
            writer.writerow([
                "파일명", "상대경로", "총 처리시간(초)", 
                "엑셀 로드(초)", "시트 분석(초)", 
                "시트 수", "오류"
            ])
            
            # 파일별 정보 작성
            for rel_path, log in file_logs.items():
                writer.writerow([
                    log["file_name"],
                    log["rel_path"],
                    f"{log['total_time']:.4f}",
                    f"{log['steps'].get('excel_load', 0):.4f}",
                    f"{log['steps'].get('sheets_analysis', 0):.4f}",
                    len(log.get("sheets_detail", {})),
                    log.get("errors", "")
                ])
    except Exception as e:
        logger.error(f"CSV 로그 저장 오류: {path} - {e}")

# def update_db_cache(db_folder_path, folder_path, base_cache_dir=".cache"):
#     """DB 캐시 업데이트"""
#     if not os.path.exists(folder_path):
#         raise FileNotFoundError(f"[경고] 잘못된 Excel 경로: {folder_path}")

#     if not os.path.exists(db_folder_path):
#         raise FileNotFoundError(f"[경고] 잘못된 DB 경로: {db_folder_path}")

#     os.makedirs(base_cache_dir, exist_ok=True)
#     cache_id = hash_paths(folder_path, db_folder_path)
#     cache_dir = os.path.join(base_cache_dir, cache_id)
#     os.makedirs(cache_dir, exist_ok=True)

#     db_cache_path = os.path.join(cache_dir, "db_cache.json")
#     old_cache = load_cached_data(db_cache_path)
#     new_cache = {}

#     for file in os.listdir(db_folder_path):
#         if file.endswith(".db"):
#             db_path = os.path.join(db_folder_path, file)
#             mtime = get_file_mtime(db_path)
#             if file in old_cache and old_cache[file]["mtime"] == mtime:
#                 new_cache[file] = old_cache[file]
#                 continue
#             try:
#                 conn = sqlite3.connect(db_path)
#                 cursor = conn.cursor()
#                 cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
#                 tables = [row[0] for row in cursor.fetchall()]
#                 new_cache[file] = {"mtime": mtime, "tables": tables}
#                 conn.close()
#             except Exception as e:
#                 logger.error(f"DB 분석 실패: {db_path} - {e}")

#     save_cache(db_cache_path, new_cache)
#     return db_folder_path, folder_path, new_cache

def update_db_cache(db_folder_path, folder_path, base_cache_dir=".cache"):
    """DB 캐시 업데이트"""
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"[경고] 잘못된 Excel 경로: {folder_path}")

    if not os.path.exists(db_folder_path):
        raise FileNotFoundError(f"[경고] 잘못된 DB 경로: {db_folder_path}")

    os.makedirs(base_cache_dir, exist_ok=True)
    cache_id = hash_paths(folder_path, db_folder_path)
    cache_dir = os.path.join(base_cache_dir, cache_id)
    os.makedirs(cache_dir, exist_ok=True)

    db_cache_path = os.path.join(cache_dir, "db_cache.json")
    old_cache = load_cached_data(db_cache_path)
    new_cache = {}

    for file in os.listdir(db_folder_path):
        if file.endswith(".db"):
            db_path = os.path.join(db_folder_path, file)
            
            # 추가: 파일 크기 확인 - 0KB 파일은 건너뛰기
            if os.path.getsize(db_path) == 0:
                logger.error(f"0KB DB 파일 무시: {db_path}")
                continue
                
            mtime = get_file_mtime(db_path)
            if file in old_cache and old_cache[file]["mtime"] == mtime:
                new_cache[file] = old_cache[file]
                continue
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tables = [row[0] for row in cursor.fetchall()]
                new_cache[file] = {"mtime": mtime, "tables": tables}
                conn.close()
            except Exception as e:
                logger.error(f"DB 분석 실패: {db_path} - {e}")
                
                # 추가: 오류 발생 후 0KB 파일이 생성되었다면 삭제
                if os.path.exists(db_path) and os.path.getsize(db_path) == 0:
                    try:
                        os.remove(db_path)
                        logger.info(f"빈 DB 파일 삭제: {db_path}")
                    except Exception as del_e:
                        logger.error(f"빈 DB 파일 삭제 실패: {db_path} - {del_e}")

    save_cache(db_cache_path, new_cache)
    return db_folder_path, folder_path, new_cache

def build_table_sheet_index(folder_path, db_folder_path, base_cache_dir=".cache", status_callback=None):
    """테이블-시트 인덱스 생성"""
    from collections import defaultdict
    
    start_time = time.time()

    # 캐시 ID 및 경로 설정
    cache_id = hashlib.md5(f"{os.path.abspath(folder_path)}|{os.path.abspath(db_folder_path)}".encode("utf-8")).hexdigest()
    cache_dir = os.path.join(base_cache_dir, cache_id)
    os.makedirs(cache_dir, exist_ok=True)
    index_path = os.path.join(cache_dir, "table_sheet_index.json")
    
    # 기존 인덱스 로드 (없으면 빈 객체 생성)
    old_index_data = {}
    if os.path.exists(index_path):
        try:
            with open(index_path, "r", encoding="utf-8") as f:
                old_index_data = json.load(f)
        except Exception as e:
            logger.error(f"인덱스 로드 오류: {e}")
    
    old_index = old_index_data.get("index", {})
    old_mtimes = old_index_data.get("mtimes", {})
    
    # 결과를 저장할 딕셔너리
    index = defaultdict(list, {k: v for k, v in old_index.items()})
    file_mtimes = dict(old_mtimes)  # 기존 mtime 정보 복사
    changed_files = []

    # 변경된 파일만 처리하기 위해 모든 엑셀 파일의 수정 시간 확인
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") and not file.startswith("~$"):
                full_path = os.path.join(root, file)
                rel_path = os.path.normpath(os.path.relpath(full_path, folder_path))
                
                # 파일의 mtime 확인
                mtime = get_file_mtime(full_path)
                
                # 파일이 새로 추가되었거나 수정되었는지 확인
                if rel_path not in file_mtimes or file_mtimes[rel_path] != mtime:
                    changed_files.append((root, file))
                    file_mtimes[rel_path] = mtime
    
    # 변경된 파일만 처리
    if status_callback:
        status_callback(f"변경된 파일 {len(changed_files)}개 처리 중...")
    
    # 변경된 파일에 대해 기존 인덱스 항목 제거
    for root, file in changed_files:
        full_path = os.path.join(root, file)
        rel_path = os.path.normpath(os.path.relpath(full_path, folder_path))
        
        # 이 파일이 포함된 모든 테이블 항목 찾기
        tables_to_update = []
        for table_name, entries in index.items():
            updated_entries = [entry for entry in entries if entry[0] != rel_path]
            if len(updated_entries) != len(entries):
                tables_to_update.append(table_name)
                index[table_name] = updated_entries
    
    # 변경된 파일 처리
    for idx, (root, file) in enumerate(changed_files):
        if status_callback:
            status_callback(f"[{idx+1}/{len(changed_files)}] {file} 분석 중...")

        full_path = os.path.join(root, file)
        rel_path = os.path.relpath(full_path, folder_path)

        try:
            # 시트 이름만 빠르게 - with 문 사용하여 자동으로 닫히도록 함
            with pd.ExcelFile(full_path, engine="openpyxl") as xls:
                for sheet in xls.sheet_names:
                    if "@" not in sheet:
                        continue
                    table_name = sheet.split("@")[0]
                    index[table_name].append((rel_path, sheet))
        except Exception as e:
            logger.error(f"시트 인덱스 오류: {file} - {e}")

    elapsed = time.time() - start_time
    if status_callback:
        status_callback(f"테이블 인덱스 생성 완료 (변경: {len(changed_files)}개 파일, {elapsed:.2f}초)")

    # 인덱스와 mtime 정보를 함께 저장
    index_data = {
        "index": dict(index),  # defaultdict를 일반 dict로 변환
        "mtimes": file_mtimes,
        "last_updated": int(time.time())
    }
    
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(index_data, f, indent=2, ensure_ascii=False)

    return index_path

def get_tables_from_db_folder(db_folder_path):
    """DB 폴더에서 테이블 목록 가져오기"""
    if not os.path.exists(db_folder_path):
        return []
        
    return [os.path.splitext(f)[0] for f in os.listdir(db_folder_path) if f.endswith(".db")]