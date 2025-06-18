import os
import json
import sqlite3
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from ..utils.analyzer_snippet import apply_typecode_mappings

def get_file_mtime(path):
    try:
        return int(os.path.getmtime(path))
    except:
        return -1

def analyze_table_relationships(db_folder, output_path=None, progress_callback=None, manual_relationships=None, excel_data=None, typecode_mappings=None):
    """
    데이터베이스 폴더에서 테이블 간의 관계를 분석합니다.
    
    Args:
        db_folder: DB 파일들이 있는 폴더 경로
        output_path: 결과를 저장할 파일 경로 (기본값: None, 반환만 하고 저장 안 함)
        progress_callback: 진행 상황을 보고할 콜백 함수
        manual_relationships: 수동으로 추가한 관계 정보
        excel_data: 엑셀에서 추출한 테이블/컬럼 정보 (딕셔너리)
        typecode_mappings: TSV에서 로드한 타입코드 기반 관계 정보
        
    Returns:
        딕셔너리 형태의 테이블 관계 매핑
    """
    # 결과를 저장할 딕셔너리
    relationships = {}
    
    # 모든 테이블과 컬럼 정보 수집
    tables_info = {}
    db_files = [f for f in os.listdir(db_folder) if f.endswith('.db')]
    
    # 열거형(Enum) 타입 목록
    enum_types = [
        "RewardType", "HeroCategoryType", "CostType", "RewardGroupType", 
        "QuestType", "QuestMissionType", "ConditionType", "StageType"
    ]
    
    for idx, db_file in enumerate(db_files):
        if progress_callback:
            progress_callback(f"스키마 분석 중... ({idx+1}/{len(db_files)}) {db_file}")
        
        db_path = os.path.join(db_folder, db_file)
        table_name = os.path.splitext(db_file)[0]
        
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 테이블 컬럼 정보 가져오기
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns_info = cursor.fetchall()
            columns = [row[1] for row in columns_info]
            column_types = {row[1]: row[2] for row in columns_info}  # 컬럼 타입 정보 저장
            
            if not columns:
                continue
                
            tables_info[table_name] = {
                "columns": columns,
                "column_types": column_types,
                "primary_key": None,  # 기본키는 일단 None으로 설정
                "foreign_keys": []    # 외래키 목록
            }
            
            # 각 컬럼 타입 정보 수집 (ID 컬럼 찾기)
            for col in columns:
                if col.lower() == 'id' or col.endswith('ID') or col.endswith('Id'):
                    tables_info[table_name]["primary_key"] = col
                    break
            
            conn.close()
        except Exception as e:
            print(f"[DB 분석 오류] {db_file}: {e}")
    
    # 엑셀 데이터 활용하여 테이블 정보 보강
    if excel_data and 'tables' in excel_data:
        for table_data in excel_data['tables']:
            table_name = table_data.get('table_name')
            if table_name in tables_info:
                # 엑셀에서 정의된 컬럼 관계 정보 추가
                for col_info in table_data.get('columns', []):
                    col_name = col_info.get('column_name')
                    related_table = col_info.get('related_table')
                    related_column = col_info.get('related_column')
                    
                    if col_name and related_table and related_column:
                        if "foreign_keys" not in tables_info[table_name]:
                            tables_info[table_name]["foreign_keys"] = []
                        
                        tables_info[table_name]["foreign_keys"].append({
                            "column": col_name,
                            "foreign_table": related_table,
                            "foreign_column": related_column
                        })
    
    # 패턴 기반 관계 분석
    for table_name, table_info in tables_info.items():
        if progress_callback:
            progress_callback(f"관계 분석 중... {table_name}")
            
        relationships[table_name] = {}
        
        # 1. 엑셀에서 정의된 외래키 관계 처리
        for fk in table_info.get("foreign_keys", []):
            target_table = fk["foreign_table"]
            target_column = fk["foreign_column"]
            source_column = fk["column"]
            
            if target_table not in relationships[table_name]:
                relationships[table_name][target_table] = []
                
            # 중복 검사
            is_duplicate = False
            for existing_rel in relationships[table_name].get(target_table, []):
                if existing_rel.get("source_column") == source_column and existing_rel.get("target_column") == target_column:
                    is_duplicate = True
                    break
                    
            if not is_duplicate:
                relationships[table_name][target_table].append({
                    "source_column": source_column,
                    "target_column": target_column,
                    "relation_type": "excel_foreign_key"
                })
        
        # 2. Enum 타입 컬럼 분석 (TypeCode 기반 관계)
        for column in table_info["columns"]:
            # Enum 타입 컬럼 찾기 (예: RewardType, QuestType 등)
            for enum_type in enum_types:
                if column == enum_type:
                    # 해당 테이블을 Enum 타입의 참조 테이블로 설정
                    for other_table, other_info in tables_info.items():
                        if other_table == table_name:
                            continue
                            
                        # 다른 테이블의 컬럼 중 Type 컬럼과 Value/Code 컬럼 쌍 찾기
                        for other_col in other_info["columns"]:
                            if other_col == "Type" or other_col.endswith("Type"):
                                # Value/Code 컬럼 찾기 시도
                                value_col = None
                                for val_candidate in ["Value", "Code", "ID", "Id"]:
                                    if val_candidate in other_info["columns"]:
                                        value_col = val_candidate
                                        break
                                
                                if value_col:
                                    if other_table not in relationships:
                                        relationships[other_table] = {}
                                    if table_name not in relationships[other_table]:
                                        relationships[other_table][table_name] = []
                                    
                                    # 중복 검사
                                    is_duplicate = False
                                    for existing_rel in relationships[other_table].get(table_name, []):
                                        if existing_rel.get("source_column") == value_col and existing_rel.get("target_column") == "ID" and existing_rel.get("filter_column") == other_col:
                                            is_duplicate = True
                                            break
                                    
                                    if not is_duplicate:
                                        relationships[other_table][table_name].append({
                                            "source_column": value_col,
                                            "target_column": "ID" if "ID" in table_info["columns"] else "Id",
                                            "filter_column": other_col,
                                            "filter_value": enum_type,
                                            "relation_type": "enum_typecode"
                                        })
        
        # 3. 기존 패턴 기반 분석 (ID 컬럼 등)
        for column in table_info["columns"]:
            # ID 컬럼 분석 (예: HeroTemplateID, Item_ID 등)
            if column.endswith('ID') or column.endswith('_id') or column.endswith('Id'):
                # 테이블 이름 패턴 추출 (예: HeroTemplate, Item 등)
                possible_tables = []
                
                # 패턴 1: TableNameID
                pattern1 = re.match(r'(.+?)(?:ID|Id)$', column)
                if pattern1:
                    possible_tables.append(pattern1.group(1))
                
                # 패턴 2: TableName_ID
                pattern2 = re.match(r'(.+?)_(?:ID|Id)$', column)
                if pattern2:
                    possible_tables.append(pattern2.group(1))
                
                # 패턴 3: FK_TableName
                pattern3 = re.match(r'FK_(.+)', column)
                if pattern3:
                    possible_tables.append(pattern3.group(1))
                
                # 패턴 4: 짧은 이름 (예: "ID" -> 해당 테이블의 주 키)
                if column.lower() == 'id':
                    possible_tables.append(table_name)
                
                # 발견된 테이블이 실제로 존재하는지 확인
                for possible_table in possible_tables:
                    # 정확히 일치하는 테이블이 있는 경우
                    if possible_table in tables_info:
                        target_table = possible_table
                        target_column = tables_info[target_table]["primary_key"] or "ID"
                        
                        # 중복 검사
                        is_duplicate = False
                        if target_table in relationships[table_name]:
                            for existing_rel in relationships[table_name][target_table]:
                                if existing_rel.get("source_column") == column and existing_rel.get("target_column") == target_column:
                                    is_duplicate = True
                                    break
                        
                        # 관계 정보 저장 (중복이 아닌 경우만)
                        if not is_duplicate:
                            if target_table not in relationships[table_name]:
                                relationships[table_name][target_table] = []
                                
                            relationships[table_name][target_table].append({
                                "source_column": column,
                                "target_column": target_column,
                                "relation_type": "foreign_key"
                            })
                    
                    # 유사한 테이블이 있는 경우 (예: HeroTemplate -> Hero)
                    else:
                        for existing_table in tables_info.keys():
                            # 테이블명이 서로 포함 관계인 경우
                            if possible_table in existing_table or existing_table in possible_table:
                                similarity_score = len(set(possible_table.lower()) & set(existing_table.lower())) / max(len(possible_table), len(existing_table))
                                
                                # 유사도가 높은 경우 (70% 이상)
                                if similarity_score >= 0.7:
                                    target_table = existing_table
                                    target_column = tables_info[target_table]["primary_key"] or "ID"
                                    
                                    # 중복 검사
                                    is_duplicate = False
                                    if target_table in relationships[table_name]:
                                        for existing_rel in relationships[table_name][target_table]:
                                            if existing_rel.get("source_column") == column and existing_rel.get("target_column") == target_column:
                                                is_duplicate = True
                                                break
                                    
                                    # 관계 정보 저장 (중복이 아닌 경우만)
                                    if not is_duplicate:
                                        if target_table not in relationships[table_name]:
                                            relationships[table_name][target_table] = []
                                            
                                        relationships[table_name][target_table].append({
                                            "source_column": column,
                                            "target_column": target_column,
                                            "relation_type": "possible_foreign_key",
                                            "similarity": f"{similarity_score:.2f}"
                                        })
    
    # 수동으로 추가한 관계 정보 처리
    if manual_relationships:
        for item in manual_relationships:
            source_table = item["table"]
            source_column = item["column"]
            target_info = item["related_columns"]
            
            # 관계 정보 파싱 (예: "HeroID:Hero:ID, ItemID:Item:ID")
            for relation in target_info.split(','):
                parts = relation.strip().split(':')
                if len(parts) >= 2:
                    target_column = parts[0].strip()
                    target_table = parts[1].strip()
                    
                    # 대상 테이블 컬럼 (기본값: "ID")
                    target_table_column = parts[2].strip() if len(parts) >= 3 else "ID"
                    
                    # 관계 정보 저장
                    if source_table not in relationships:
                        relationships[source_table] = {}
                    
                    if target_table not in relationships[source_table]:
                        relationships[source_table][target_table] = []
                    
                    # 중복 추가 방지
                    existing = False
                    for rel in relationships[source_table][target_table]:
                        if rel["source_column"] == source_column and rel["target_column"] == target_table_column:
                            existing = True
                            break
                    
                    if not existing:
                        relationships[source_table][target_table].append({
                            "source_column": source_column,
                            "target_column": target_table_column,
                            "relation_type": "manual_foreign_key"
                        })
    
    # 타입코드 매핑 정보 처리
    if typecode_mappings:
        for mapping in typecode_mappings:
            src_tbl = mapping.get("source_table")
            src_col = mapping.get("source_column")
            tgt_tbl = mapping.get("target_table")
            tgt_col = mapping.get("target_column")
            filter_col = mapping.get("filter_column")
            filter_val = mapping.get("filter_value")
            rel_type = mapping.get("relation_type", "타입코드 외래키")
            
            if not src_tbl or not src_col or not tgt_tbl or not tgt_col:
                continue
            
            # 관계 정보 저장
            if src_tbl not in relationships:
                relationships[src_tbl] = {}
            
            if tgt_tbl not in relationships[src_tbl]:
                relationships[src_tbl][tgt_tbl] = []
            
            # 매핑 정보 생성
            relation = {
                "source_column": src_col,
                "target_column": tgt_col,
                "relation_type": "typecode_foreign_key"
            }
            
            # 필터 정보가 있는 경우 추가
            if filter_col and filter_val:
                relation["filter_column"] = filter_col
                relation["filter_value"] = filter_val
            
            # 중복 검사
            is_duplicate = False
            for existing_rel in relationships[src_tbl][tgt_tbl]:
                match = True
                for key, value in relation.items():
                    if key not in existing_rel or existing_rel[key] != value:
                        match = False
                        break
                
                if match:
                    is_duplicate = True
                    break
            
            # 중복이 아닌 경우에만 추가
            if not is_duplicate:
                relationships[src_tbl][tgt_tbl].append(relation)
    
    # 결과 저장
    if output_path:
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(relationships, f, indent=2, ensure_ascii=False)
            print(f"테이블 관계 매핑 저장 완료: {output_path}")
        except Exception as e:
            print(f"[저장 오류] {e}")
    
    return relationships

def apply_typecode_mappings(relationships, typecode_mappings):
    for item in typecode_mappings:
        src_tbl = item["source_table"]
        tgt_tbl = item["target_table"]
        tgt_col = item["target_column"]
        src_col = item["source_column"]
        rel = {
            "source_column": src_col,
            "target_column": tgt_col,
            "relation_type": "typecode_foreign_key",
            "filter_column": item.get("filter_column"),
            "filter_value": item.get("filter_value")
        }

        if src_tbl not in relationships:
            relationships[src_tbl] = {}
        if tgt_tbl not in relationships[src_tbl]:
            relationships[src_tbl][tgt_tbl] = []

        if rel not in relationships[src_tbl][tgt_tbl]:
            relationships[src_tbl][tgt_tbl].append(rel)

    return relationships

def load_excel_data(excel_path):
    """
    엑셀 파일에서 테이블 및 컬럼 정보를 로드합니다.
    
    Args:
        excel_path: 엑셀 파일 경로
        
    Returns:
        딕셔너리 형태의 엑셀 데이터와 관계 매핑 정보
    """
    try:
        import pandas as pd
        
        # 엑셀 파일 로드
        excel_data = {"tables": [], "enums": [], "relationships": []}
        
        # 관계 정보 시트 로드 (새로운 형식)
        try:
            # 관계 정보가 포함된 시트 로드 (Sheet1 또는 다른 이름일 수 있음)
            relationships_df = pd.read_excel(excel_path, sheet_name="Sheet1")
            
            # 필요한 컬럼이 있는지 확인
            expected_columns = ["소스 테이블", "소스 컬럼", "타겟 테이블", "타겟 컬럼", "조건 컬럼", "조건 값", "관계 타입"]
            required_columns = ["소스 테이블", "소스 컬럼", "타겟 테이블", "타겟 컬럼"]
            
            # 필수 컬럼이 모두 있는지 확인
            if all(col in relationships_df.columns for col in required_columns):
                for _, row in relationships_df.iterrows():
                    # 필수 값이 없는 행은 건너뛰기
                    if pd.isna(row["소스 테이블"]) or pd.isna(row["소스 컬럼"]) or \
                       pd.isna(row["타겟 테이블"]) or pd.isna(row["타겟 컬럼"]):
                        continue
                    
                    src_tbl = str(row["소스 테이블"]).strip()
                    src_col = str(row["소스 컬럼"]).strip()
                    tgt_tbl = str(row["타겟 테이블"]).strip()
                    tgt_col = str(row["타겟 컬럼"]).strip()
                    
                    # 선택적 컬럼 처리
                    filter_col = str(row["조건 컬럼"]).strip() if "조건 컬럼" in row and not pd.isna(row["조건 컬럼"]) else ""
                    filter_val = str(row["조건 값"]).strip() if "조건 값" in row and not pd.isna(row["조건 값"]) else ""
                    rel_type = str(row["관계 타입"]).strip() if "관계 타입" in row and not pd.isna(row["관계 타입"]) else "외래키"
                    
                    # 관계 매핑 정보 생성
                    relationship = {
                        "source_table": src_tbl,
                        "source_column": src_col,
                        "target_table": tgt_tbl,
                        "target_column": tgt_col,
                        "relation_type": rel_type
                    }
                    
                    # 필터 조건이 있는 경우 추가
                    if filter_col and filter_val:
                        relationship["filter_column"] = filter_col
                        relationship["filter_value"] = filter_val
                    
                    # relationships 목록에 추가
                    excel_data["relationships"].append(relationship)
                    
                    # 수동 관계 추가를 위한 형식으로도 변환
                    # related_columns 형식: SourceColumn:TargetTable:TargetColumn
                    related_info = f"{src_col}:{tgt_tbl}:{tgt_col}"
                    if filter_col and filter_val:
                        related_info += f":{filter_col}:{filter_val}"
                    
                    manual_relationship = {
                        "table": src_tbl,
                        "column": src_col,
                        "related_columns": related_info,
                        "source": "excel",
                        "relation_type": rel_type
                    }
                    
                    # 테이블 정보에도 추가
                    found = False
                    for table in excel_data["tables"]:
                        if table["table_name"] == src_tbl:
                            found = True
                            column_exists = False
                            for column in table["columns"]:
                                if column["column_name"] == src_col:
                                    column_exists = True
                                    if "related_table" not in column:
                                        column["related_table"] = tgt_tbl
                                        column["related_column"] = tgt_col
                                    break
                            
                            if not column_exists:
                                table["columns"].append({
                                    "column_name": src_col,
                                    "related_table": tgt_tbl,
                                    "related_column": tgt_col
                                })
                            break
                    
                    if not found:
                        excel_data["tables"].append({
                            "table_name": src_tbl,
                            "columns": [{
                                "column_name": src_col,
                                "related_table": tgt_tbl,
                                "related_column": tgt_col
                            }]
                        })
            
        except Exception as e:
            print(f"[관계 시트 로드 오류] {e}")
        
        # enum 시트 로드는 기존과 동일
        try:
            enums_df = pd.read_excel(excel_path, sheet_name="enum")
            enum_types = {}
            
            for _, row in enums_df.iterrows():
                enum_type = row.get("EnumType", "")
                if pd.isna(enum_type):
                    continue
                    
                enum_type = str(enum_type).strip()
                enum_value = row.get("Value", 0)
                enum_name = row.get("Name", "")
                
                if enum_type not in enum_types:
                    enum_types[enum_type] = []
                
                enum_types[enum_type].append({
                    "value": enum_value,
                    "name": str(enum_name).strip() if not pd.isna(enum_name) else ""
                })
            
            # 딕셔너리를 리스트로 변환
            for enum_type, values in enum_types.items():
                excel_data["enums"].append({
                    "type": enum_type,
                    "values": values
                })
        except Exception as e:
            print(f"[enum 시트 로드 오류] {e}")
        
        return excel_data
    except Exception as e:
        print(f"[엑셀 로드 오류] {e}")
        return None



def load_typecode_mappings_from_tsv(tsv_path):
    """
    TSV 파일에서 타입코드 기반 관계 매핑 정보를 로드합니다.
    
    Args:
        tsv_path: TSV 파일 경로
        
    Returns:
        타입코드 매핑 목록
    """
    import csv
    
    mappings = []
    
    try:
        with open(tsv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter='\t')
            headers = next(reader)  # 헤더 행 읽기
            
            for row in reader:
                if len(row) < 7:  # 최소 7개 컬럼 필요
                    continue
                    
                src_tbl = row[0].strip()
                src_col = row[1].strip()
                tgt_tbl = row[2].strip()
                tgt_col = row[3].strip()
                filter_col = row[4].strip() if len(row) > 4 and row[4].strip() else None
                filter_val = row[5].strip() if len(row) > 5 and row[5].strip() else None
                rel_type = row[6].strip() if len(row) > 6 and row[6].strip() else "외래키"
                
                # 타입코드 기반 관계인 경우
                if filter_col and filter_val:
                    mappings.append({
                        "source_table": src_tbl,
                        "source_column": src_col,
                        "target_table": tgt_tbl,
                        "target_column": tgt_col,
                        "filter_column": filter_col,
                        "filter_value": filter_val,
                        "relation_type": rel_type
                    })
                # 일반 외래키 관계인 경우
                else:
                    mappings.append({
                        "source_table": src_tbl,
                        "source_column": src_col,
                        "target_table": tgt_tbl,
                        "target_column": tgt_col,
                        "relation_type": rel_type
                    })
        
        return mappings
    except Exception as e:
        print(f"[TSV 로드 오류] {e}")
        return []

class AddRelationshipDialog(tk.Toplevel):
    def __init__(self, parent, db_folder, tables=None, on_add=None):
        super().__init__(parent)
        self.title("관계 추가")
        self.geometry("600x400")  # 크기 확대
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        
        self.db_folder = db_folder  # DB 폴더 경로 저장
        self.tables = tables or []
        self.on_add = on_add
        
        self.table_var = tk.StringVar()
        self.column_var = tk.StringVar()
        self.related_var = tk.StringVar()
        self.bulk_var = tk.StringVar()  # 일괄 입력용
        
        self.build_ui()
    
    def build_ui(self):
        # 탭 컨트롤 생성
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 단일 입력 탭
        single_tab = ttk.Frame(self.notebook)
        self.notebook.add(single_tab, text="단일 입력")
        
        # 일괄 입력 탭
        bulk_tab = ttk.Frame(self.notebook)
        self.notebook.add(bulk_tab, text="일괄 입력")
        
        # 단일 입력 UI 구성
        self.build_single_ui(single_tab)
        
        # 일괄 입력 UI 구성
        self.build_bulk_ui(bulk_tab)
    
    def build_single_ui(self, parent):
        # 테이블 선택 영역
        ttk.Label(parent, text="테이블:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.table_combobox = ttk.Combobox(parent, textvariable=self.table_var, values=self.tables, width=30)
        self.table_combobox.grid(row=0, column=1, sticky="we", padx=5, pady=5)
        self.table_combobox.bind("<<ComboboxSelected>>", self.on_table_selected)
        
        # 컬럼 선택 영역
        ttk.Label(parent, text="컬럼:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.column_combobox = ttk.Combobox(parent, textvariable=self.column_var, width=30)
        self.column_combobox.grid(row=1, column=1, sticky="we", padx=5, pady=5)
        
        # 관계 컬럼 입력 영역
        ttk.Label(parent, text="관계 정보:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        
        # 설명 레이블
        ttk.Label(parent, text="형식: TargetColumn:TargetTable:TargetTableColumn, ...").grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        # 입력 필드
        self.related_entry = ttk.Entry(parent, textvariable=self.related_var, width=50)
        self.related_entry.grid(row=3, column=0, columnspan=2, sticky="we", padx=5, pady=5)
        
        # 예시 텍스트
        example_text = "예시: HeroID:Hero:ID, ItemID:Item:ID\n(각 관계는 쉼표로 구분, 형식은 '컬럼:테이블:테이블컬럼')"
        ttk.Label(parent, text=example_text, foreground="gray").grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        
        # 버튼 영역
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="추가", command=self.add_relationship).pack(side="left", padx=5)
        ttk.Button(button_frame, text="취소", command=self.destroy).pack(side="left", padx=5)
    
    def build_bulk_ui(self, parent):
        # 설명 레이블
        ttk.Label(parent, text="여러 관계를 한번에 추가합니다. 각 줄마다 한 관계를 입력하세요.").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(parent, text="형식: 테이블명,컬럼명,관계정보").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        
        # 예시 텍스트
        example_text = "예시:\nHero,SkillID,SkillID:Skill:ID\nInventory,ItemID,ItemID:Item:ID"
        ttk.Label(parent, text=example_text, foreground="gray").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        
        # 텍스트 영역
        self.bulk_text = tk.Text(parent, width=50, height=10)
        self.bulk_text.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(parent, command=self.bulk_text.yview)
        scrollbar.grid(row=3, column=1, sticky="ns")
        self.bulk_text.config(yscrollcommand=scrollbar.set)
        
        # 버튼 영역
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="일괄 추가", command=self.add_bulk_relationships).pack(side="left", padx=5)
        ttk.Button(button_frame, text="취소", command=self.destroy).pack(side="left", padx=5)
        
        # 그리드 설정
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)
    
    def on_table_selected(self, event=None):
        table = self.table_var.get()
        if not table:
            return
        
        # 선택된 테이블의 컬럼 목록 가져오기
        columns = self.get_table_columns(table)
        self.column_combobox['values'] = columns
        
        if columns:
            self.column_var.set(columns[0])
    
    def get_table_columns(self, table):
        # 테이블의 컬럼 목록을 가져오는 함수
        try:
            db_path = os.path.join(self.db_folder, f"{table}.db")
            
            if not os.path.exists(db_path):
                print(f"[오류] DB 파일이 존재하지 않음: {db_path}")
                return []
                
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA table_info({table})")
            columns = [row[1] for row in cursor.fetchall()]
            conn.close()
            
            print(f"[정보] 테이블 '{table}'의 컬럼: {columns}")
            return columns
        except Exception as e:
            print(f"[컬럼 정보 오류] {table}: {e}")
            return []
    
    def add_relationship(self):
        table = self.table_var.get()
        column = self.column_var.get()
        related = self.related_var.get()
        
        if not table or not column or not related:
            messagebox.showwarning("입력 오류", "모든 필드를 입력해주세요.")
            return
        
        # 입력된 관계 형식 검증
        valid = True
        for rel in related.split(','):
            parts = rel.strip().split(':')
            if len(parts) < 2:
                valid = False
                break
        
        if not valid:
            messagebox.showwarning("형식 오류", "관계 정보는 'TargetColumn:TargetTable[:TargetTableColumn]' 형식이어야 합니다.\n여러 관계는 쉼표로 구분해주세요.")
            return
        
        # 결과 전달 및 창 닫기
        if self.on_add:
            self.on_add({
                "table": table,
                "column": column,
                "related_columns": related
            })
        
        self.destroy()
    
    def add_bulk_relationships(self):
        bulk_text = self.bulk_text.get(1.0, tk.END).strip()
        
        if not bulk_text:
            messagebox.showwarning("입력 오류", "일괄 추가할 내용을 입력해주세요.")
            return
        
        # 각 줄 파싱
        relationships = []
        invalid_lines = []
        
        for line_num, line in enumerate(bulk_text.split('\n'), 1):
            line = line.strip()
            if not line:
                continue
                
            # CSV 형태로 파싱
            import csv
            from io import StringIO
            
            reader = csv.reader(StringIO(line))
            parts = next(reader, None)
            
            if not parts or len(parts) < 3:
                invalid_lines.append(f"{line_num}: {line}")
                continue
            
            table, column, related = parts[:3]
            
            # 관계 형식 검증
            valid = True
            for rel in related.split(','):
                rel_parts = rel.strip().split(':')
                if len(rel_parts) < 2:
                    valid = False
                    break
            
            if not valid:
                invalid_lines.append(f"{line_num}: {line}")
                continue
            
            relationships.append({
                "table": table.strip(),
                "column": column.strip(),
                "related_columns": related.strip()
            })
        
        # 오류 있는 경우 보고
        if invalid_lines:
            messagebox.showwarning("형식 오류", 
                f"다음 {len(invalid_lines)}개 줄의 형식이 잘못되었습니다:\n\n" + 
                "\n".join(invalid_lines[:5]) + 
                ("\n..." if len(invalid_lines) > 5 else ""))
            return
        
        # 결과 전달 및 창 닫기
        if self.on_add and relationships:
            for rel in relationships:
                self.on_add(rel)
        
        messagebox.showinfo("추가 완료", f"{len(relationships)}개의 관계가 추가되었습니다.")
        self.destroy()
       
class TableRelationshipAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("테이블 관계 분석기")
        self.root.geometry("800x600")
        
        self.db_folder_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.excel_path = tk.StringVar()  # 엑셀 파일 경로
        self.status_var = tk.StringVar(value="대기 중...")
        
        # 수동 관계 목록
        self.manual_relationships = []
        # 파일 타임스탬프 저장
        self.excel_last_modified = -1
        
        # 주기적 체크 타이머 (30초마다)
        self.check_interval = 30000  # 30초
        
        self.build_ui()
    
    def build_ui(self):
        # 경로 입력 프레임
        path_frame = ttk.LabelFrame(self.root, text="경로 설정")
        path_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(path_frame, text="DB 폴더 경로:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.db_folder_path, width=60).grid(row=0, column=1, sticky="we", padx=5, pady=5)
        ttk.Button(path_frame, text="찾기", command=self.select_db_folder).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(path_frame, text="결과 저장 폴더:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.output_folder_path, width=60).grid(row=1, column=1, sticky="we", padx=5, pady=5)
        ttk.Button(path_frame, text="찾기", command=self.select_output_folder).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Label(path_frame, text="테이블 정리 엑셀:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(path_frame, textvariable=self.excel_path, width=55).grid(row=2, column=1, sticky="we", padx=5, pady=5)
        
        excel_button_frame = ttk.Frame(path_frame)
        excel_button_frame.grid(row=2, column=2, padx=5, pady=5)
        
        ttk.Button(excel_button_frame, text="찾기", command=self.select_excel_file).pack(side="left", padx=2)
        ttk.Button(excel_button_frame, text="새로고침", command=self.refresh_excel_data).pack(side="left", padx=2)
        
        
        # 관계 관리 프레임
        relation_frame = ttk.LabelFrame(self.root, text="수동 관계 관리")
        relation_frame.pack(fill="x", padx=10, pady=5)
        
        # 관계 목록 트리뷰
        self.relation_tree = ttk.Treeview(relation_frame, columns=("table", "column", "related", "type"), show="headings", height=5)
        self.relation_tree.heading("table", text="테이블")
        self.relation_tree.heading("column", text="컬럼")
        self.relation_tree.heading("related", text="관계 정보")
        
        self.relation_tree.column("table", width=150)
        self.relation_tree.column("column", width=150)
        self.relation_tree.column("related", width=300)
        
        self.relation_tree.pack(fill="x", padx=5, pady=5)
        
        # 관계 관리 버튼
        rel_button_frame = ttk.Frame(relation_frame)
        rel_button_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(rel_button_frame, text="관계 추가", command=self.add_relationship).pack(side="left", padx=5)
        ttk.Button(rel_button_frame, text="관계 삭제", command=self.delete_relationship).pack(side="left", padx=5)
        ttk.Button(rel_button_frame, text="관계 저장", command=self.save_relationships).pack(side="left", padx=5)
        ttk.Button(rel_button_frame, text="관계 불러오기", command=self.load_relationships).pack(side="left", padx=5)
        ttk.Button(rel_button_frame, text="엑셀에서 가져오기", command=self.import_relationships_from_excel).pack(side="left", padx=5)
        ttk.Button(rel_button_frame, text="TSV에서 가져오기", command=self.import_relationships_from_tsv).pack(side="left", padx=5)

        
        # 버튼 프레임
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(button_frame, text="분석 시작", command=self.start_analysis).pack(side="left", padx=5)
        ttk.Button(button_frame, text="결과 보기", command=self.view_results).pack(side="left", padx=5)
        ttk.Button(button_frame, text="엑셀 정보 보기", command=self.view_excel_info).pack(side="left", padx=5)
        
        # 결과 표시 영역
        result_frame = ttk.LabelFrame(self.root, text="분석 결과")
        result_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.result_text = tk.Text(result_frame, wrap="word", width=80, height=20)
        self.result_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 상태 표시줄
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom", padx=5, pady=5)
    
    def select_db_folder(self):
        folder = filedialog.askdirectory(title="DB 폴더 선택")
        if folder:
            self.db_folder_path.set(folder)
            
            # 기본 출력 폴더 설정 (.cache/relationships)
            default_output = os.path.join(os.path.dirname(folder), ".cache", "relationships")
            self.output_folder_path.set(default_output)
    
    def select_output_folder(self):
        folder = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if folder:
            self.output_folder_path.set(folder)

    def select_excel_file(self):
        excel_file = filedialog.askopenfilename(
            title="테이블 정리 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("Excel 파일(이전 버전)", "*.xls")]
        )
        if excel_file:
            self.excel_path.set(excel_file)
            # 선택 즉시 로드하고 관계 정보 추가
            self.refresh_excel_data()
            self.excel_last_modified = get_file_mtime(excel_file)

    def view_excel_info(self):
        excel_path = self.excel_path.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showwarning("파일 오류", "엑셀 파일을 선택해주세요.")
            return

        try:
            excel_data = load_excel_data(excel_path)
            if not excel_data:
                messagebox.showwarning("로드 오류", "엑셀 파일을 로드하는 중 오류가 발생했습니다.")
                return
            
            # 결과 표시
            self.result_text.delete(1.0, tk.END)
            
            self.result_text.insert(tk.END, "=== 테이블 정보 ===\n", "header")
            for table in excel_data.get("tables", []):
                table_name = table.get("table_name", "")
                self.result_text.insert(tk.END, f"\n[{table_name}]\n", "table_name")
                
                for column in table.get("columns", []):
                    col_name = column.get("column_name", "")
                    related_table = column.get("related_table", "")
                    related_column = column.get("related_column", "")
                    
                    self.result_text.insert(tk.END, f"  {col_name}")
                    if related_table and related_column:
                        self.result_text.insert(tk.END, f" → {related_table}.{related_column}\n", "relation")
                    else:
                        self.result_text.insert(tk.END, "\n")
            
            if "enums" in excel_data and excel_data["enums"]:
                self.result_text.insert(tk.END, "\n\n=== Enum 정보 ===\n", "header")
                
                for enum in excel_data["enums"]:
                    enum_type = enum.get("type", "")
                    self.result_text.insert(tk.END, f"\n[{enum_type}]\n", "table_name")
                    
                    for value in enum.get("values", []):
                        val = value.get("value", "")
                        name = value.get("name", "")
                        self.result_text.insert(tk.END, f"  {val}: {name}\n")
            
            # 텍스트 태그 설정
            self.result_text.tag_configure("header", font=("Arial", 12, "bold"))
            self.result_text.tag_configure("table_name", font=("Arial", 10, "bold"))
            self.result_text.tag_configure("relation", foreground="blue")
            
            # 맨 위로 스크롤
            self.result_text.see("1.0")
            
            self.status_var.set(f"엑셀 정보 로드 완료: {len(excel_data.get('tables', []))} 테이블, {len(excel_data.get('enums', []))} Enum 타입")
            
        except Exception as e:
            messagebox.showerror("분석 오류", f"엑셀 파일 분석 중 오류 발생: {str(e)}")

    def import_relationships_from_tsv(self):
        """TSV 파일에서 관계 정보를 가져옵니다."""
        tsv_file = filedialog.askopenfilename(
            title="관계 정보 TSV/TXT 파일 선택",
            filetypes=[("TSV 파일", "*.tsv"), ("텍스트 파일", "*.txt")]
        )
        
        if not tsv_file:
            return
            
        mappings = load_typecode_mappings_from_tsv(tsv_file)
        if not mappings:
            messagebox.showwarning("로드 오류", "TSV 파일에서 관계 정보를 불러오지 못했습니다.")
            return
        
        # 타입코드 매핑 파일 저장
        output_folder = self.output_folder_path.get()
        if not output_folder:
            output_folder = os.path.join(os.path.dirname(self.db_folder_path.get()), ".cache", "relationships")
            self.output_folder_path.set(output_folder)
        
        os.makedirs(output_folder, exist_ok=True)
        typecode_path = os.path.join(output_folder, "typecode_mapping.json")
        
        try:
            with open(typecode_path, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("가져오기 완료", f"{len(mappings)}개의 관계를 가져와 저장했습니다.\n파일 경로: {typecode_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"관계 정보 저장 중 오류 발생: {str(e)}")

    def add_relationship(self):
        db_folder = self.db_folder_path.get()
        if not db_folder or not os.path.exists(db_folder):
            messagebox.showwarning("경로 오류", "유효한 DB 폴더를 선택해주세요.")
            return
        
        # DB 폴더에서 테이블 목록 가져오기 (DB 파일명 기준)
        tables = [os.path.splitext(f)[0] for f in os.listdir(db_folder) if f.endswith('.db')]
        
        if not tables:
            messagebox.showwarning("테이블 없음", "선택한 DB 폴더에서 테이블을 찾을 수 없습니다.")
            return
        
        # 관계 추가 다이얼로그 표시
        def on_add(relationship):
            self.manual_relationships.append(relationship)
            self.update_relationship_tree()
        
        dialog = AddRelationshipDialog(self.root, db_folder, tables=tables, on_add=on_add)
        self.root.wait_window(dialog)
        
    def delete_relationship(self):
        selected = self.relation_tree.selection()
        if not selected:
            messagebox.showwarning("선택 오류", "삭제할 관계를 선택해주세요.")
            return
        
        for item in selected:
            index = int(item)
            if 0 <= index < len(self.manual_relationships):
                del self.manual_relationships[index]
        
        self.update_relationship_tree()
    
    def update_relationship_tree(self):
        # 트리뷰 초기화
        for item in self.relation_tree.get_children():
            self.relation_tree.delete(item)
        
        # 컬럼 설정 업데이트 (관계 유형 추가)
        if len(self.relation_tree["columns"]) < 4:
            self.relation_tree["columns"] = ("table", "column", "related", "type")
            self.relation_tree.heading("type", text="관계 유형")
            self.relation_tree.column("type", width=100)
        
        # 관계 목록 표시
        for idx, rel in enumerate(self.manual_relationships):
            source = rel.get("source", "수동")
            table = rel["table"]
            column = rel["column"]
            related = rel["related_columns"]
            rel_type = rel.get("relation_type", "외래키")
            
            # 출처에 따라 다른 태그 사용
            tag = "excel" if source == "excel" else "manual"
            
            self.relation_tree.insert("", "end", iid=str(idx), values=(
                table,
                column,
                related,
                rel_type
            ), tags=(tag,))
        
        # 태그 설정 (엑셀 출처는 파란색, 수동은 기본색)
        self.relation_tree.tag_configure("excel", foreground="blue")
        
    def save_relationships(self):
        if not self.manual_relationships:
            messagebox.showwarning("저장 오류", "저장할 수동 관계가 없습니다.")
            return
        
        output_folder = self.output_folder_path.get()
        if not output_folder:
            output_folder = os.path.join(os.path.dirname(self.db_folder_path.get()), ".cache", "relationships")
            self.output_folder_path.set(output_folder)
        
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, "manual_relationships.json")
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(self.manual_relationships, f, indent=2, ensure_ascii=False)
            
            self.status_var.set(f"수동 관계 저장 완료: {output_path}")
            messagebox.showinfo("저장 완료", f"수동 관계 정보가 저장되었습니다.\n파일 경로: {output_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"수동 관계 정보 저장 중 오류 발생: {str(e)}")
    
    def load_relationships(self):
        output_folder = self.output_folder_path.get()
        if not output_folder:
            output_folder = os.path.join(os.path.dirname(self.db_folder_path.get()), ".cache", "relationships")
            self.output_folder_path.set(output_folder)
        
        os.makedirs(output_folder, exist_ok=True)
        input_path = os.path.join(output_folder, "manual_relationships.json")
        
        if not os.path.exists(input_path):
            messagebox.showwarning("파일 없음", f"수동 관계 정보 파일을 찾을 수 없습니다: {input_path}")
            return
        
        try:
            with open(input_path, 'r', encoding='utf-8') as f:
                self.manual_relationships = json.load(f)
            
            self.update_relationship_tree()
            self.status_var.set(f"수동 관계 로드 완료: {len(self.manual_relationships)}개")
            messagebox.showinfo("로드 완료", f"수동 관계 정보가 로드되었습니다.\n총 {len(self.manual_relationships)}개의 관계")
        except Exception as e:
            messagebox.showerror("로드 오류", f"수동 관계 정보 로드 중 오류 발생: {str(e)}")
    
    def start_analysis(self):
        db_folder = self.db_folder_path.get()
        if not db_folder:
            messagebox.showwarning("경로 오류", "DB 폴더 경로를 선택해주세요.")
            return
        
        if not os.path.exists(db_folder):
            messagebox.showwarning("경로 오류", f"선택한 DB 폴더가 존재하지 않습니다: {db_folder}")
            return
        
        output_folder = self.output_folder_path.get()
        if not output_folder:
            output_folder = os.path.join(os.path.dirname(db_folder), ".cache", "relationships")
            self.output_folder_path.set(output_folder)
        
        # 결과 파일 경로
        output_path = os.path.join(output_folder, "table_relationships.json")
        
        # 엑셀 데이터 로드
        excel_data = None
        excel_path = self.excel_path.get()
        if excel_path and os.path.exists(excel_path):
            excel_data = load_excel_data(excel_path)
            if excel_data:
                self.status_var.set(f"엑셀 데이터 로드 완료: {len(excel_data.get('tables', []))} 테이블")
        
        # 타입코드 매핑 로드
        typecode_mappings = None
        typecode_path = os.path.join(output_folder, "typecode_mapping.json")
        if os.path.exists(typecode_path):
            try:
                with open(typecode_path, 'r', encoding='utf-8') as f:
                    typecode_mappings = json.load(f)
                    self.status_var.set(f"{self.status_var.get()} | 타입코드 매핑 로드: {len(typecode_mappings)}개")
            except Exception as e:
                print(f"[타입코드 매핑 로드 오류] {e}")
        
        # 진행 상태 표시 창
        progress_window = tk.Toplevel(self.root)
        progress_window.title("분석 중...")
        progress_window.geometry("400x100")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        progress_label = ttk.Label(progress_window, text="준비 중...", anchor="center")
        progress_label.pack(pady=20, padx=10, fill="x")
        
        # 진행 상태 업데이트 함수
        def update_progress(message):
            self.root.after(0, lambda: progress_label.configure(text=message))
            self.root.update_idletasks()
        
        # 분석 작업 함수
        def do_analyze():
            try:
                relationships = analyze_table_relationships(
                    db_folder,
                    output_path=output_path,
                    progress_callback=update_progress,
                    manual_relationships=self.manual_relationships,
                    excel_data=excel_data,
                    typecode_mappings=typecode_mappings
                )
                
                # 성공시 UI 업데이트
                self.root.after(0, lambda: [
                    progress_window.destroy(),
                    self.status_var.set(f"분석 완료: {len(relationships)} 테이블, 결과 저장 경로: {output_path}"),
                    messagebox.showinfo("분석 완료", f"테이블 관계 분석 완료!\n총 {len(relationships)} 테이블 분석됨\n결과 저장 경로: {output_path}"),
                    self.display_results(relationships)
                ])
            except Exception as e:
                # 오류 발생시 UI 업데이트
                self.root.after(0, lambda: [
                    progress_window.destroy(),
                    self.status_var.set(f"분석 오류: {str(e)}"),
                    messagebox.showerror("분석 오류", f"테이블 관계 분석 중 오류 발생:\n{str(e)}")
                ])
        
        # 스레드 시작
        import threading
        thread = threading.Thread(target=do_analyze)
        thread.daemon = True
        thread.start()

    def view_results(self):
        output_folder = self.output_folder_path.get()
        if not output_folder or not os.path.exists(output_folder):
            messagebox.showwarning("경로 오류", "유효한 결과 폴더를 선택해주세요.")
            return
        
        result_path = os.path.join(output_folder, "table_relationships.json")
        if not os.path.exists(result_path):
            messagebox.showwarning("파일 없음", f"결과 파일을 찾을 수 없습니다: {result_path}")
            return
        
        try:
            with open(result_path, 'r', encoding='utf-8') as f:
                relationships = json.load(f)
            self.display_results(relationships)
            self.status_var.set(f"결과 로드 완료: {len(relationships)} 테이블")
        except Exception as e:
            messagebox.showerror("로드 오류", f"결과 파일을 로드하는 중 오류 발생: {str(e)}")
    
    def display_results(self, relationships):
        self.result_text.delete(1.0, tk.END)
        
        total_relations = 0
        
        # 결과 텍스트 형식으로 표시
        for source_table, targets in relationships.items():
            if not targets:  # 관계가 없는 테이블은 건너뛰기
                continue
                
            self.result_text.insert(tk.END, f"== {source_table} ==\n", "table_header")
            
            for target_table, relations in targets.items():
                for relation in relations:
                    total_relations += 1
                    source_col = relation["source_column"]
                    target_col = relation["target_column"]
                    rel_type = relation["relation_type"]
                    
                    if rel_type == "foreign_key":
                        rel_desc = "외래키 (확정)"
                    elif rel_type == "manual_foreign_key":
                        rel_desc = "외래키 (수동 추가)"
                    elif rel_type == "typecode_foreign_key":
                        filter_col = relation.get("filter_column", "")
                        filter_val = relation.get("filter_value", "")
                        rel_desc = f"조건 외래키 ({filter_col}={filter_val})"
                    else:
                        similarity = relation.get("similarity", "N/A")
                        rel_desc = f"유사 외래키 (유사도: {similarity})"
                    
                    self.result_text.insert(tk.END, f"  {source_col} → {target_table}.{target_col} [{rel_desc}]\n")
        
                self.result_text.insert(tk.END, "\n")
        
            self.result_text.insert(tk.END, f"총 {len(relationships)} 테이블, {total_relations} 관계 발견\n")
            
            # 텍스트 스타일 설정
            self.result_text.tag_configure("table_header", font=("Arial", 10, "bold"))
            
            # 맨 위로 스크롤
            self.result_text.see("1.0")

    def import_relationships_from_excel(self):
        """엑셀 파일에서 관계 정보를 가져옵니다."""
        excel_file = filedialog.askopenfilename(
            title="관계 정보 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("Excel 파일(이전 버전)", "*.xls")]
        )
        
        if not excel_file:
            return
            
        try:
            import pandas as pd
            df = pd.read_excel(excel_file)
            
            required_columns = ["테이블", "컬럼", "관계정보"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                messagebox.showwarning("형식 오류", 
                    f"필요한 열이 없습니다: {', '.join(missing_columns)}\n\n"
                    f"엑셀 파일에는 '테이블', '컬럼', '관계정보' 열이 필요합니다.")
                return
            
            count = 0
            for _, row in df.iterrows():
                if pd.isna(row["테이블"]) or pd.isna(row["컬럼"]) or pd.isna(row["관계정보"]):
                    continue
                    
                relationship = {
                    "table": str(row["테이블"]).strip(),
                    "column": str(row["컬럼"]).strip(),
                    "related_columns": str(row["관계정보"]).strip()
                }
                
                self.manual_relationships.append(relationship)
                count += 1
            
            self.update_relationship_tree()
            messagebox.showinfo("가져오기 완료", f"{count}개의 관계를 가져왔습니다.")
            
        except Exception as e:
            messagebox.showerror("가져오기 오류", f"엑셀 파일 가져오기 오류: {str(e)}")

    def check_excel_update(self):
        """엑셀 파일 변경 여부를 확인하고 자동으로 업데이트합니다."""
        excel_path = self.excel_path.get()
        if excel_path and os.path.exists(excel_path):
            current_mtime = get_file_mtime(excel_path)
            if current_mtime > self.excel_last_modified:
                # 파일이 변경된 경우
                if self.excel_last_modified != -1:  # 처음이 아닌 경우만
                    self.status_var.set("엑셀 파일 변경 감지됨. 데이터 갱신 중...")
                    self.refresh_excel_data()
                self.excel_last_modified = current_mtime
        
        # 다음 체크 예약
        self.root.after(self.check_interval, self.check_excel_update)

    def refresh_excel_data(self):
        """엑셀 데이터를 새로 불러와 관계 정보를 업데이트합니다."""
        excel_path = self.excel_path.get()
        if not excel_path or not os.path.exists(excel_path):
            return
            
        try:
            excel_data = load_excel_data(excel_path)
            if excel_data:
                # 기존 엑셀 기반 관계 제거
                self.manual_relationships = [rel for rel in self.manual_relationships 
                                            if rel.get("source", "") != "excel"]
                
                # 새 관계 추가
                typecode_mappings = []
                
                # 타입코드 매핑용 출력 폴더 설정
                output_folder = self.output_folder_path.get()
                if not output_folder:
                    output_folder = os.path.join(os.path.dirname(self.db_folder_path.get()), ".cache", "relationships")
                    self.output_folder_path.set(output_folder)
                
                os.makedirs(output_folder, exist_ok=True)
                
                # 관계 정보 처리
                for rel in excel_data.get("relationships", []):
                    # 타입코드 매핑 정보 구성
                    src_tbl = rel["source_table"]
                    src_col = rel["source_column"]
                    tgt_tbl = rel["target_table"]
                    tgt_col = rel["target_column"]
                    filter_col = rel.get("filter_column", "")
                    filter_val = rel.get("filter_value", "")
                    rel_type = rel.get("relation_type", "외래키")
                    
                    # 수동 관계 목록용
                    related_info = f"{src_col}:{tgt_tbl}:{tgt_col}"
                    manual_rel = {
                        "table": src_tbl,
                        "column": src_col,
                        "related_columns": related_info,
                        "source": "excel",
                        "relation_type": rel_type
                    }
                    
                    if manual_rel not in self.manual_relationships:
                        self.manual_relationships.append(manual_rel)
                    
                    # 타입코드 매핑용
                    mapping = {
                        "source_table": src_tbl,
                        "source_column": src_col,
                        "target_table": tgt_tbl,
                        "target_column": tgt_col,
                        "relation_type": rel_type
                    }
                    
                    # 필터 정보가 있는 경우만 추가
                    if filter_col and filter_val:
                        mapping["filter_column"] = filter_col
                        mapping["filter_value"] = filter_val
                    
                    typecode_mappings.append(mapping)
                
                # 타입코드 매핑 저장
                if typecode_mappings:
                    typecode_path = os.path.join(output_folder, "typecode_mapping.json")
                    try:
                        with open(typecode_path, 'w', encoding='utf-8') as f:
                            json.dump(typecode_mappings, f, indent=2, ensure_ascii=False)
                        print(f"타입코드 매핑 저장 완료: {typecode_path}")
                    except Exception as e:
                        print(f"[타입코드 매핑 저장 오류] {e}")
                
                # UI 업데이트
                self.update_relationship_tree()
                rel_count = len([r for r in self.manual_relationships if r.get("source") == "excel"])
                self.status_var.set(f"엑셀 데이터 자동 갱신 완료: {rel_count}개 관계")
                
        except Exception as e:
            self.status_var.set(f"엑셀 데이터 갱신 오류: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TableRelationshipAnalyzerApp(root)
    root.mainloop()