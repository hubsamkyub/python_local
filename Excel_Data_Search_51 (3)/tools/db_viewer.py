import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
# import subprocess  # 엑셀 열기 기능 제거로 불필요
import sys

def open_db_viewer_with_excel_match(master, db_folder_path, table_name, excel_cache, folder_path):
    # 메인 뷰어 창 생성
    viewer = tk.Toplevel(master)
    viewer.title(f"🔍 DB 보기 - {table_name}")
    viewer.geometry("1200x700")

    # 상태 표시줄
    status_label = tk.Label(viewer, text="🔄 DB 데이터를 불러오는 중...", anchor="w")
    status_label.pack(fill="x", padx=10, pady=5)

    # DB 연결 및 데이터 로드
    db_path = os.path.join(db_folder_path, f"{table_name}.db")
    if not os.path.exists(db_path):
        messagebox.showerror("DB 없음", f"{table_name}.db 파일을 찾을 수 없습니다.")
        return

    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA table_info({table_name})")
        column_names = [row[1] for row in cursor.fetchall()]

        query = f"SELECT * FROM {table_name}"
        df = pd.read_sql_query(query, conn)
        conn.close()
    except Exception as e:
        messagebox.showerror("DB 오류", f"{e}")
        return

    # 필터 컬럼 선택 프레임
    filter_selection_frame = tk.Frame(viewer)
    filter_selection_frame.pack(fill="x", padx=10, pady=(5, 0))
    
    tk.Label(filter_selection_frame, text="🔍 필터 컬럼 선택 (최대 3개):").pack(side="left", padx=(0, 10))

    # 기본 필터 컬럼 설정 - 최대 3개로 제한
    default_filter_columns = ["컬럼 1", "컬럼 2", "컬럼 3"][:min(3, len(column_names))]
    if len(default_filter_columns) < min(3, len(column_names)):
        default_filter_columns.extend(column_names[:min(3, len(column_names)) - len(default_filter_columns)])
    
    # 필터 드롭다운 생성
    dropdown_vars = []
    max_dropdowns = min(3, len(column_names))
    for i in range(max_dropdowns):
        dropdown_var = tk.StringVar()
        if i < len(default_filter_columns):
            dropdown_var.set(default_filter_columns[i])
        dropdown = ttk.Combobox(filter_selection_frame, textvariable=dropdown_var, 
                              values=column_names, width=15, state="readonly")
        dropdown.pack(side="left", padx=5)
        dropdown_vars.append(dropdown_var)

    # 필터 UI 컨테이너
    filter_container = tk.Frame(viewer)
    filter_container.pack(fill="x", padx=10, pady=5)
    
    # 버튼과 필터 입력 필드를 한 줄에 배치
    filter_row = tk.Frame(filter_container)
    filter_row.pack(fill="x", expand=True)
    
    # 버튼 프레임
    btn_frame = tk.Frame(filter_row)
    btn_frame.pack(side="left", padx=(0, 10))
    
    # 필터 필드 프레임
    filter_fields_frame = tk.Frame(filter_row)
    filter_fields_frame.pack(side="left", fill="x", expand=True)
    
    # 필터 엔트리 딕셔너리
    filter_entries = {}

    # 셀 상태 표시 레이블
    status_cell_label = tk.Label(viewer, text="", anchor="w", fg="blue")
    status_cell_label.pack(fill="x", padx=10, pady=(5, 0))
    
    # 결과 테이블 프레임
    results_frame = tk.Frame(viewer)
    results_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
    # 스크롤바
    tree_scroll_y = tk.Scrollbar(results_frame)
    tree_scroll_y.pack(side="right", fill="y")
    
    tree_scroll_x = tk.Scrollbar(results_frame, orient="horizontal")
    tree_scroll_x.pack(side="bottom", fill="x")
    
    # 트리뷰 생성
    tree = ttk.Treeview(results_frame, columns=column_names, show="headings",
                      yscrollcommand=tree_scroll_y.set,
                      xscrollcommand=tree_scroll_x.set)
    
    # 데이터 열 설정
    for col in column_names:
        tree.heading(col, text=col)
        tree.column(col, anchor="w", width=120, stretch=True)  # stretch=True로 변경하여 열 크기 조절 문제 해결
    
    tree.pack(fill="both", expand=True)
    
    # 스크롤바 연결
    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)
    
    # 트리뷰 스타일 설정
    style = ttk.Style()
    style.configure("Treeview", rowheight=25)
    style.map("Treeview", 
             background=[("selected", "#e0e0ff")],
             foreground=[("selected", "black")])
    
    # 선택 관련 변수들
    selected_item = None
    selected_column = None
    selected_cell_value = None
    last_selected_cell_column = None
    
    # 트리뷰 태그 설정
    tree.tag_configure("evenrow", background="#f0f0f0")
    tree.tag_configure("oddrow", background="white")
    tree.tag_configure("rowselect", background="#ffcccb")
    tree.tag_configure("cellselect", background="#ccdfff")
    
    # 필터 UI 설정 함수
    def setup_filter_ui(selected_columns):
        # 기존 위젯 제거
        for widget in filter_fields_frame.winfo_children():
            widget.destroy()
        filter_entries.clear()
        
        # 새 필터 필드 생성
        for i, col in enumerate(selected_columns[:3]):
            entry_frame = tk.Frame(filter_fields_frame)
            entry_frame.pack(side="left", padx=5)
            
            tk.Label(entry_frame, text=f"{col}:").pack(side="left")
            entry = tk.Entry(entry_frame, width=15)
            entry.pack(side="left", padx=2)
            filter_entries[col] = entry
    
    # 필터 컬럼 업데이트 함수
    def update_filter_columns():
        # 현재 선택된 컬럼 가져오기
        selected_cols = [var.get() for var in dropdown_vars if var.get()]
        # 중복 제거 및 최대 3개로 제한
        selected_cols = list(dict.fromkeys(selected_cols))[:min(3, len(selected_cols))]
        # UI 업데이트
        setup_filter_ui(selected_cols)
    
    # 필터 적용 버튼
    tk.Button(btn_frame, text="🔎 필터 적용", 
             command=lambda: apply_filter()).pack(side="left", padx=2)
    
    # 필터 초기화 버튼
    tk.Button(btn_frame, text="🔄 필터 초기화", 
             command=lambda: reset_filters()).pack(side="left", padx=2)
    
    # 필터 컬럼 적용 버튼
    tk.Button(filter_selection_frame, text="✅ 필터 컬럼 적용", 
             command=lambda: update_filter_columns()).pack(side="left", padx=5)
    
    # 필터 초기화 함수
    def reset_filters():
        for entry in filter_entries.values():
            entry.delete(0, tk.END)
        apply_filter()
    
    # 필터 적용 함수
    def apply_filter():
        filtered = df.copy()
        for col, entry in filter_entries.items():
            val = entry.get().strip()
            if val and col in filtered.columns:
                filtered = filtered[filtered[col].astype(str).str.contains(val, case=False, na=False)]
        update_treeview(filtered)
        status_label.config(text=f"🔍 결과: {len(filtered)}건")
    
    # 트리뷰 업데이트 함수
    def update_treeview(data):
        tree.delete(*tree.get_children())
        for idx, (_, row) in enumerate(data.iterrows()):
            row_id = tree.insert("", "end", text=str(idx+1), values=list(row))
            # 행 스타일 적용
            tree.item(row_id, tags=("evenrow" if idx % 2 == 0 else "oddrow",))
    
    # 셀 선택 처리 함수
    def on_cell_click(event):
        nonlocal selected_item, selected_column, selected_cell_value, last_selected_cell_column
        
        # 모든 아이템 초기화
        for item in tree.get_children():
            idx = int(tree.index(item))
            tree.item(item, tags=("evenrow" if idx % 2 == 0 else "oddrow",))
        
        region = tree.identify("region", event.x, event.y)
        
        if region == "cell":
            selected_item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            column_idx = int(column[1:]) - 1
            
            if 0 <= column_idx < len(column_names):
                selected_column = column_names[column_idx]
                row_values = tree.item(selected_item, "values")
                selected_cell_value = row_values[column_idx]
                
                # 행 스타일 유지하면서 셀 하이라이트 효과
                idx = int(tree.index(selected_item))
                base_tag = "evenrow" if idx % 2 == 0 else "oddrow"
                tree.selection_set(selected_item)
                status_cell_label.config(text=f"선택됨: {selected_column} = {selected_cell_value}")
    
    # 셀 내용 복사 함수
    def copy_cell_content(event):
        if selected_cell_value is not None:
            viewer.clipboard_clear()
            viewer.clipboard_append(str(selected_cell_value))
            status_label.config(text=f"복사됨: {selected_cell_value}")
        elif selected_item is not None:
            # 행 전체 복사
            row_values = tree.item(selected_item, "values")
            viewer.clipboard_clear()
            viewer.clipboard_append("\t".join(str(v) for v in row_values))
            status_label.config(text="행 내용 복사됨")
    
    # 더블 클릭 처리 함수 - 엑셀 열기 기능 제거
    def on_double_click(event):
        try:
            item = tree.selection()[0]
            row_values = tree.item(item, "values")
            # 더블클릭 이벤트를 행 선택으로만 사용
            status_label.config(text=f"행 선택됨: {tree.index(item) + 1}")
        except Exception as e:
            print(f"[더블 클릭 오류] {e}")
            status_label.config(text="선택 실패")
    
    # 이벤트 바인딩
    tree.bind("<ButtonRelease-1>", on_cell_click)
    tree.bind("<Double-1>", on_double_click)
    viewer.bind("<Control-c>", copy_cell_content)
    
    # 초기 설정
    setup_filter_ui(default_filter_columns)
    update_treeview(df)
    status_label.config(text=f"✅ 총 {len(df)}건 로딩 완료")