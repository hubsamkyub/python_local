import ast
import os

# --- 설정 부분 ---

# 원본 파일 (모든 코드가 다 들어있는 파일)
SOURCE_FILE = "smart_translation_manager.py" 
# 대상 파일 (제가 드린 '뼈대' 파일. 이 파일에 코드를 추가합니다.)
DEST_FILE = "smart_translation_manager_skeleton.py" # 임시 뼈대 파일명
# 최종 결과가 저장될 파일
OUTPUT_FILE = "smart_translation_manager_filled.py"

# setup_ 함수를 제외하고, 복사해야 할 함수 이름 목록
# 이전 답변에서 드린 목록을 그대로 사용합니다.
FUNCTIONS_TO_COPY = [
    # UI 이벤트 처리 함수
    'on_engine_changed', 'on_scenario_option_changed', 'handle_drop', 'on_tree_click',
    'on_conflict_row_selected', 'on_inline_edit_complete', 'on_speaker_saved',

    # 버튼 동작 및 기능 실행 함수
    'select_file', 'load_data', 'analyze_translations', 'execute_translation',
    '_execute_translation_thread', 'save_results', 'force_retranslate_selected',
    '_force_retranslate_thread', 'edit_translation_inline', 'remove_from_tm',
    'view_tm_entry', 'sync_glossary_from_gsheet', '_import_gsheet_thread',
    'start_db_build', 'search_excel_for_import', 'start_excel_import',
    '_excel_import_thread', 'update_tm_from_excel', '_update_tm_thread',
    'show_update_preview', 'apply_tm_updates', 'import_tm_from_folder',
    
    # UI 업데이트 및 헬퍼 함수
    'update_translation_table', 'filter_translations', 'update_stats_label',
    'determine_status', 'find_item_id_by_string_id', 'toggle_all_selections',
    'toggle_selected_checkboxes', 'toggle_single_item', 'toggle_multiple_items',
    'select_all_items', 'deselect_all_items', 'update_select_all_checkbox',
    'invert_selection', 'clear_selected_translations', 'edit_selected_item',
    'update_status',

    # 컨텍스트 메뉴 및 프롬프트 관련 함수
    'show_context_menu', 'set_prompt_template', 'set_scenario_mode_prompt', 'get_llm_prompt',

    # 시나리오 및 레퍼런스 관련 함수
    'ensure_scenario_manager', 'refresh_speaker_list', 'update_speaker_list', 'add_speaker',
    'edit_speaker', 'delete_speaker', 'select_reference_file', 'load_reference_from_gsheet',
    '_load_reference_gsheet_thread', 'prepare_scenario_translation', 'handle_unknown_speakers',
    'auto_create_speaker_profiles', 'infer_gender_from_name', 'manual_add_speakers',
    'analyze_reference_data_smart', 'debug_file_structure_detailed', 'show_latest_debug_log',
    'auto_save_reference_analysis', 'show_reference_dataset_manager',
    'translate_with_enhanced_scenario', 'get_speaker_for_item_enhanced',
    'debug_file_structure', 'show_debug_result', 'copy_to_clipboard', 
    'manual_column_mapping_dialog', 'analyze_with_skiprows', 'find_language_columns', 
    'show_manual_column_mapping',

    # 초기 설정 및 파일 처리
    'check_and_create_config_files', 'create_config_templates', 'show_setup_instructions',
    
    # 기타 유틸리티 함수
    'find_similar_translation', 'apply_glossary', 'check_multilang_prerequisites', 'get_speaker_for_item',
    'get_speaker_for_translation', 'translate_with_protection', 'translate_complex_markup',
    'show_efficiency_report', 'create_summary_tab', 'create_api_usage_tab', 
    'create_tm_efficiency_tab', 'create_cost_analysis_tab', 'show_translation_report',
    
    # 이미 setup_... 함수 목록도 여기에 포함시키는 것이 안전합니다.
    'setup_ui', 'setup_scenario_tab', 'setup_keyboard_shortcuts', 'setup_compact_tabs',
    'setup_progress_bar', 'setup_status_help', 'setup_responsive_layout', 'setup_theme_support',
    'setup_memory_monitor', 'setup_enhanced_ui', 'setup_translation_tab', 'setup_glossary_tab',
    'setup_history_tab', 'setup_conflict_tab', 'setup_exclusion_tab', 'setup_tm_management_tab',
    'setup_tm_view_edit_tab', 'setup_excel_import_tab',
    
    # DB 관련 함수는 이미 DatabaseManager로 옮겨졌으므로 여기 목록에는 없습니다.
]


def get_functions_from_source(source_code):
    """
    소스 코드 문자열에서 모든 함수 정의(def)를 찾아
    {함수명: 전체 소스코드} 형태의 딕셔너리로 반환합니다.
    """
    try:
        tree = ast.parse(source_code)
        functions = {}
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef):
                # ast.get_source_segment를 사용하여 함수의 전체 소스코드를 정확히 추출합니다.
                function_source = ast.get_source_segment(source_code, node)
                if function_source:
                    functions[node.name] = function_source
        return functions
    except SyntaxError as e:
        print(f"Error parsing source code: {e}")
        return {}


def main():
    # 1. 원본 파일과 대상 파일 읽기
    try:
        with open(SOURCE_FILE, 'r', encoding='utf-8') as f:
            source_code = f.read()
        
        # 대상 파일이 없을 경우, 비어있는 문자열로 시작
        if os.path.exists(DEST_FILE):
            with open(DEST_FILE, 'r', encoding='utf-8') as f:
                dest_code = f.read()
        else:
            print(f"경고: 대상 파일 '{DEST_FILE}'을 찾을 수 없습니다. 빈 파일에 코드를 추가합니다.")
            dest_code = ""

    except FileNotFoundError as e:
        print(f"오류: 파일을 찾을 수 없습니다. {e}")
        return

    # 2. 각 파일에서 함수 목록 추출
    source_functions = get_functions_from_source(source_code)
    dest_functions = get_functions_from_source(dest_code)
    dest_function_names = set(dest_functions.keys())
    
    print(f"'{SOURCE_FILE}'에서 {len(source_functions)}개의 함수를 찾았습니다.")
    print(f"'{DEST_FILE}'에서 {len(dest_functions)}개의 함수를 찾았습니다.\n")

    # 3. 옮겨야 할 함수들을 식별하고 리스트에 추가
    functions_to_add = []
    for func_name in FUNCTIONS_TO_COPY:
        if func_name in source_functions and func_name not in dest_function_names:
            print(f"✅ 추가: '{func_name}' 함수를 복사 목록에 추가합니다.")
            functions_to_add.append(source_functions[func_name])
        elif func_name in dest_function_names:
            print(f"☑️ 건너뛰기: '{func_name}' 함수는 이미 대상 파일에 존재합니다.")
        else:
            print(f"⚠️ 경고: '{func_name}' 함수를 원본 파일에서 찾을 수 없습니다.")

    # 4. 최종 코드를 생성하여 새 파일에 저장
    if functions_to_add:
        # 각 함수 정의 사이에 두 줄씩 띄워서 가독성을 높입니다.
        code_to_append = "\n\n".join(functions_to_add)
        
        # 대상 코드의 마지막에 코드를 추가
        final_code = dest_code + "\n\n" + code_to_append
        
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write(final_code)
        
        print(f"\n🎉 작업 완료! 총 {len(functions_to_add)}개의 함수가 추가되었습니다.")
        print(f"결과가 '{OUTPUT_FILE}' 파일에 저장되었습니다.")
    else:
        print("\n✨ 추가할 함수가 없습니다. 이미 모든 함수가 대상 파일에 있는 것 같습니다.")

if __name__ == "__main__":
    # 스크립트 실행 전, 뼈대 파일(DEST_FILE)이 없으면 하나 만들어줍니다.
    if not os.path.exists(DEST_FILE):
        with open(DEST_FILE, 'w', encoding='utf-8') as f:
            f.write(f'# {DEST_FILE} - 이 파일은 비어있어도 됩니다.\n\n')
            
    main()