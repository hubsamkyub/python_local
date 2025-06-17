"""
클래스 자동 분리 도구
"""
import re
import os

def extract_class_from_file(file_path, class_name, output_path):
    """파일에서 특정 클래스를 추출하여 새 파일로 저장"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 클래스 정의 시작점 찾기
    class_pattern = rf'class {class_name}.*?:'
    class_match = re.search(class_pattern, content, re.MULTILINE)
    
    if not class_match:
        print(f"클래스 {class_name}을 찾을 수 없습니다.")
        return
    
    start_pos = class_match.start()
    
    # 클래스 끝점 찾기 (다음 class나 함수 정의 또는 파일 끝)
    remaining_content = content[start_pos:]
    
    # 간단한 방법: 들여쓰기 레벨로 클래스 끝 판단
    lines = remaining_content.split('\n')
    class_lines = [lines[0]]  # 클래스 정의 라인
    
    for i, line in enumerate(lines[1:], 1):
        if line.strip() == '':
            class_lines.append(line)
            continue
        if line.startswith('class ') or line.startswith('def ') and not line.startswith('    '):
            break
        class_lines.append(line)
    
    class_content = '\n'.join(class_lines)
    
    # 필요한 import 추가
    imports = """import tkinter as tk
from tkinter import ttk, messagebox
import json
"""
    
    # 출력 파일에 저장
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(imports + '\n\n' + class_content)
    
    print(f"클래스 {class_name}이 {output_path}로 추출되었습니다.")

# 사용 예시
if __name__ == "__main__":
    # 필요한 클래스들 추출
    classes_to_extract = [
        ("TextProtector", "utils/text_protector.py"),
        ("TranslationMetrics", "utils/metrics.py"),
        ("InlineEditDialog", "dialogs/edit_dialogs.py"),
        ("ScrollableCheckList", "utils/widgets.py"),
    ]
    
    for class_name, output_path in classes_to_extract:
        extract_class_from_file("smart_translation_manager.py", class_name, output_path)