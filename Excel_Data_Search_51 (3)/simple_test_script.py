#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
간단한 외부 링크 테스트 스크립트
"""

import sys
from openpyxl import load_workbook
import re

def simple_test(file_path):
    """간단한 테스트"""
    print(f"=== 간단한 테스트: {file_path} ===")
    
    # 패턴들
    ref_patterns = [r'#REF!', r'OFFSET\(#REF!']
    external_patterns = [r'\[\d+\]!']
    
    try:
        workbook = load_workbook(file_path, data_only=False)
        print("✅ 파일 열기 성공")
        
        found_count = 0
        
        # 명명된 범위 검사
        if hasattr(workbook, 'defined_names') and workbook.defined_names:
            keys = list(workbook.defined_names.keys())
            print(f"명명된 범위 키들: {keys}")
            
            for name_key in keys:
                defined_name = workbook.defined_names[name_key]
                if hasattr(defined_name, 'value') and defined_name.value:
                    value = str(defined_name.value)
                    print(f"{name_key}: '{value}'")
                    
                    # #REF! 검사
                    for pattern in ref_patterns:
                        if re.search(pattern, value):
                            print(f"  🚨 #REF! 패턴 '{pattern}' 매칭!")
                            found_count += 1
                            break
                    else:
                        # 외부 참조 검사
                        for pattern in external_patterns:
                            if re.search(pattern, value):
                                print(f"  🔗 외부 참조 패턴 '{pattern}' 매칭!")
                                found_count += 1
                                break
        
        print(f"\n총 {found_count}개 문제 발견")
        workbook.close()
        
    except Exception as e:
        print(f"❌ 오류: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        simple_test(sys.argv[1])
    else:
        file_path = input("파일 경로를 입력하세요: ").strip().strip('"')
        simple_test(file_path)