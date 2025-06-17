# 파일명: glossary_handler.py
import csv
import os

def load_glossary(glossary_path):
    """
    CSV 파일에서 용어집을 로드하여 딕셔너리로 반환합니다.
    (기존 코드와 동일하며, 파일만 분리되었습니다.)
    """
    glossary = {}
    if glossary_path and os.path.exists(glossary_path):
        with open(glossary_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            # 헤더 스킵
            next(reader, None)
            for row in reader:
                if len(row) >= 2:
                    source_term, target_term = row[0].strip(), row[1].strip()
                    if source_term:
                        glossary[source_term] = target_term
    return glossary