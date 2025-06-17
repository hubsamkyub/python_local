# 파일명: text_processor.py
import re

def is_html(text):
    """
    주어진 텍스트에 HTML 태그가 포함되어 있는지 확인합니다.
    """
    return bool(re.search(r'<[^>]+>', text))

def wrap_with_span(text, class_name="glossary-term"):
    """
    용어집에 있는 단어를 식별하기 위한 span 태그로 감쌉니다.
    """
    return f'<span class="{class_name}">{text}</span>'

def unwrap_span(text, class_name="glossary-term"):
    """
    전처리를 위해 추가했던 span 태그를 제거합니다.
    """
    pattern = re.compile(f'<span class="{class_name}">(.*?)</span>', re.DOTALL)
    return pattern.sub(r'\1', text)

def split_text_with_html(text, max_len=4500):
    """
    HTML 태그를 보존하면서 텍스트를 분할합니다.
    """
    if not is_html(text):
        # HTML이 아니면 간단히 분할
        return [text[i:i+max_len] for i in range(0, len(text), max_len)]

    chunks = []
    current_chunk = ""
    # HTML 태그 또는 일반 텍스트로 분리
    parts = re.split(r'(<[^>]+>)', text)

    for part in parts:
        if not part:
            continue
        # 현재 청크에 추가했을 때 길이를 초과하는지 확인
        if len(current_chunk) + len(part) > max_len:
            if current_chunk:
                chunks.append(current_chunk)
            current_chunk = part
        else:
            current_chunk += part

    if current_chunk:
        chunks.append(current_chunk)

    return chunks