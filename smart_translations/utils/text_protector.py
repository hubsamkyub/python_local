"""
텍스트 보호 관련 유틸리티
"""
import re


class TextProtector:
    """번역 시 특수 태그와 변수를 보호하는 클래스"""
    
    def __init__(self):
        # 보호할 패턴들 정의
        self.protection_patterns = [
            # 색상 태그: [#ffffff]텍스트[#]
            (r'\[#[a-fA-F0-9]{6}\]([^[]*?)\[#\]', 'COLOR'),
            # 변수 플레이스홀더: {숫자}, {숫자%}, {변수명} 등
            (r'\{[^}]+\}', 'VAR'),
            # HTML/XML 스타일 태그: <태그>내용</태그>
            (r'<[^>]+>[^<]*</[^>]+>', 'HTML'),
            # 단순 마크업: [태그]내용[/태그]
            (r'\[[^\]]+\][^[]*?\[/[^\]]+\]', 'MARKUP'),
            # 숫자 + % 조합
            (r'\d+%', 'PERCENT'),
            # @ 시작 변수
            (r'@\w+', 'AT_VAR'),
        ]
        self.reset()
    
    def reset(self):
        """새 텍스트 처리를 위한 초기화"""
        self.protected_items = {}
        self.placeholder_counter = 0
    
    def protect_text(self, text):
        """텍스트에서 특수 태그들을 플레이스홀더로 치환"""
        if not text or not isinstance(text, str):
            return text, {}
            
        protected_text = text
        protection_map = {}
        
        # 각 패턴에 대해 보호 처리
        for pattern, tag_type in self.protection_patterns:
            matches = list(re.finditer(pattern, protected_text))
            
            # 뒤에서부터 처리 (인덱스 변경 방지)
            for match in reversed(matches):
                original = match.group(0)
                placeholder = f"<PROTECTED_{self.placeholder_counter}>"
                
                protection_map[placeholder] = {
                    'original': original,
                    'type': tag_type,
                    'position': match.span()
                }
                
                # 텍스트에서 원본을 플레이스홀더로 교체
                protected_text = protected_text[:match.start()] + placeholder + protected_text[match.end():]
                self.placeholder_counter += 1
        
        return protected_text, protection_map
    
    def restore_text(self, translated_text, protection_map):
        """번역된 텍스트에서 플레이스홀더를 원본으로 복원"""
        if not translated_text or not protection_map:
            return translated_text
            
        restored_text = translated_text
        
        # 모든 플레이스홀더를 원본으로 복원
        for placeholder, info in protection_map.items():
            if placeholder in restored_text:
                restored_text = restored_text.replace(placeholder, info['original'])
        
        return restored_text
    
    def extract_translatable_parts(self, text):
        """색상 태그 내부의 번역 가능한 부분만 추출"""
        if not text:
            return []
            
        translatable_parts = []
        
        # 색상 태그 패턴: [#ffffff]텍스트[#]
        color_pattern = r'\[#[a-fA-F0-9]{6}\]([^[]+?)\[#\]'
        matches = re.finditer(color_pattern, text)
        
        for match in matches:
            inner_text = match.group(1).strip()
            if inner_text and not re.match(r'^[|]+$', inner_text):  # |만 있는 것은 제외
                translatable_parts.append({
                    'text': inner_text,
                    'start': match.start(1),
                    'end': match.end(1),
                    'full_match': match.group(0)
                })
        
        return translatable_parts
    
