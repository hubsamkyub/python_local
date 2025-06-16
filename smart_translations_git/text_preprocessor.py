"""
텍스트 전처리 및 특수 태그 필터링 모듈
용도: 번역 텍스트에서 [@...], [#...#] 태그는 제외하고 [일반텍스트]는 포함하여 처리
"""
import re
from typing import List, Tuple, Dict

class TextPreprocessor:
    """특수 태그 필터링 및 텍스트 전처리 클래스"""
    
    def __init__(self):
        # 제외할 패턴들
        self.exclude_patterns = [
            r'\[@[^\]]*\]',      # [@변수명], [@michel] 등
            r'\[#[^#]*#\]'       # [#color#red#], [#sound#] 등
        ]
        
        # 포함할 패턴: [일반텍스트] (@ 또는 #로 시작하지 않는)
        self.include_pattern = r'\[[^@#\]]+\]'
        
        # 디버깅을 위한 통계
        self.stats = {
            'total_processed': 0,
            'excluded_count': 0,
            'included_count': 0
        }
    
    def extract_searchable_text(self, text: str) -> str:
        """
        용어집 검색에 사용할 텍스트만 추출
        
        Args:
            text: 원본 텍스트
            
        Returns:
            검색 가능한 텍스트 (특수 태그 제거됨)
        """
        if not text or not isinstance(text, str):
            return ""
        
        self.stats['total_processed'] += 1
        
        # 작업용 텍스트 복사
        temp_text = text
        
        # 1. 제외할 태그들을 공백으로 대체
        excluded_count = 0
        for pattern in self.exclude_patterns:
            matches = re.findall(pattern, temp_text)
            excluded_count += len(matches)
            temp_text = re.sub(pattern, ' ', temp_text)
        
        self.stats['excluded_count'] += excluded_count
        
        # 2. 포함할 대괄호 태그 처리: [무기] → 무기
        include_matches = re.findall(self.include_pattern, temp_text)
        included_count = len(include_matches)
        
        for match in include_matches:
            # 대괄호 제거하고 내용만 추출
            clean_content = match[1:-1]  # [무기] → 무기
            temp_text = temp_text.replace(match, f' {clean_content} ', 1)
        
        self.stats['included_count'] += included_count
        
        # 3. 남은 대괄호가 있다면 일반 태그로 간주하고 처리
        # [기타태그] → 기타태그 (안전한 처리)
        temp_text = re.sub(r'\[([^\]@#]*)\]', r' \1 ', temp_text)
        
        # 4. 연속된 공백을 하나로 정리하고 앞뒤 공백 제거
        temp_text = re.sub(r'\s+', ' ', temp_text).strip()
        
        return temp_text
    
    def get_excluded_tags(self, text: str) -> List[str]:
        """제외된 태그들 목록 반환 (디버깅용)"""
        if not text:
            return []
        
        excluded = []
        for pattern in self.exclude_patterns:
            matches = re.findall(pattern, text)
            excluded.extend(matches)
        
        return excluded
    
    def get_included_tags(self, text: str) -> List[str]:
        """포함된 태그들 목록 반환 (디버깅용)"""
        if not text:
            return []
        
        return re.findall(self.include_pattern, text)
    
    def get_debug_info(self, text: str) -> Dict:
        """디버깅용 상세 정보 반환"""
        if not text:
            return {
                'original': '',
                'searchable': '',
                'excluded_tags': [],
                'included_tags': [],
                'processing_stats': self.stats.copy()
            }
        
        excluded_tags = self.get_excluded_tags(text)
        included_tags = self.get_included_tags(text)
        searchable_text = self.extract_searchable_text(text)
        
        return {
            'original': text,
            'searchable': searchable_text,
            'excluded_tags': excluded_tags,
            'included_tags': included_tags,
            'excluded_count': len(excluded_tags),
            'included_count': len(included_tags),
            'length_reduction': len(text) - len(searchable_text),
            'processing_stats': self.stats.copy()
        }
    
    def reset_stats(self):
        """통계 초기화"""
        self.stats = {
            'total_processed': 0,
            'excluded_count': 0,
            'included_count': 0
        }
    
    def get_stats(self) -> Dict:
        """현재 처리 통계 반환"""
        return self.stats.copy()

# 편의 함수들
def extract_searchable_text(text: str) -> str:
    """빠른 사용을 위한 편의 함수"""
    processor = TextPreprocessor()
    return processor.extract_searchable_text(text)

def get_debug_info(text: str) -> Dict:
    """빠른 디버깅을 위한 편의 함수"""
    processor = TextPreprocessor()
    return processor.get_debug_info(text)

# 테스트 코드 (모듈이 직접 실행될 때)
if __name__ == "__main__":
    print("=== TextPreprocessor 테스트 ===")
    
    # 테스트 케이스들
    test_cases = [
        "[@michel]이 [무기]를 장착했습니다",
        "[#color#red#]경고[#color#] 메시지입니다",
        "캐릭터가 스킬을 사용했습니다",
        "[@system] [아이템] 획득! [@player_name] 레벨업 [#sound#beep#]",
        "[무기] [방어구] [스킬] 모두 업그레이드!",
        "[@variable] [일반태그] [#special#] 혼합 테스트",
        "",  # 빈 문자열 테스트
        None  # None 테스트 (오류 처리)
    ]
    
    processor = TextPreprocessor()
    
    for i, text in enumerate(test_cases, 1):
        print(f"\n📝 테스트 {i}: {repr(text)}")
        try:
            debug_info = processor.get_debug_info(text or "")
            print(f"   🔍 검색대상: '{debug_info['searchable']}'")
            print(f"   🚫 제외태그: {debug_info['excluded_tags']}")
            print(f"   ✅ 포함태그: {debug_info['included_tags']}")
            print(f"   📏 길이감소: {debug_info['length_reduction']}글자")
        except Exception as e:
            print(f"   ❌ 오류: {e}")
    
    print(f"\n📊 전체 처리 통계:")
    stats = processor.get_stats()
    for key, value in stats.items():
        print(f"   {key}: {value}")