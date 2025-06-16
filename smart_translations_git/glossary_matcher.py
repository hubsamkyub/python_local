"""
스마트 용어집 매칭 모듈
용도: Trie 자료구조를 이용한 고속 용어집 검색 및 LLM 프롬프트 최적화
"""
import sqlite3
import re
import time
from typing import List, Dict, Optional, Tuple
from collections import defaultdict

try:
    from text_preprocessor import TextPreprocessor
except ImportError:
    print("Warning: text_preprocessor 모듈을 찾을 수 없습니다. 기본 전처리 사용")
    TextPreprocessor = None

class GlossaryTrie:
    """용어집 검색용 Trie 자료구조"""
    
    def __init__(self):
        self.root = {}
        self.word_end_marker = '__END__'
        self.term_count = 0
    
    def insert(self, korean_word: str, english_word: str, category: str = 'general', string_id: str = None):
        """
        용어를 Trie에 삽입
        
        Args:
            korean_word: 한국어 용어
            english_word: 영어 번역
            category: 용어 카테고리
            string_id: 용어 ID (선택사항)
        """
        if not korean_word or not english_word:
            return
        
        node = self.root
        for char in korean_word:
            if char not in node:
                node[char] = {}
            node = node[char]
        
        # 용어 정보 저장
        node[self.word_end_marker] = {
            'korean': korean_word,
            'english': english_word,
            'category': category,
            'string_id': string_id,
            'frequency': 0  # 사용 빈도 (향후 확장)
        }
        self.term_count += 1
    
    # glossary_matcher.py (일부)
    def search_in_text(self, text: str, max_results: int = 20) -> List[Dict]:
        """
        텍스트에서 모든 매칭되는 용어 찾기 (수정: 최장 일치 보장)
        """
        if not text:
            return []
        
        matches = []
        text_len = len(text)
        
        for i in range(text_len):
            node = self.root
            # --- 이 부분이 수정되었습니다 ---
            longest_match_at_i = None
            j = i
            while j < text_len and text[j] in node:
                node = node[text[j]]
                # 단어가 완성될 때마다, 일단 '가장 긴 후보'로 저장해둡니다.
                if self.word_end_marker in node:
                    word_info = node[self.word_end_marker]
                    longest_match_at_i = {
                        'korean': word_info['korean'],
                        'english': word_info['english'],
                        'category': word_info['category'],
                        'string_id': word_info['string_id'],
                        'start_pos': i,
                        'end_pos': j + 1,
                        'length': (j + 1) - i
                    }
                j += 1
            
            # 현재 시작 위치(i)에서 찾은 가장 긴 매칭만 최종 결과에 추가합니다.
            if longest_match_at_i:
                matches.append(longest_match_at_i)
        
        # 겹치는 매칭 처리 및 정렬
        filtered_matches = self._remove_overlapping_matches(matches)
        
        return filtered_matches[:max_results]
    
    def _remove_overlapping_matches(self, matches: List[Dict]) -> List[Dict]:
        """
        겹치는 매칭에서 가장 긴 것 우선 선택 + 복합어 보호
        """
        if not matches:
            return []
        
        # 1. 복합어 내 짧은 매칭 제거 (대장정 → 대장 문제 해결)
        filtered_matches = []
        for match in matches:
            # 2글자 이하 매칭은 제외 (일단 안전하게)
            if len(match['korean']) <= 2:
                print(f"⚠️ 짧은 매칭 제외: '{match['korean']}'")
                continue
            filtered_matches.append(match)
        
        # 2. 길이순으로 정렬 (긴 것 우선)
        sorted_matches = sorted(filtered_matches, key=lambda x: x['length'], reverse=True)
        
        final_matches = []
        used_positions = set()
        
        for match in sorted_matches:
            match_positions = set(range(match['start_pos'], match['end_pos']))
            
            if not match_positions.intersection(used_positions):
                final_matches.append(match)
                used_positions.update(match_positions)
        
        return sorted(final_matches, key=lambda x: x['start_pos'])

    def get_stats(self) -> Dict:
        """Trie 통계 정보 반환"""
        def count_nodes(node):
            count = 1
            for key, child in node.items():
                if key != self.word_end_marker and isinstance(child, dict):
                    count += count_nodes(child)
            return count
        
        return {
            'total_terms': self.term_count,
            'total_nodes': count_nodes(self.root),
            'memory_efficiency': self.term_count / max(1, count_nodes(self.root))
        }

class SmartGlossaryMatcher:
    """스마트 용어집 매칭 메인 클래스"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.preprocessor = TextPreprocessor() if TextPreprocessor else None
        self.trie = GlossaryTrie()
        self.exact_matches = {}  # 빠른 완전 매칭용
        self.cache = {}  # 검색 결과 캐싱
        
        # 성능 설정
        self.max_cache_size = 1000
        self.max_terms_per_request = 12
        self.cache_hit_count = 0
        self.cache_miss_count = 0
        
        # 카테고리별 우선순위
        self.category_priorities = {
            'character': 10,    # 캐릭터/인물명 최우선
            'item': 8,         # 아이템/장비명
            'place': 7,        # 장소명
            'skill': 6,        # 스킬/능력명
            'ui': 5,           # UI 요소
            'general': 4       # 일반 용어
        }
        
        self._load_glossary()
    

    def _load_glossary(self):
        """DB에서 용어집 로드 및 Trie 구축 (STRING_ID 제거 대응)"""
        start_time = time.time()
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # STRING_ID 없이 용어집 로드
            cursor.execute("""
                SELECT kr, en 
                FROM glossary 
                WHERE kr IS NOT NULL AND en IS NOT NULL 
                AND TRIM(kr) != '' AND TRIM(en) != ''
                ORDER BY LENGTH(kr) DESC
            """)
            
            load_count = 0
            error_count = 0
            
            for kr, en in cursor.fetchall():
                try:
                    kr = kr.strip()
                    en = en.strip()
                    
                    if not kr or not en:
                        continue
                    
                    # 카테고리 자동 분류 (STRING_ID 없이)
                    category = self._categorize_term_without_id(kr)
                    
                    # Trie에 삽입 (string_id를 None으로)
                    self.trie.insert(kr, en, category, string_id=None)
                    
                    # 빠른 완전 매칭을 위한 해시맵
                    self.exact_matches[kr] = {
                        'english': en,
                        'category': category,
                        'string_id': None
                    }
                    
                    load_count += 1
                    
                except Exception as e:
                    error_count += 1
                    if error_count <= 5:
                        print(f"용어 로드 오류: {kr} -> {en}, 오류: {e}")
            
            conn.close()
            
            load_time = time.time() - start_time
            print(f"✅ 용어집 로드 완료: {load_count}개 용어, {load_time:.2f}초")
            
            if error_count > 0:
                print(f"⚠️ 로드 중 오류: {error_count}개")
            
            # Trie 통계 출력
            trie_stats = self.trie.get_stats()
            print(f"📊 Trie 통계: {trie_stats['total_terms']}개 용어, {trie_stats['total_nodes']}개 노드")
            
        except Exception as e:
            print(f"❌ 용어집 로드 실패: {e}")

    def _categorize_term_without_id(self, korean_term: str) -> str:
        """
        STRING_ID 없이 용어 카테고리 자동 분류 (한국어 기준)
        
        Args:
            korean_term: 한국어 용어
            
        Returns:
            카테고리 문자열
        """
        korean_lower = korean_term.lower()
        
        # 캐릭터/인물명 감지
        character_keywords = ['이름', '캐릭터', '인물', '왕', '공주', '왕자', '기사']
        if any(keyword in korean_lower for keyword in character_keywords):
            return 'character'
        
        # 아이템/장비 감지
        item_keywords = ['무기', '방어구', '아이템', '장비', '도구', '검', '갑옷', '방패']
        if any(keyword in korean_lower for keyword in item_keywords):
            return 'item'
        
        # 장소 감지
        place_keywords = ['성', '마을', '던전', '지역', '장소', '숲', '산', '강']
        if any(keyword in korean_lower for keyword in place_keywords):
            return 'place'
        
        # 스킬 감지
        skill_keywords = ['스킬', '능력', '마법', '기술', '주문', '치료']
        if any(keyword in korean_lower for keyword in skill_keywords):
            return 'skill'
        
        # UI 요소 감지
        ui_keywords = ['버튼', '메뉴', '창', '팝업', '인터페이스', '화면', '목록']
        if any(keyword in korean_lower for keyword in ui_keywords):
            return 'ui'
        
        return 'general'

    def _categorize_term(self, string_id: str, korean_term: str) -> str:
        """
        용어 카테고리 자동 분류
        
        Args:
            string_id: 용어 식별자
            korean_term: 한국어 용어
            
        Returns:
            카테고리 문자열
        """
        if not string_id:
            return 'general'
        
        string_id_lower = string_id.lower()
        korean_lower = korean_term.lower()
        
        # 캐릭터/인물명 감지
        character_keywords = ['name', 'character', 'npc', '@', '이름', '캐릭터', '인물']
        if any(keyword in string_id_lower for keyword in character_keywords[:4]):
            return 'character'
        if any(keyword in korean_lower for keyword in character_keywords[4:]):
            return 'character'
        
        # 아이템/장비 감지
        item_keywords = ['item', 'weapon', 'armor', 'equipment', '무기', '방어구', '아이템', '장비', '도구']
        if any(keyword in string_id_lower for keyword in item_keywords[:4]):
            return 'item'
        if any(keyword in korean_lower for keyword in item_keywords[4:]):
            return 'item'
        
        # 장소 감지
        place_keywords = ['place', 'location', 'area', '성', '마을', '던전', '지역', '장소']
        if any(keyword in string_id_lower for keyword in place_keywords[:3]):
            return 'place'
        if any(keyword in korean_lower for keyword in place_keywords[3:]):
            return 'place'
        
        # 스킬 감지
        skill_keywords = ['skill', 'ability', 'magic', '스킬', '능력', '마법', '기술']
        if any(keyword in string_id_lower for keyword in skill_keywords[:3]):
            return 'skill'
        if any(keyword in korean_lower for keyword in skill_keywords[3:]):
            return 'skill'
        
        # UI 요소 감지
        ui_keywords = ['ui', 'button', 'menu', 'popup', '버튼', '메뉴', '창', '팝업', '인터페이스']
        if any(keyword in string_id_lower for keyword in ui_keywords[:4]):
            return 'ui'
        if any(keyword in korean_lower for keyword in ui_keywords[4:]):
            return 'ui'
        
        return 'general'
    
    def find_relevant_terms(self, text: str, max_terms: int = None) -> List[str]:
        """
        텍스트에서 관련 용어 찾기 (메인 함수)
        
        Args:
            text: 검색할 텍스트
            max_terms: 최대 반환할 용어 수
            
        Returns:
            "한국어→영어" 형태의 용어 리스트
        """
        if not text or not isinstance(text, str):
            return []
        
        max_terms = max_terms or self.max_terms_per_request
        
        # 캐시 확인
        cache_key = hash(text[:150])  # 처음 150자로 캐시키 생성
        if cache_key in self.cache:
            self.cache_hit_count += 1
            return self.cache[cache_key]
        
        self.cache_miss_count += 1
        
        # 1. 텍스트 전처리
        if self.preprocessor:
            searchable_text = self.preprocessor.extract_searchable_text(text)
        else:
            # 전처리 모듈이 없으면 기본 처리
            searchable_text = re.sub(r'\[@[^\]]*\]', ' ', text)  # [@...] 제거
            searchable_text = re.sub(r'\[#[^#]*#\]', ' ', searchable_text)  # [#...#] 제거
            searchable_text = re.sub(r'\s+', ' ', searchable_text).strip()
        
        if not searchable_text:
            return []
        
        # 2. Trie 기반 용어 검색
        trie_matches = self.trie.search_in_text(searchable_text)
        
        # 3. 우선순위 적용 및 정렬
        prioritized_matches = self._prioritize_matches(trie_matches)
        
        # 4. 결과 포맷팅
        relevant_terms = []
        for match in prioritized_matches[:max_terms]:
            formatted_term = f"{match['korean']}→{match['english']}"
            relevant_terms.append(formatted_term)
        
        # 5. 캐싱 (크기 제한)
        if len(self.cache) >= self.max_cache_size:
            # 오래된 캐시 절반 삭제 (FIFO)
            old_keys = list(self.cache.keys())[:self.max_cache_size // 2]
            for key in old_keys:
                del self.cache[key]
        
        self.cache[cache_key] = relevant_terms
        
        return relevant_terms
    
    def _prioritize_matches(self, matches: List[Dict]) -> List[Dict]:
        """
        매칭된 용어들을 우선순위에 따라 정렬
        
        Args:
            matches: 원본 매칭 리스트
            
        Returns:
            우선순위가 적용된 매칭 리스트
        """
        if not matches:
            return []
        
        # 각 매칭에 우선순위 점수 계산
        for match in matches:
            score = 0
            
            # 1. 카테고리별 가중치
            category = match.get('category', 'general')
            score += self.category_priorities.get(category, 4)
            
            # 2. 용어 길이 가중치 (긴 용어 우선)
            score += len(match['korean']) * 0.5
            
            # 3. 특수 보너스
            korean = match['korean']
            
            # 고유명사 보너스 (첫 글자 대문자 등)
            if any(char.isupper() for char in match['english']):
                score += 2
            
            # 자주 사용되는 게임 용어 보너스
            game_terms = ['캐릭터', '스킬', '아이템', '레벨', '경험치']
            if any(term in korean for term in game_terms):
                score += 1
            
            match['priority_score'] = score
        
        # 점수순으로 정렬 (높은 점수 우선)
        return sorted(matches, key=lambda x: x['priority_score'], reverse=True)
    
    def create_enhanced_prompt(self, korean_text: str, base_prompt: str) -> str:
        """
        용어집이 적용된 향상된 프롬프트 생성
        
        Args:
            korean_text: 번역할 한국어 텍스트
            base_prompt: 기본 프롬프트
            
        Returns:
            용어집이 포함된 향상된 프롬프트
        """
        relevant_terms = self.find_relevant_terms(korean_text, max_terms=10)
        
        if not relevant_terms:
            # 관련 용어가 없으면 기본 프롬프트 사용
            return f"{base_prompt}\n\n{korean_text}"
        
        # 용어집 포함 프롬프트 생성
        terms_str = ", ".join(relevant_terms)
        
        enhanced_prompt = f"""{base_prompt}

다음 용어집을 참고하여 번역하세요:
{terms_str}

번역할 텍스트: {korean_text}
번역:"""
        
        return enhanced_prompt
    
    def get_debug_info(self, text: str) -> Dict:
        """
        디버깅 정보 반환
        
        Args:
            text: 분석할 텍스트
            
        Returns:
            상세한 디버깅 정보
        """
        debug_info = {
            'original_text': text,
            'cache_stats': {
                'hits': self.cache_hit_count,
                'misses': self.cache_miss_count,
                'hit_rate': self.cache_hit_count / max(1, self.cache_hit_count + self.cache_miss_count),
                'cache_size': len(self.cache)
            }
        }
        
        if not text:
            return debug_info
        
        # 전처리 정보
        if self.preprocessor:
            preprocess_info = self.preprocessor.get_debug_info(text)
            debug_info.update(preprocess_info)
            searchable_text = preprocess_info['searchable']
        else:
            searchable_text = text
            debug_info['searchable'] = searchable_text
        
        # 매칭 정보
        if searchable_text:
            raw_matches = self.trie.search_in_text(searchable_text)
            prioritized_matches = self._prioritize_matches(raw_matches)
            final_terms = self.find_relevant_terms(text)
            
            debug_info.update({
                'raw_matches': raw_matches,
                'prioritized_matches': prioritized_matches,
                'final_terms': final_terms,
                'match_count': len(raw_matches),
                'final_count': len(final_terms)
            })
        else:
            debug_info.update({
                'raw_matches': [],
                'final_terms': [],
                'match_count': 0,
                'final_count': 0
            })
        
        return debug_info
    
    def get_performance_stats(self) -> Dict:
        """성능 통계 반환"""
        trie_stats = self.trie.get_stats()
        
        return {
            'glossary_terms': trie_stats['total_terms'],
            'trie_nodes': trie_stats['total_nodes'],
            'cache_hits': self.cache_hit_count,
            'cache_misses': self.cache_miss_count,
            'cache_hit_rate': self.cache_hit_count / max(1, self.cache_hit_count + self.cache_miss_count),
            'cache_size': len(self.cache),
            'memory_efficiency': trie_stats['memory_efficiency']
        }

# 편의 함수들
def create_matcher(db_path: str) -> SmartGlossaryMatcher:
    """GlossaryMatcher 인스턴스 생성"""
    return SmartGlossaryMatcher(db_path)

def enhance_prompt_with_glossary(korean_text: str, base_prompt: str, matcher: SmartGlossaryMatcher) -> str:
    """프롬프트에 용어집 정보 추가"""
    return matcher.create_enhanced_prompt(korean_text, base_prompt)

# 테스트 코드
if __name__ == "__main__":
    print("=== SmartGlossaryMatcher 테스트 ===")
    
    # 실제 DB 파일이 있다고 가정한 테스트
    # 실제 사용 시에는 올바른 DB 경로를 제공해야 함
    try:
        matcher = SmartGlossaryMatcher("smart_translations.db")
        
        test_cases = [
            "[@michel]이 [무기]를 장착했습니다",
            "캐릭터가 스킬을 사용했습니다",
            "[#color#red#]경고[#color#] 메시지",
            "[@system] [아이템] 획득!"
        ]
        
        for i, text in enumerate(test_cases, 1):
            print(f"\n📝 테스트 {i}: {text}")
            
            debug_info = matcher.get_debug_info(text)
            print(f"   🔍 검색대상: '{debug_info.get('searchable', 'N/A')}'")
            print(f"   🎯 관련용어: {debug_info.get('final_terms', [])}")
            
            # 프롬프트 생성 테스트
            base_prompt = "다음 텍스트를 영어로 번역하세요:"
            enhanced = matcher.create_enhanced_prompt(text, base_prompt)
            print(f"   📝 프롬프트 길이: {len(enhanced)}글자")
        
        # 성능 통계
        print(f"\n📊 성능 통계:")
        stats = matcher.get_performance_stats()
        for key, value in stats.items():
            print(f"   {key}: {value}")
            
    except Exception as e:
        print(f"❌ 테스트 실패: {e}")
        print("📝 참고: 실제 DB 파일 경로를 확인하세요.")

    def fix_inconsistent_translations(self):
        """용어집의 일관성 없는 번역 수정"""
        print("🔧 용어집 일관성 문제 수정 중...")
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 같은 한국어에 대한 여러 영어 번역 찾기
            cursor.execute("""
                SELECT kr, GROUP_CONCAT(en) as translations, COUNT(DISTINCT en) as count
                FROM glossary 
                GROUP BY kr 
                HAVING COUNT(DISTINCT en) > 1
                ORDER BY count DESC
            """)
            
            inconsistent = cursor.fetchall()
            
            if inconsistent:
                print(f"⚠️ 일관성 없는 번역 {len(inconsistent)}개 발견:")
                for kr, translations, count in inconsistent[:5]:
                    print(f"   '{kr}': {translations}")
            
            # 자동 수정: 각 한국어에 대해 가장 빈번한 영어 번역만 남기기
            fixed_count = 0
            for kr, _, _ in inconsistent:
                # 가장 빈번한 번역 찾기
                cursor.execute("""
                    SELECT en, COUNT(*) as freq 
                    FROM glossary 
                    WHERE kr = ? 
                    GROUP BY en 
                    ORDER BY freq DESC 
                    LIMIT 1
                """, (kr,))
                
                most_frequent = cursor.fetchone()
                if most_frequent:
                    best_en = most_frequent[0]
                    
                    # 다른 번역들 삭제
                    cursor.execute("""
                        DELETE FROM glossary 
                        WHERE kr = ? AND en != ?
                    """, (kr, best_en))
                    
                    fixed_count += cursor.rowcount
            
            conn.commit()
            conn.close()
            
            print(f"✅ 일관성 수정 완료: {fixed_count}개 중복 항목 제거")
            
            # Trie 재구축
            self._load_glossary()
            
        except Exception as e:
            print(f"❌ 일관성 수정 실패: {e}")