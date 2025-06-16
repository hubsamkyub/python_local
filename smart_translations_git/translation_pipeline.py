# translation_pipeline.py - 완전 개선 버전

import re
import time
from typing import List, Dict, Callable, Optional, Tuple
from collections import defaultdict
import json

class ImprovedTranslationPipeline:
    """
    완전히 개선된 번역 파이프라인
    - 재귀 완전 제거
    - 배치 처리로 효율성 극대화  
    - 안정적인 특수 태그 처리
    - 스마트 캐싱
    - 강력한 오류 처리
    """
    
    def __init__(self, glossary_matcher, text_preprocessor):
        self.glossary_matcher = glossary_matcher
        self.text_preprocessor = text_preprocessor
        
        # 캐싱 시스템
        self.translation_cache = {}
        self.glossary_cache = {}
        
        # 배치 처리 설정
        self.batch_size = 15
        self.max_retries = 3
        
        # 통계
        self.stats = {
            'cache_hits': 0,
            'cache_misses': 0,
            'total_translations': 0,
            'batch_count': 0,
            'errors': 0
        }
        
        # 특수 태그 패턴
        self.special_tag_patterns = [
            r'\{[^}]+\}',           # {1}, {player_name}
            r'\[@[^\]]+\]',         # [@variable_name]
            r'\[#[^#]*#\]',         # [#color#red#]
            r'\[[A-Z_]+\]'          # [DIFFICULTY], [MAIN_QUEST]
        ]

    def translate_single_text(self, text: str, llm_callback: Callable, base_prompt: str = None) -> str:
        """
        단일 텍스트 번역 (재귀 없음, 안전함)
        """
        if not text or not text.strip():
            return text
            
        # 캐시 확인
        cache_key = hash(text)
        if cache_key in self.translation_cache:
            self.stats['cache_hits'] += 1
            return self.translation_cache[cache_key]
        
        self.stats['cache_misses'] += 1
        self.stats['total_translations'] += 1
        
        try:
            # 1. 특수 태그 보호
            protected_text, tag_map = self._protect_special_tags(text)
            
            # 2. 용어집 용어 찾기 및 보호
            terms_found, term_map = self._find_and_protect_glossary_terms(protected_text)
            
            # 3. 번역 실행
            if terms_found:
                # 용어가 있는 경우: 스마트 번역
                result = self._translate_with_glossary_terms(
                    terms_found, term_map, tag_map, llm_callback, base_prompt or "Translate to natural English:"
                )
            else:
                # 용어가 없는 경우: 직접 번역
                result = self._direct_translate(protected_text, tag_map, llm_callback, base_prompt)
            
            # 4. 결과 캐싱
            if result and result != text:
                self.translation_cache[cache_key] = result
            
            return result or text
            
        except Exception as e:
            print(f"❌ 번역 오류: {e} (텍스트: {text[:50]}...)")
            self.stats['errors'] += 1
            return text

    def translate_batch(self, texts: List[str], llm_callback: Callable, 
                       progress_callback: Optional[Callable] = None, 
                       base_prompt: str = None) -> List[str]:
        """
        배치 번역 처리 (효율성 극대화)
        """
        if not texts:
            return []
        
        print(f"🚀 배치 번역 시작: {len(texts)}개 항목")
        start_time = time.time()
        
        results = []
        
        # 배치 단위로 처리
        for i in range(0, len(texts), self.batch_size):
            batch = texts[i:i + self.batch_size]
            batch_number = (i // self.batch_size) + 1
            
            print(f"  📦 배치 {batch_number}: {len(batch)}개 처리 중...")
            
            try:
                batch_results = self._process_batch_intelligent(batch, llm_callback, base_prompt)
                results.extend(batch_results)
                
                self.stats['batch_count'] += 1
                
                # 진행률 업데이트
                if progress_callback:
                    progress_callback(min(i + len(batch), len(texts)), len(texts))
                    
            except Exception as e:
                print(f"❌ 배치 {batch_number} 오류: {e}")
                # 개별 처리로 폴백
                for text in batch:
                    try:
                        result = self.translate_single_text(text, llm_callback, base_prompt)
                        results.append(result)
                    except:
                        results.append(text)
        
        elapsed = time.time() - start_time
        print(f"✅ 배치 번역 완료: {elapsed:.2f}초, {len(results)}개 결과")
        print(f"📊 캐시 효율: {self.stats['cache_hits']}/{self.stats['cache_hits'] + self.stats['cache_misses']} hits")
        
        return results

    def _protect_special_tags(self, text: str) -> Tuple[str, Dict[str, str]]:
        """특수 태그를 안전하게 보호"""
        protected_text = text
        tag_map = {}
        counter = 0
        
        for pattern in self.special_tag_patterns:
            matches = re.finditer(pattern, protected_text)
            for match in matches:
                placeholder = f"__SPECIALTAG_{counter}__"
                tag_map[placeholder] = match.group()
                protected_text = protected_text.replace(match.group(), placeholder, 1)
                counter += 1
        
        return protected_text, tag_map

    def _find_and_protect_glossary_terms(self, text: str) -> Tuple[str, Dict[str, Dict]]:
        """용어집 용어 찾기 및 보호 (개선된 알고리즘)"""
        try:
            # 용어집에서 매칭 찾기 (개선된 필터링)
            raw_matches = self.glossary_matcher.trie.search_in_text(text)
            
            # 품질 필터링
            filtered_matches = self._filter_quality_matches(raw_matches, text)
            
            if not filtered_matches:
                return text, {}
            
            # 용어 보호 및 맵핑
            protected_text = text
            term_map = {}
            
            # 긴 용어부터 처리 (겹침 방지)
            sorted_matches = sorted(filtered_matches, key=lambda x: x['length'], reverse=True)
            
            for i, match in enumerate(sorted_matches):
                placeholder = f"__TERM_{i}__"
                term_map[placeholder] = {
                    'korean': match['korean'],
                    'english': match['english'],
                    'original_text': match['korean']
                }
                
                # 첫 번째 발생만 교체
                protected_text = protected_text.replace(match['korean'], placeholder, 1)
            
            return protected_text, term_map
            
        except Exception as e:
            print(f"❌ 용어 찾기 오류: {e}")
            return text, {}

    def _filter_quality_matches(self, matches: List[Dict], text: str) -> List[Dict]:
        """품질 기준으로 매칭 필터링 (대장정→대장 문제 해결)"""
        if not matches:
            return []
        
        filtered = []
        
        for match in matches:
            korean_term = match['korean']
            start_pos = match['start_pos']
            end_pos = match['end_pos']
            
            # 1. 너무 짧은 용어 제외 (1-2글자)
            if len(korean_term) <= 2:
                # 주변 문맥 확인
                before = text[max(0, start_pos-1):start_pos] if start_pos > 0 else ""
                after = text[end_pos:end_pos+1] if end_pos < len(text) else ""
                
                # 한글이 연속으로 이어지면 복합어의 일부로 판단하여 제외
                if (before and self._is_korean_char(before) or 
                    after and self._is_korean_char(after)):
                    print(f"⚠️ 복합어 내 단어 제외: '{korean_term}' in '{text[max(0,start_pos-3):end_pos+3]}'")
                    continue
            
            # 2. 의미없는 매칭 제외
            if korean_term in ['의', '을', '를', '이', '가', '은', '는', '과', '와']:
                continue
            
            # 3. 품질 점수 계산
            quality_score = self._calculate_match_quality(match, text)
            if quality_score < 0.5:
                continue
            
            filtered.append(match)
        
        return filtered

    def _is_korean_char(self, char: str) -> bool:
        """한글 문자 확인"""
        return '가' <= char <= '힣'

    def _calculate_match_quality(self, match: Dict, text: str) -> float:
        """매칭 품질 점수 계산"""
        score = 1.0
        
        # 길이 보너스
        length = len(match['korean'])
        if length >= 3:
            score += 0.3
        elif length == 1:
            score -= 0.4
        
        # 고유명사 보너스
        if any(c.isupper() for c in match['english']):
            score += 0.2
        
        # 카테고리 보너스
        category = match.get('category', 'general')
        if category in ['character', 'item', 'place']:
            score += 0.1
        
        return max(0, min(1, score))

    def _translate_with_glossary_terms(self, protected_text: str, term_map: Dict, 
                                     tag_map: Dict, llm_callback: Callable, base_prompt: str) -> str:
        """용어가 포함된 텍스트의 스마트 번역"""
        
        # 용어 정보 추출
        term_info = []
        for placeholder, term_data in term_map.items():
            term_info.append(f"{term_data['korean']} → {term_data['english']}")
        
        # 최적화된 프롬프트 생성
        prompt = f"""{base_prompt}

Key terms to use exactly as shown:
{'; '.join(term_info)}

Translate naturally while using the exact terms provided:
{protected_text}

Translation:"""
        
        try:
            result = llm_callback(protected_text, prompt)
            
            if not result:
                return self._fallback_assembly(protected_text, term_map, tag_map)
            
            # 플레이스홀더 복원
            final_result = self._restore_placeholders(result, term_map, tag_map)
            return final_result
            
        except Exception as e:
            print(f"❌ 용어 번역 오류: {e}")
            return self._fallback_assembly(protected_text, term_map, tag_map)

    def _direct_translate(self, protected_text: str, tag_map: Dict, 
                         llm_callback: Callable, base_prompt: str) -> str:
        """직접 번역 (용어 없음)"""
        try:
            prompt = f"""{base_prompt}

Preserve all special formatting and translate naturally:
{protected_text}

Translation:"""
            
            result = llm_callback(protected_text, prompt)
            
            if result:
                # 태그만 복원
                return self._restore_tags(result, tag_map)
            else:
                return protected_text
                
        except Exception as e:
            print(f"❌ 직접 번역 오류: {e}")
            return protected_text

    def _fallback_assembly(self, protected_text: str, term_map: Dict, tag_map: Dict) -> str:
        """폴백: 단순 조합"""
        result = protected_text
        
        # 용어 교체
        for placeholder, term_data in term_map.items():
            result = result.replace(placeholder, term_data['english'])
        
        # 태그 복원
        result = self._restore_tags(result, tag_map)
        
        # 기본 정리
        result = ' '.join(result.split())  # 중복 공백 제거
        
        return result

    def _restore_placeholders(self, text: str, term_map: Dict, tag_map: Dict) -> str:
        """모든 플레이스홀더 복원"""
        result = text
        
        # 용어 복원
        for placeholder, term_data in term_map.items():
            if placeholder in result:
                result = result.replace(placeholder, term_data['english'])
        
        # 태그 복원
        result = self._restore_tags(result, tag_map)
        
        return result

    def _restore_tags(self, text: str, tag_map: Dict) -> str:
        """특수 태그 복원"""
        result = text
        for placeholder, original_tag in tag_map.items():
            result = result.replace(placeholder, original_tag)
        return result

    def _process_batch_intelligent(self, batch: List[str], llm_callback: Callable, base_prompt: str) -> List[str]:
        """지능적 배치 처리"""
        
        # 1. 캐시된 결과와 새로 번역할 텍스트 분리
        cached_results = {}
        new_texts = []
        
        for text in batch:
            cache_key = hash(text)
            if cache_key in self.translation_cache:
                cached_results[text] = self.translation_cache[cache_key]
                self.stats['cache_hits'] += 1
            else:
                new_texts.append(text)
                self.stats['cache_misses'] += 1
        
        # 2. 새 텍스트들을 분류
        simple_texts = []  # 용어 없음
        complex_texts = []  # 용어 있음
        
        for text in new_texts:
            protected_text, _ = self._protect_special_tags(text)
            terms_found, _ = self._find_and_protect_glossary_terms(protected_text)
            
            if terms_found == protected_text:  # 용어 없음
                simple_texts.append(text)
            else:  # 용어 있음
                complex_texts.append(text)
        
        # 3. 배치별 번역
        new_results = {}
        
        # 간단한 텍스트들 배치 번역
        if simple_texts:
            simple_results = self._batch_translate_simple(simple_texts, llm_callback, base_prompt)
            new_results.update(simple_results)
        
        # 복잡한 텍스트들 개별 번역
        for text in complex_texts:
            try:
                result = self.translate_single_text(text, llm_callback, base_prompt)
                new_results[text] = result
            except Exception as e:
                print(f"❌ 복잡한 텍스트 번역 실패: {text[:30]}... ({e})")
                new_results[text] = text
        
        # 4. 결과 조합 및 순서 유지
        final_results = []
        for text in batch:
            if text in cached_results:
                final_results.append(cached_results[text])
            elif text in new_results:
                final_results.append(new_results[text])
            else:
                final_results.append(text)
        
        return final_results

    def _batch_translate_simple(self, texts: List[str], llm_callback: Callable, base_prompt: str) -> Dict[str, str]:
        """간단한 텍스트들의 배치 번역"""
        if not texts:
            return {}
        
        try:
            # 배치 프롬프트 생성
            prompt = f"""{base_prompt}

Translate these Korean texts to natural English. Format as:
[1]: translation of first text
[2]: translation of second text
etc.

Texts:"""
            
            for i, text in enumerate(texts, 1):
                prompt += f"\n[{i}]: {text}"
            
            prompt += "\n\nTranslations:"
            
            # 배치 번역 실행
            batch_result = llm_callback("batch_simple", prompt)
            
            # 결과 파싱
            results = self._parse_batch_results(batch_result, texts)
            
            # 캐싱
            for text, translation in results.items():
                if translation and translation != text:
                    cache_key = hash(text)
                    self.translation_cache[cache_key] = translation
            
            return results
            
        except Exception as e:
            print(f"❌ 간단한 배치 번역 오류: {e}")
            return {text: text for text in texts}

    def _parse_batch_results(self, result: str, original_texts: List[str]) -> Dict[str, str]:
        """배치 결과 파싱 (강화된 버전)"""
        results = {}
        
        try:
            # [숫자]: 패턴으로 파싱
            pattern = r'\[(\d+)\]:\s*([^\[\n]*?)(?=\[\d+\]:|$)'
            matches = re.findall(pattern, result, re.DOTALL)
            
            for num_str, translation in matches:
                try:
                    index = int(num_str) - 1
                    if 0 <= index < len(original_texts):
                        cleaned_translation = translation.strip()
                        if cleaned_translation:
                            results[original_texts[index]] = cleaned_translation
                except (ValueError, IndexError):
                    continue
            
            # 결과가 없는 텍스트들은 원본 사용
            for text in original_texts:
                if text not in results:
                    results[text] = text
            
        except Exception as e:
            print(f"❌ 배치 결과 파싱 오류: {e}")
            results = {text: text for text in original_texts}
        
        return results

    def get_statistics(self) -> Dict:
        """번역 통계 반환"""
        total_requests = self.stats['cache_hits'] + self.stats['cache_misses']
        cache_rate = self.stats['cache_hits'] / max(1, total_requests)
        
        return {
            'total_translations': self.stats['total_translations'],
            'cache_hits': self.stats['cache_hits'],
            'cache_misses': self.stats['cache_misses'],
            'cache_hit_rate': f"{cache_rate:.1%}",
            'batch_count': self.stats['batch_count'],
            'errors': self.stats['errors'],
            'cache_size': len(self.translation_cache)
        }

    def clear_cache(self):
        """캐시 초기화"""
        self.translation_cache.clear()
        self.glossary_cache.clear()
        print("🗑️ 번역 캐시 초기화 완료")

    def export_cache(self, filepath: str):
        """캐시를 파일로 내보내기"""
        try:
            cache_data = {
                'translation_cache': {str(k): v for k, v in self.translation_cache.items()},
                'stats': self.stats,
                'timestamp': time.time()
            }
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            
            print(f"💾 번역 캐시 저장 완료: {filepath}")
            
        except Exception as e:
            print(f"❌ 캐시 저장 실패: {e}")

    def import_cache(self, filepath: str):
        """파일에서 캐시 불러오기"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            # 캐시 복원
            translation_cache = cache_data.get('translation_cache', {})
            for k, v in translation_cache.items():
                try:
                    self.translation_cache[int(k)] = v
                except ValueError:
                    continue
            
            print(f"📂 번역 캐시 로드 완료: {len(self.translation_cache)}개 항목")
            
        except Exception as e:
            print(f"❌ 캐시 로드 실패: {e}")


# 편의 함수들
def create_improved_pipeline(glossary_matcher, text_preprocessor) -> ImprovedTranslationPipeline:
    """개선된 파이프라인 생성"""
    return ImprovedTranslationPipeline(glossary_matcher, text_preprocessor)

def safe_translate_texts(texts: List[str], pipeline: ImprovedTranslationPipeline, 
                        llm_callback: Callable, progress_callback: Optional[Callable] = None) -> List[str]:
    """안전한 텍스트 배열 번역"""
    return pipeline.translate_batch(texts, llm_callback, progress_callback)

def safe_translate_dict(texts_dict: Dict[str, str], pipeline: ImprovedTranslationPipeline,
                       llm_callback: Callable, progress_callback: Optional[Callable] = None) -> Dict[str, str]:
    """딕셔너리 형태 텍스트 번역"""
    if not texts_dict:
        return {}
    
    original_texts = list(texts_dict.keys())
    translated_texts = pipeline.translate_batch(original_texts, llm_callback, progress_callback)
    
    return dict(zip(original_texts, translated_texts))

# 테스트 및 검증
if __name__ == "__main__":
    print("=== 개선된 번역 파이프라인 테스트 ===")
    
    # 테스트는 실제 glossary_matcher와 text_preprocessor가 있을 때 실행
    # pipeline = create_improved_pipeline(glossary_matcher, text_preprocessor)
    
    test_texts = [
        "어빈의 대장정",
        "도전의 탑 2층", 
        "미첼의 영웅 소환권",
        "[@dialogue_yellow][쉬움][#] [벨트루나] [메인 퀘스트] 아이다의 모험",
        "어빈과 {1}의 대모험"
    ]
    
    print(f"테스트할 텍스트 {len(test_texts)}개:")
    for i, text in enumerate(test_texts, 1):
        print(f"  {i}. {text}")
    
    print("\n✅ 파이프라인 모듈 로드 완료")
    print("📝 실제 테스트는 main.py에서 실행하세요.")