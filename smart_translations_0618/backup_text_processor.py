import re
import json
from difflib import SequenceMatcher
from collections import defaultdict

# === 기존 HTML 처리 함수들 (글로벌 함수로 유지) ===

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


# === 새로운 TextProcessor 클래스 ===

class TextProcessor:
    def __init__(self, manager):
        self.manager = manager
        self.text_stats = {
            'processed_count': 0,
            'glossary_applied': 0,
            'tags_protected': 0,
            'errors_detected': 0,
            'html_processed': 0
        }
    
    def preprocess_for_translation(self, kr_text, context=None):
        """번역용 텍스트 전처리 - 메인 진입점 (HTML 처리 통합)"""
        result = {
            'original': kr_text,
            'processed': kr_text,
            'glossary_matches': [],
            'protected_elements': [],
            'warnings': [],
            'html_info': {
                'is_html': False,
                'chunks': [],
                'needs_chunking': False
            },
            'metadata': {
                'length': len(kr_text),
                'complexity': 'simple',
                'has_markup': False,
                'has_variables': False,
                'estimated_difficulty': 1.0
            }
        }
        
        try:
            # 1. HTML 감지 및 처리
            self._analyze_html_structure(result)
            
            # 2. 기본 텍스트 정제
            result['processed'] = self._clean_text(kr_text)
            
            # 3. 텍스트 분석
            self._analyze_text_structure(result)
            
            # 4. 용어집 매칭 및 적용 (HTML 고려)
            self._apply_glossary_matches(result)
            
            # 5. 특수 요소 보호
            self._protect_special_elements(result)
            
            # 6. HTML 청킹 처리 (필요한 경우)
            self._handle_html_chunking(result)
            
            # 7. 번역 난이도 추정
            self._estimate_translation_difficulty(result)
            
            # 8. 품질 검사
            self._validate_preprocessed_text(result)
            
            self.text_stats['processed_count'] += 1
            if result['html_info']['is_html']:
                self.text_stats['html_processed'] += 1
            
        except Exception as e:
            result['warnings'].append(f"전처리 오류: {e}")
            self.text_stats['errors_detected'] += 1
        
        return result
    
    def postprocess_translation(self, original_kr, translated_text, target_lang='EN'):
        """번역 후 후처리 (HTML 복원 포함)"""
        result = {
            'original_translation': translated_text,
            'processed_translation': translated_text,
            'fixes_applied': [],
            'quality_score': 1.0,
            'warnings': [],
            'html_restored': False
        }
        
        try:
            # 1. 기본 정제 (앞뒤 공백, 불필요한 문자 제거)
            result['processed_translation'] = self._clean_translated_text(translated_text)
            
            # 2. HTML 복원 처리
            if is_html(original_kr):
                result['processed_translation'] = self._restore_html_structure(
                    original_kr, result['processed_translation']
                )
                result['html_restored'] = True
            
            # 3. 용어집 span 태그 처리
            result['processed_translation'] = self._handle_glossary_spans(
                original_kr, result['processed_translation']
            )
            
            # 4. 용어집 일관성 검사
            self._check_glossary_consistency(original_kr, result)
            
            # 5. 특수 태그 복원 검증
            self._verify_tag_preservation(original_kr, result)
            
            # 6. 번역 품질 검사
            result['quality_score'] = self._assess_translation_quality(original_kr, result['processed_translation'], target_lang)
            
            # 7. 자동 수정 적용
            self._apply_auto_fixes(result)
            
        except Exception as e:
            result['warnings'].append(f"후처리 오류: {e}")
        
        return result
    
    def prepare_for_chunked_translation(self, kr_text, max_len=4500):
        """HTML을 고려한 청킹 번역 준비"""
        if not is_html(kr_text):
            # HTML이 아니면 기본 청킹
            return [kr_text[i:i+max_len] for i in range(0, len(kr_text), max_len)]
        
        # HTML 보존 청킹 (기존 함수 활용)
        chunks = split_text_with_html(kr_text, max_len)
        
        # 각 청크에 대해 전처리 적용
        processed_chunks = []
        for chunk in chunks:
            preprocessed = self.preprocess_for_translation(chunk)
            processed_chunks.append({
                'original': chunk,
                'processed': preprocessed['processed'],
                'metadata': preprocessed['metadata']
            })
        
        return processed_chunks
    
    def merge_chunked_translations(self, chunk_results):
        """청킹된 번역 결과들을 병합"""
        if not chunk_results:
            return ""
        
        if len(chunk_results) == 1:
            return chunk_results[0]
        
        # 여러 청크를 자연스럽게 병합
        merged_text = ""
        for i, chunk_text in enumerate(chunk_results):
            if i > 0:
                # 청크 사이에 적절한 연결 처리
                if not merged_text.endswith(' ') and not chunk_text.startswith(' '):
                    merged_text += " "
            merged_text += chunk_text
        
        return merged_text.strip()
    
    def enhance_with_context(self, kr_text, context_data):
        """컨텍스트 정보를 활용한 텍스트 향상"""
        enhanced = {
            'text': kr_text,
            'context_applied': False,
            'enhancements': []
        }
        
        if not context_data:
            return enhanced
        
        # 1. 화자 정보 활용
        if 'speaker' in context_data:
            speaker_enhancement = self._apply_speaker_context(kr_text, context_data['speaker'])
            if speaker_enhancement:
                enhanced['enhancements'].append(speaker_enhancement)
                enhanced['context_applied'] = True
        
        # 2. 이전 번역 패턴 활용
        if 'previous_translations' in context_data:
            pattern_enhancement = self._apply_translation_patterns(kr_text, context_data['previous_translations'])
            if pattern_enhancement:
                enhanced['enhancements'].append(pattern_enhancement)
        
        # 3. 장르/도메인 정보 활용
        if 'domain' in context_data:
            domain_enhancement = self._apply_domain_knowledge(kr_text, context_data['domain'])
            if domain_enhancement:
                enhanced['enhancements'].append(domain_enhancement)
        
        return enhanced
    
    def batch_preprocess(self, kr_texts):
        """여러 텍스트 일괄 전처리 (HTML 처리 포함)"""
        results = []
        
        # 통계 초기화
        batch_stats = {
            'total': len(kr_texts),
            'successful': 0,
            'failed': 0,
            'glossary_hits': 0,
            'complex_texts': 0,
            'html_texts': 0
        }
        
        for i, kr_text in enumerate(kr_texts):
            try:
                result = self.preprocess_for_translation(kr_text)
                results.append(result)
                
                batch_stats['successful'] += 1
                if result['glossary_matches']:
                    batch_stats['glossary_hits'] += 1
                if result['metadata']['complexity'] != 'simple':
                    batch_stats['complex_texts'] += 1
                if result['html_info']['is_html']:
                    batch_stats['html_texts'] += 1
                    
            except Exception as e:
                results.append({
                    'original': kr_text,
                    'error': str(e),
                    'processed': kr_text
                })
                batch_stats['failed'] += 1
        
        return results, batch_stats
    
    # === HTML 처리 관련 메서드들 ===
    
    def _analyze_html_structure(self, result):
        """HTML 구조 분석"""
        text = result['original']
        html_info = result['html_info']
        
        # HTML 감지
        html_info['is_html'] = is_html(text)
        
        if html_info['is_html']:
            # 길이가 길면 청킹 필요
            if len(text) > 4000:
                html_info['needs_chunking'] = True
                html_info['chunks'] = split_text_with_html(text, 4500)
            
            # HTML 복잡도 반영
            result['metadata']['has_markup'] = True
            if result['metadata']['complexity'] == 'simple':
                result['metadata']['complexity'] = 'markup'
    
    def _handle_html_chunking(self, result):
        """HTML 청킹 처리"""
        html_info = result['html_info']
        
        if html_info['needs_chunking'] and html_info['chunks']:
            # 청킹된 텍스트로 처리된 텍스트 업데이트
            processed_chunks = []
            for chunk in html_info['chunks']:
                processed_chunk = self._clean_text(chunk)
                processed_chunks.append(processed_chunk)
            
            result['processed'] = processed_chunks
            result['warnings'].append(f"텍스트가 {len(html_info['chunks'])}개 청크로 분할되었습니다")
    
    def _restore_html_structure(self, original_kr, translated_text):
        """HTML 구조 복원"""
        try:
            # 기본적인 HTML 태그 복원 로직
            # 원본에서 HTML 태그 추출
            original_tags = re.findall(r'<[^>]+>', original_kr)
            
            # 번역문에 태그가 누락되었으면 적절한 위치에 복원
            for tag in original_tags:
                if tag not in translated_text:
                    # 간단한 복원 로직 (실제로는 더 정교해야 함)
                    if tag.startswith('<span class="glossary-term">'):
                        # 용어집 span 태그는 별도 처리
                        continue
                    # 다른 태그들은 텍스트 끝에 추가 (임시)
                    translated_text += f" {tag}"
            
            return translated_text
            
        except Exception as e:
            print(f"HTML 구조 복원 오류: {e}")
            return translated_text
    
    def _handle_glossary_spans(self, original_kr, translated_text):
        """용어집 span 태그 처리"""
        try:
            # 원본에 glossary-term span이 있었다면
            if 'class="glossary-term"' in original_kr:
                # 기존 함수 활용하여 span 태그 제거
                cleaned_text = unwrap_span(translated_text, "glossary-term")
                
                # 필요하면 번역된 용어에 다시 span 태그 적용
                # (실제 구현에서는 용어집 매칭 로직과 연계)
                return cleaned_text
            
            return translated_text
            
        except Exception as e:
            print(f"Glossary span 처리 오류: {e}")
            return translated_text
    
    # === 기존 내부 헬퍼 메서드들 (HTML 처리 강화) ===
    
    def _clean_text(self, text):
        """기본 텍스트 정제 (HTML 고려)"""
        if not text:
            return ""
        
        # HTML인 경우 태그 내용은 보존
        if is_html(text):
            # HTML 태그는 보존하면서 내용만 정제
            def clean_text_content(match):
                return re.sub(r'\s+', ' ', match.group(0).strip())
            
            # 태그 밖의 텍스트만 정제
            cleaned = re.sub(r'>([^<]+)<', lambda m: f'>{clean_text_content(m)}<', text)
            return cleaned.strip()
        else:
            # 일반 텍스트 정제
            cleaned = text.strip()
            cleaned = re.sub(r'\s+', ' ', cleaned)
            cleaned = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', cleaned)
            return cleaned
    
    def _analyze_text_structure(self, result):
        """텍스트 구조 분석 (HTML 정보 반영)"""
        text = result['processed']
        metadata = result['metadata']
        
        # 이미 HTML 분석에서 markup으로 설정되었으면 유지
        if metadata['complexity'] != 'markup':
            # 길이 기반 복잡도
            if len(text) > 200:
                metadata['complexity'] = 'complex'
            elif len(text) > 100:
                metadata['complexity'] = 'medium'
        
        # 마크업 감지 (HTML 외의 다른 마크업도)
        if re.search(r'\[#[^]]+\]|\{[^}]+\}', text):
            metadata['has_markup'] = True
            if metadata['complexity'] == 'simple':
                metadata['complexity'] = 'markup'
        
        # 변수/플레이스홀더 감지
        if re.search(r'\{[^}]*\}|%[sd]|\$\w+', text):
            metadata['has_variables'] = True
        
        # HTML 특수 처리
        if result['html_info']['is_html']:
            metadata['has_markup'] = True
            # HTML 청킹이 필요하면 복잡도 증가
            if result['html_info']['needs_chunking']:
                metadata['complexity'] = 'complex'
    
    def _apply_glossary_matches(self, result):
        """용어집 매칭 및 적용 (HTML 고려)"""
        if not self.manager.glossary:
            return
        
        text = result['processed']
        matches = []
        
        # HTML인 경우 태그 내용은 제외하고 매칭
        if is_html(text):
            # HTML 태그를 제외한 순수 텍스트에서만 용어집 매칭
            text_only = re.sub(r'<[^>]*>', '', text)
        else:
            text_only = text
        
        # 용어집 항목들을 길이 순으로 정렬 (긴 것부터 - 중복 매칭 방지)
        sorted_terms = sorted(self.manager.glossary.items(), 
                            key=lambda x: len(x[0]), reverse=True)
        
        for kr_term, translations in sorted_terms:
            if kr_term in text_only:
                # 정확한 위치와 컨텍스트 정보 수집
                for match in re.finditer(re.escape(kr_term), text_only):
                    match_info = {
                        'term': kr_term,
                        'position': match.start(),
                        'end_position': match.end(),
                        'translations': translations,
                        'context_before': text_only[max(0, match.start()-10):match.start()],
                        'context_after': text_only[match.end():match.end()+10],
                        'confidence': self._calculate_glossary_confidence(kr_term, text_only, match.start()),
                        'in_html': is_html(text)
                    }
                    matches.append(match_info)
        
        # 겹치는 매칭 제거 (더 긴 용어 우선)
        filtered_matches = self._filter_overlapping_matches(matches)
        result['glossary_matches'] = filtered_matches
        
        if filtered_matches:
            self.text_stats['glossary_applied'] += 1
    
    def _protect_special_elements(self, result):
        """특수 요소 보호 (HTML 태그 포함)"""
        text = result['processed']
        protected_elements = []
        
        # 보호할 패턴들 (HTML 태그 추가)
        protection_patterns = [
            (r'<[^>]+>', 'html_tag'),              # HTML 태그 (추가)
            (r'\[#[^]]+\]', 'color_tag'),          # 색상 태그
            (r'\{[^}]+\}', 'variable'),            # 변수
            (r'\\[nt]', 'escape_char'),            # 이스케이프 문자
            (r'[@]\w+', 'mention'),                # 멘션
            (r'#\w+', 'hashtag'),                  # 해시태그
            (r'\d+%', 'percentage'),               # 퍼센트
            (r'\$\d+', 'currency')                 # 통화
        ]
        
        for pattern, element_type in protection_patterns:
            for match in re.finditer(pattern, text):
                element_info = {
                    'type': element_type,
                    'content': match.group(),
                    'position': match.start(),
                    'end_position': match.end(),
                    'placeholder': f'<PROTECTED_{len(protected_elements)}>'
                }
                protected_elements.append(element_info)
        
        result['protected_elements'] = protected_elements
        
        if protected_elements:
            self.text_stats['tags_protected'] += len(protected_elements)
    
    def _estimate_translation_difficulty(self, result):
        """번역 난이도 추정 (HTML 복잡도 반영)"""
        metadata = result['metadata']
        difficulty = 1.0  # 기본 난이도
        
        # 길이에 따른 난이도 증가
        length = metadata['length']
        if length > 200:
            difficulty += 0.5
        elif length > 100:
            difficulty += 0.3
        
        # 복잡도에 따른 난이도 증가
        complexity_scores = {
            'simple': 0,
            'medium': 0.3,
            'complex': 0.6,
            'markup': 0.8
        }
        difficulty += complexity_scores.get(metadata['complexity'], 0)
        
        # HTML 특수 처리
        if result['html_info']['is_html']:
            difficulty += 0.4
            if result['html_info']['needs_chunking']:
                difficulty += 0.3  # 청킹이 필요하면 추가 난이도
        
        # 특수 요소에 따른 난이도 증가
        if metadata['has_markup']:
            difficulty += 0.4
        if metadata['has_variables']:
            difficulty += 0.3
        
        # 용어집 매칭이 많으면 난이도 감소
        if len(result['glossary_matches']) > 3:
            difficulty -= 0.2
        
        metadata['estimated_difficulty'] = max(1.0, min(3.0, difficulty))
    
    def _validate_preprocessed_text(self, result):
        """전처리된 텍스트 검증 (HTML 유효성 포함)"""
        text = result['processed']
        warnings = result['warnings']
        
        # 1. 빈 텍스트 검사
        if not text.strip():
            warnings.append("텍스트가 비어있습니다")
        
        # 2. 너무 긴 텍스트 검사
        if len(text) > 1000:
            warnings.append("텍스트가 너무 깁니다 (1000자 초과)")
        
        # 3. HTML 유효성 검사
        if result['html_info']['is_html']:
            if not self._validate_html_tags(text):
                warnings.append("HTML 태그가 올바르지 않습니다")
        
        # 4. 불완전한 태그 검사 (기존 로직)
        open_brackets = text.count('[')
        close_brackets = text.count(']')
        if open_brackets != close_brackets:
            warnings.append("불완전한 대괄호 태그가 감지되었습니다")
        
        open_braces = text.count('{')
        close_braces = text.count('}')
        if open_braces != close_braces:
            warnings.append("불완전한 중괄호 태그가 감지되었습니다")
    
    def _validate_html_tags(self, text):
        """HTML 태그 유효성 검사"""
        try:
            # 기본적인 HTML 태그 쌍 검사
            tag_stack = []
            tag_pattern = r'<(/?)([^>\s]+)[^>]*>'
            
            for match in re.finditer(tag_pattern, text):
                is_closing = bool(match.group(1))
                tag_name = match.group(2).lower()
                
                # 자체 닫힌 태그들 (br, img 등)
                self_closing_tags = {'br', 'img', 'hr', 'input', 'meta', 'link'}
                
                if tag_name in self_closing_tags:
                    continue
                
                if is_closing:
                    if not tag_stack or tag_stack[-1] != tag_name:
                        return False
                    tag_stack.pop()
                else:
                    tag_stack.append(tag_name)
            
            return len(tag_stack) == 0
            
        except Exception:
            return True  # 검사 실패 시 통과로 처리
    
    # === 기존 메서드들 (변경사항 없음) ===
    
    def _clean_translated_text(self, text):
        """번역된 텍스트 정제"""
        if not text:
            return ""
        
        # 1. 기본 정제
        cleaned = text.strip()
        
        # 2. 연속된 공백 정리
        cleaned = re.sub(r'\s+', ' ', cleaned)
        
        # 3. 불필요한 따옴표 제거 (번역 시 종종 추가됨)
        if cleaned.startswith('"') and cleaned.endswith('"'):
            cleaned = cleaned[1:-1]
        if cleaned.startswith("'") and cleaned.endswith("'"):
            cleaned = cleaned[1:-1]
        
        # 4. 문장 부호 정리
        cleaned = re.sub(r'\s+([.!?:;,])', r'\1', cleaned)
        cleaned = re.sub(r'([.!?])\s*$', r'\1', cleaned)
        
        return cleaned
    
    def _check_glossary_consistency(self, original_kr, result):
        """용어집 일관성 검사"""
        if not self.manager.glossary:
            return
        
        original_text = original_kr
        translated_text = result['processed_translation']
        
        # HTML 태그 제거 후 비교
        if is_html(original_text):
            original_text = re.sub(r'<[^>]*>', '', original_text)
        if is_html(translated_text):
            translated_text = re.sub(r'<[^>]*>', '', translated_text)
        
        inconsistencies = []
        
        for kr_term, translations in self.manager.glossary.items():
            if kr_term in original_text:
                expected_en = translations.get('EN', '')
                if expected_en and expected_en.lower() not in translated_text.lower():
                    inconsistencies.append({
                        'kr_term': kr_term,
                        'expected_en': expected_en,
                        'issue': 'glossary_term_missing'
                    })
        
        if inconsistencies:
            result['warnings'].extend([f"용어집 불일치: {inc['kr_term']} -> {inc['expected_en']}" 
                                     for inc in inconsistencies])
    
    def _verify_tag_preservation(self, original_kr, result):
        """특수 태그 보존 검증 (HTML 포함)"""
        original_tags = self._extract_tags(original_kr)
        translated_tags = self._extract_tags(result['processed_translation'])
        
        if len(original_tags) != len(translated_tags):
            result['warnings'].append(f"태그 개수 불일치: 원본 {len(original_tags)}개, 번역 {len(translated_tags)}개")
        
        # 태그 내용 비교
        for orig_tag in original_tags:
            if orig_tag not in translated_tags:
                result['warnings'].append(f"누락된 태그: {orig_tag}")
    
    def _assess_translation_quality(self, original_kr, translated_text, target_lang):
        """번역 품질 평가"""
        if not translated_text:
            return 0.0
        
        quality_score = 1.0
        
        # HTML 태그 제거 후 길이 비교
        orig_text_only = re.sub(r'<[^>]*>', '', original_kr) if is_html(original_kr) else original_kr
        trans_text_only = re.sub(r'<[^>]*>', '', translated_text) if is_html(translated_text) else translated_text
        
        # 1. 길이 비율 검사
        length_ratio = len(trans_text_only) / len(orig_text_only) if orig_text_only else 0
        if target_lang == 'EN':
            if length_ratio < 0.5 or length_ratio > 2.0:
                quality_score -= 0.3
        
        # 2. 반복 패턴 검사
        if self._has_repetitive_patterns(trans_text_only):
            quality_score -= 0.2
        
        # 3. 불완전한 문장 검사
        if not self._is_complete_sentence(trans_text_only):
            quality_score -= 0.1
        
        # 4. 언어별 특수 검사
        if target_lang == 'EN':
            quality_score *= self._assess_english_quality(trans_text_only)
        
        return max(0.0, min(1.0, quality_score))
    
    def _apply_auto_fixes(self, result):
        """자동 수정 적용"""
        text = result['processed_translation']
        fixes = []
        
        # HTML이 아닌 경우에만 기본 수정 적용
        if not is_html(text):
            # 1. 대소문자 수정
            if text and not text[0].isupper():
                text = text[0].upper() + text[1:]
                fixes.append("첫 글자 대문자화")
            
            # 2. 마침표 추가 (필요한 경우)
            if text and text[-1] not in '.!?':
                if len(text.split()) > 3:
                    text += '.'
                    fixes.append("마침표 추가")
        
        # 3. 이중 공백 제거 (HTML이든 아니든 적용)
        original_spaces = text.count('  ')
        text = re.sub(r'\s+', ' ', text)
        if text.count('  ') < original_spaces:
            fixes.append("이중 공백 제거")
        
        result['processed_translation'] = text
        result['fixes_applied'] = fixes
    
    # === 컨텍스트 관련 헬퍼 메서드들 (변경사항 없음) ===
    
    def _apply_speaker_context(self, text, speaker_info):
        """화자 컨텍스트 적용"""
        if not speaker_info:
            return None
        
        enhancement = {
            'type': 'speaker_context',
            'speaker': speaker_info.get('name', ''),
            'suggestions': []
        }
        
        if 'tone' in speaker_info:
            tone = speaker_info['tone']
            if tone in ['formal', '정중한']:
                enhancement['suggestions'].append("정중한 어조로 번역")
            elif tone in ['casual', '친근한']:
                enhancement['suggestions'].append("캐주얼한 어조로 번역")
        
        return enhancement if enhancement['suggestions'] else None
    
    def _apply_translation_patterns(self, text, previous_translations):
        """이전 번역 패턴 적용"""
        patterns = []
        
        for prev in previous_translations:
            if 'pattern' in prev:
                patterns.append(prev['pattern'])
        
        if patterns:
            return {
                'type': 'translation_pattern',
                'patterns': patterns
            }
        return None
    
    def _apply_domain_knowledge(self, text, domain):
        """도메인 지식 적용"""
        domain_hints = {
            'game': ['게임 용어 우선 사용', '플레이어 중심 표현'],
            'technical': ['기술 용어 정확성 중시', '명확한 표현'],
            'casual': ['친근한 표현 사용', '간결한 문체']
        }
        
        hints = domain_hints.get(domain, [])
        if hints:
            return {
                'type': 'domain_knowledge',
                'domain': domain,
                'hints': hints
            }
        return None
    
    # === 유틸리티 메서드들 (변경사항 없음) ===
    
    def _calculate_glossary_confidence(self, term, text, position):
        """용어집 매칭 신뢰도 계산"""
        is_word_boundary = (position == 0 or not text[position-1].isalnum()) and \
                          (position + len(term) >= len(text) or not text[position + len(term)].isalnum())
        
        confidence = 0.8 if is_word_boundary else 0.6
        
        if len(term) >= 4:
            confidence += 0.1
        
        return min(1.0, confidence)
    
    def _filter_overlapping_matches(self, matches):
        """겹치는 매칭 필터링"""
        if not matches:
            return matches
        
        sorted_matches = sorted(matches, key=lambda x: x['position'])
        filtered = []
        
        for match in sorted_matches:
            overlaps = False
            for existing in filtered:
                if (match['position'] < existing['end_position'] and 
                    match['end_position'] > existing['position']):
                    overlaps = True
                    break
            
            if not overlaps:
                filtered.append(match)
        
        return filtered
    
    def _extract_tags(self, text):
        """텍스트에서 태그 추출 (HTML 태그 포함)"""
        patterns = [
            r'<[^>]+>',      # HTML 태그 (추가)
            r'\[#[^]]+\]',   # 색상 태그
            r'\{[^}]+\}'     # 변수
        ]
        
        tags = []
        for pattern in patterns:
            tags.extend(re.findall(pattern, text))
        
        return tags
    
    def _has_repetitive_patterns(self, text):
        """반복 패턴 감지"""
        words = text.split()
        if len(words) < 4:
            return False
        
        for i in range(len(words) - 1):
            if words[i] == words[i + 1]:
                return True
        
        return False
    
    def _is_complete_sentence(self, text):
        """완전한 문장인지 검사"""
        if not text:
            return False
        
        if len(text.strip()) < 2:
            return False
        
        words = text.split()
        if len(words) < 1:
            return False
        
        return True
    
    def _assess_english_quality(self, text):
        """영어 번역 품질 평가"""
        quality = 1.0
        
        if not re.search(r'[a-zA-Z]', text):
            quality -= 0.5
        
        error_patterns = [
            r'\ba a\b',
            r'\bthe the\b',
            r'\s{2,}',
        ]
        
        for pattern in error_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                quality -= 0.1
        
        return max(0.1, quality)

    def get_processing_stats(self):
        """전처리 통계 반환 (HTML 처리 통계 포함)"""
        return self.text_stats.copy()

    def reset_stats(self):
        """통계 초기화"""
        self.text_stats = {
            'processed_count': 0,
            'glossary_applied': 0,
            'tags_protected': 0,
            'errors_detected': 0,
            'html_processed': 0
        }