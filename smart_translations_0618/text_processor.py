# 파일명: text_processor.py

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
            'template_text': None,
            'placeholder_map': {},
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
            self._analyze_html_structure(result)
            result['processed'] = self._clean_text(kr_text)
            self._analyze_text_structure(result)
            self._apply_glossary_matches(result)
            if result['glossary_matches']:
                template, p_map = self._create_placeholders(result['processed'], result['glossary_matches'])
                result['template_text'] = template
                result['placeholder_map'] = p_map
            self._protect_special_elements(result)
            self._handle_html_chunking(result)
            self._estimate_translation_difficulty(result)
            self._validate_preprocessed_text(result)
            self.text_stats['processed_count'] += 1
            if result['html_info']['is_html']: self.text_stats['html_processed'] += 1
        except Exception as e:
            result['warnings'].append(f"전처리 오류: {e}")
            self.text_stats['errors_detected'] += 1
        return result
    
    # === [추가된 부분] === 
    def reassemble_from_placeholders(self, template_text, placeholder_map, translated_terms_dict):
        if not template_text: return ""
        reassembled_sentence = template_text
        for placeholder, original_term in placeholder_map.items():
            translated_term = translated_terms_dict.get(original_term, f"[번역실패:{original_term}]")
            reassembled_sentence = reassembled_sentence.replace(placeholder, str(translated_term))
        return reassembled_sentence

    def _create_placeholders(self, text, matches):
        if not matches: return text, {}
        placeholder_map, template_text = {}, text
        sorted_matches = sorted(matches, key=lambda x: x['position'], reverse=True)
        placeholder_idx = 0
        for match in sorted_matches:
            placeholder = f"{{{{term_{placeholder_idx}}}}}"
            start, end = match['position'], match['end_position']
            placeholder_map[placeholder] = match['term']
            template_text = template_text[:start] + placeholder + template_text[end:]
            placeholder_idx += 1
        return template_text, placeholder_map


    def _apply_glossary_matches(self, result):
        """용어집 매칭 및 적용 (조사 및 단어 경계 처리 강화 + 상세 로그)"""
        if not self.manager.glossary:
            print("🐞 DEBUG: _apply_glossary_matches - 용어집(self.manager.glossary)이 비어있습니다.")
            return

        print(f"--- 🐞 용어집 매칭 시작: \"{result['processed'][:50]}...\" ---")
        print(f"🐞 DEBUG: 로드된 총 용어집 개수: {len(self.manager.glossary)}")
        
        text_only = re.sub(r'<[^>]*>', '', result['processed']) if is_html(result['processed']) else result['processed']
        sorted_terms = sorted(self.manager.glossary.keys(), key=len, reverse=True)
        
        matches = []
        found_ranges = set()
        
        # 한국어 조사를 판별하기 위한 리스트 (더 많은 조사 추가)
        josa_list = ['은', '는', '이', '가', '을', '를', '의', '과', '와', '에', '에서', '에게', '께', '으로', '로', '다', '만', '도', '뿐', '까지', '부터', '마저', '조차', '하고', '이며', '이라', '여']
        
        for kr_term in sorted_terms:
            try:
                term_pattern = re.compile(re.escape(kr_term))
                for match in term_pattern.finditer(text_only):
                    start, end = match.start(), match.end()
                    
                    is_overlapping = any(max(start, r_start) < min(end, r_end) for r_start, r_end in found_ranges)
                    if is_overlapping:
                        continue

                    # 단어 경계 확인 (스마트 매칭)
                    next_char = text_only[end] if end < len(text_only) else ' '
                    is_boundary = True # 기본적으로 경계가 맞는다고 가정
                    
                    # 뒤에 다른 한글이 바로 붙어있으면 더 큰 단어의 일부일 가능성 체크
                    if '가' <= next_char <= '힣':
                        # 뒤 글자가 조사가 아니면 경계가 아니라고 판단 (매칭 무시)
                        if next_char not in josa_list:
                            is_boundary = False
                    
                    # --- 상세 디버깅 로그 ---
                    print(f"  - 검사 중: '{kr_term}' (위치:{start}-{end}), 뒤따르는 문자: '{next_char}', 단어 경계: {is_boundary}")

                    if not is_boundary:
                        print(f"    -> 무시: '{kr_term}'은 더 큰 단어의 일부로 판단됨.")
                        continue

                    print(f"    -> ✅ 매칭 성공: '{kr_term}'")
                    matches.append({
                        'term': kr_term,
                        'position': start,
                        'end_position': end,
                        'category': self.manager.glossary[kr_term].get('category', 'etc'),
                        'translations': self.manager.glossary[kr_term]
                    })
                    found_ranges.add((start, end))

            except re.error as e:
                print(f"⚠️ 정규식 오류: 용어 '{kr_term}' 처리 중 오류 발생 - {e}")
                continue

        result['glossary_matches'] = sorted(matches, key=lambda x: x['position'])
        print(f"🐞 DEBUG: 최종 매칭된 용어: {[m['term'] for m in result['glossary_matches']]}")
        print("--- 🐞 용어집 매칭 종료 ---")

        if result['glossary_matches']:
            self.text_stats['glossary_applied'] += 1
    

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
            result['processed_translation'] = self._clean_translated_text(translated_text)
            if is_html(original_kr):
                result['processed_translation'] = self._restore_html_structure(original_kr, result['processed_translation'])
                result['html_restored'] = True
            
            result['processed_translation'] = self._handle_glossary_spans(original_kr, result['processed_translation'])
            self._check_glossary_consistency(original_kr, result)
            self._verify_tag_preservation(original_kr, result)
            result['quality_score'] = self._assess_translation_quality(original_kr, result['processed_translation'], target_lang)
            self._apply_auto_fixes(result)
            
        except Exception as e:
            result['warnings'].append(f"후처리 오류: {e}")
        
        return result
    
    def prepare_for_chunked_translation(self, kr_text, max_len=4500):
        """HTML을 고려한 청킹 번역 준비"""
        if not is_html(kr_text):
            return [kr_text[i:i+max_len] for i in range(0, len(kr_text), max_len)]
        
        chunks = split_text_with_html(kr_text, max_len)
        
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
        
        merged_text = ""
        for i, chunk_text in enumerate(chunk_results):
            if i > 0 and not merged_text.endswith(' ') and not chunk_text.startswith(' '):
                merged_text += " "
            merged_text += chunk_text
        
        return merged_text.strip()
    
    def enhance_with_context(self, kr_text, context_data):
        """컨텍스트 정보를 활용한 텍스트 향상"""
        enhanced = {'text': kr_text, 'context_applied': False, 'enhancements': []}
        if not context_data: return enhanced
        
        if 'speaker' in context_data:
            speaker_enhancement = self._apply_speaker_context(kr_text, context_data['speaker'])
            if speaker_enhancement:
                enhanced['enhancements'].append(speaker_enhancement)
                enhanced['context_applied'] = True
        
        if 'previous_translations' in context_data:
            pattern_enhancement = self._apply_translation_patterns(kr_text, context_data['previous_translations'])
            if pattern_enhancement: enhanced['enhancements'].append(pattern_enhancement)
        
        if 'domain' in context_data:
            domain_enhancement = self._apply_domain_knowledge(kr_text, context_data['domain'])
            if domain_enhancement: enhanced['enhancements'].append(domain_enhancement)
        
        return enhanced
    
    def batch_preprocess(self, kr_texts):
        """여러 텍스트 일괄 전처리 (HTML 처리 포함)"""
        results = []
        batch_stats = {'total': len(kr_texts), 'successful': 0, 'failed': 0, 'glossary_hits': 0, 'complex_texts': 0, 'html_texts': 0}
        
        for i, kr_text in enumerate(kr_texts):
            try:
                result = self.preprocess_for_translation(kr_text)
                results.append(result)
                batch_stats['successful'] += 1
                if result['glossary_matches']: batch_stats['glossary_hits'] += 1
                if result['metadata']['complexity'] != 'simple': batch_stats['complex_texts'] += 1
                if result['html_info']['is_html']: batch_stats['html_texts'] += 1
            except Exception as e:
                results.append({'original': kr_text, 'error': str(e), 'processed': kr_text})
                batch_stats['failed'] += 1
        
        return results, batch_stats
    
    # === HTML 처리 관련 메서드들 ===
    
    def _analyze_html_structure(self, result):
        """HTML 구조 분석"""
        text = result['original']
        html_info = result['html_info']
        html_info['is_html'] = is_html(text)
        
        if html_info['is_html']:
            if len(text) > 4000:
                html_info['needs_chunking'] = True
                html_info['chunks'] = split_text_with_html(text, 4500)
            result['metadata']['has_markup'] = True
            if result['metadata']['complexity'] == 'simple': result['metadata']['complexity'] = 'markup'
    
    def _handle_html_chunking(self, result):
        """HTML 청킹 처리"""
        html_info = result['html_info']
        if html_info['needs_chunking'] and html_info['chunks']:
            processed_chunks = [self._clean_text(chunk) for chunk in html_info['chunks']]
            result['processed'] = processed_chunks
            result['warnings'].append(f"텍스트가 {len(html_info['chunks'])}개 청크로 분할되었습니다")
    
    def _restore_html_structure(self, original_kr, translated_text):
        """HTML 구조 복원"""
        try:
            original_tags = re.findall(r'<[^>]+>', original_kr)
            for tag in original_tags:
                if tag not in translated_text:
                    if tag.startswith('<span class="glossary-term">'): continue
                    translated_text += f" {tag}"
            return translated_text
        except Exception as e:
            print(f"HTML 구조 복원 오류: {e}")
            return translated_text
    
    def _handle_glossary_spans(self, original_kr, translated_text):
        """용어집 span 태그 처리"""
        try:
            if 'class="glossary-term"' in original_kr:
                return unwrap_span(translated_text, "glossary-term")
            return translated_text
        except Exception as e:
            print(f"Glossary span 처리 오류: {e}")
            return translated_text
    
    def _clean_text(self, text):
        """기본 텍스트 정제 (HTML 고려)"""
        if not text: return ""
        if is_html(text):
            def clean_text_content(match):
                return re.sub(r'\s+', ' ', match.group(0).strip())
            return re.sub(r'>([^<]+)<', lambda m: f'>{clean_text_content(m)}<', text).strip()
        else:
            cleaned = text.strip()
            cleaned = re.sub(r'\s+', ' ', cleaned)
            return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', cleaned)
    
    def _analyze_text_structure(self, result):
        """텍스트 구조 분석 (HTML 정보 반영)"""
        text = result['processed']
        metadata = result['metadata']
        if metadata['complexity'] != 'markup':
            if len(text) > 200: metadata['complexity'] = 'complex'
            elif len(text) > 100: metadata['complexity'] = 'medium'
        if re.search(r'\[#[^]]+\]|\{[^}]+\}', text):
            metadata['has_markup'] = True
            if metadata['complexity'] == 'simple': metadata['complexity'] = 'markup'
        if re.search(r'\{[^}]*\}|%[sd]|\$\w+', text): metadata['has_variables'] = True
        if result['html_info']['is_html']:
            metadata['has_markup'] = True
            if result['html_info']['needs_chunking']: metadata['complexity'] = 'complex'
    
    def _protect_special_elements(self, result):
        """특수 요소 보호 (HTML 태그 포함)"""
        text, protected_elements = result['processed'], []
        protection_patterns = [
            (r'<[^>]+>', 'html_tag'), (r'\[#[^]]+\]', 'color_tag'), (r'\{[^}]+\}', 'variable'), (r'\\[nt]', 'escape_char'),
            (r'[@]\w+', 'mention'), (r'#\w+', 'hashtag'), (r'\d+%', 'percentage'), (r'\$\d+', 'currency')
        ]
        for pattern, element_type in protection_patterns:
            for match in re.finditer(pattern, text):
                protected_elements.append({
                    'type': element_type, 'content': match.group(), 'position': match.start(), 'end_position': match.end(),
                    'placeholder': f'<PROTECTED_{len(protected_elements)}>'
                })
        result['protected_elements'] = protected_elements
        if protected_elements: self.text_stats['tags_protected'] += len(protected_elements)
    
    def _estimate_translation_difficulty(self, result):
        """번역 난이도 추정 (HTML 복잡도 반영)"""
        metadata = result['metadata']
        difficulty = 1.0
        if metadata['length'] > 200: difficulty += 0.5
        elif metadata['length'] > 100: difficulty += 0.3
        difficulty += {'simple': 0, 'medium': 0.3, 'complex': 0.6, 'markup': 0.8}.get(metadata['complexity'], 0)
        if result['html_info']['is_html']:
            difficulty += 0.4
            if result['html_info']['needs_chunking']: difficulty += 0.3
        if metadata['has_markup']: difficulty += 0.4
        if metadata['has_variables']: difficulty += 0.3
        if len(result['glossary_matches']) > 3: difficulty -= 0.2
        metadata['estimated_difficulty'] = max(1.0, min(3.0, difficulty))
    
    def _validate_preprocessed_text(self, result):
        """전처리된 텍스트 검증 (HTML 유효성 포함)"""
        text, warnings = result['processed'], result['warnings']
        if not text.strip(): warnings.append("텍스트가 비어있습니다")
        if len(text) > 1000: warnings.append("텍스트가 너무 깁니다 (1000자 초과)")
        if result['html_info']['is_html'] and not self._validate_html_tags(text): warnings.append("HTML 태그가 올바르지 않습니다")
        if text.count('[') != text.count(']'): warnings.append("불완전한 대괄호 태그가 감지되었습니다")
        if text.count('{') != text.count('}'): warnings.append("불완전한 중괄호 태그가 감지되었습니다")
    
    def _validate_html_tags(self, text):
        """HTML 태그 유효성 검사"""
        try:
            tag_stack, tag_pattern = [], r'<(/?)([^>\s]+)[^>]*>'
            self_closing_tags = {'br', 'img', 'hr', 'input', 'meta', 'link'}
            for match in re.finditer(tag_pattern, text):
                is_closing, tag_name = bool(match.group(1)), match.group(2).lower()
                if tag_name in self_closing_tags: continue
                if is_closing:
                    if not tag_stack or tag_stack[-1] != tag_name: return False
                    tag_stack.pop()
                else:
                    tag_stack.append(tag_name)
            return len(tag_stack) == 0
        except Exception: return True
    
    def _clean_translated_text(self, text):
        """번역된 텍스트 정제"""
        if not text: return ""
        cleaned = text.strip()
        cleaned = re.sub(r'\s+', ' ', cleaned)
        if cleaned.startswith('"') and cleaned.endswith('"'): cleaned = cleaned[1:-1]
        if cleaned.startswith("'") and cleaned.endswith("'"): cleaned = cleaned[1:-1]
        cleaned = re.sub(r'\s+([.!?:;,])', r'\1', cleaned)
        return re.sub(r'([.!?])\s*$', r'\1', cleaned)
    
    def _check_glossary_consistency(self, original_kr, result):
        """용어집 일관성 검사"""
        if not self.manager.glossary: return
        original_text = re.sub(r'<[^>]*>', '', original_kr) if is_html(original_kr) else original_kr
        translated_text = re.sub(r'<[^>]*>', '', result['processed_translation']) if is_html(result['processed_translation']) else result['processed_translation']
        inconsistencies = []
        for kr_term, translations in self.manager.glossary.items():
            if kr_term in original_text:
                expected_en = translations.get('EN', '')
                if expected_en and expected_en.lower() not in translated_text.lower():
                    inconsistencies.append({'kr_term': kr_term, 'expected_en': expected_en, 'issue': 'glossary_term_missing'})
        if inconsistencies:
            result['warnings'].extend([f"용어집 불일치: {inc['kr_term']} -> {inc['expected_en']}" for inc in inconsistencies])
    
    def _verify_tag_preservation(self, original_kr, result):
        """특수 태그 보존 검증 (HTML 포함)"""
        original_tags = self._extract_tags(original_kr)
        translated_tags = self._extract_tags(result['processed_translation'])
        if len(original_tags) != len(translated_tags):
            result['warnings'].append(f"태그 개수 불일치: 원본 {len(original_tags)}개, 번역 {len(translated_tags)}개")
        for orig_tag in original_tags:
            if orig_tag not in translated_tags:
                result['warnings'].append(f"누락된 태그: {orig_tag}")
    
    def _assess_translation_quality(self, original_kr, translated_text, target_lang):
        """번역 품질 평가"""
        if not translated_text: return 0.0
        quality_score = 1.0
        orig_text_only = re.sub(r'<[^>]*>', '', original_kr) if is_html(original_kr) else original_kr
        trans_text_only = re.sub(r'<[^>]*>', '', translated_text) if is_html(translated_text) else translated_text
        length_ratio = len(trans_text_only) / len(orig_text_only) if orig_text_only else 0
        if target_lang == 'EN' and (length_ratio < 0.5 or length_ratio > 2.0): quality_score -= 0.3
        if self._has_repetitive_patterns(trans_text_only): quality_score -= 0.2
        if not self._is_complete_sentence(trans_text_only): quality_score -= 0.1
        if target_lang == 'EN': quality_score *= self._assess_english_quality(trans_text_only)
        return max(0.0, min(1.0, quality_score))
    
    def _apply_auto_fixes(self, result):
        """자동 수정 적용"""
        text, fixes = result['processed_translation'], []
        if not is_html(text):
            if text and not text[0].isupper():
                text = text[0].upper() + text[1:]
                fixes.append("첫 글자 대문자화")
            if text and text[-1] not in '.!?' and len(text.split()) > 3:
                text += '.'
                fixes.append("마침표 추가")
        original_spaces = text.count('  ')
        text = re.sub(r'\s+', ' ', text)
        if text.count('  ') < original_spaces: fixes.append("이중 공백 제거")
        result['processed_translation'] = text
        result['fixes_applied'] = fixes
    
    def _apply_speaker_context(self, text, speaker_info):
        """화자 컨텍스트 적용"""
        if not speaker_info: return None
        enhancement = {'type': 'speaker_context', 'speaker': speaker_info.get('name', ''), 'suggestions': []}
        if 'tone' in speaker_info:
            tone = speaker_info['tone']
            if tone in ['formal', '정중한']: enhancement['suggestions'].append("정중한 어조로 번역")
            elif tone in ['casual', '친근한']: enhancement['suggestions'].append("캐주얼한 어조로 번역")
        return enhancement if enhancement['suggestions'] else None
    
    def _apply_translation_patterns(self, text, previous_translations):
        """이전 번역 패턴 적용"""
        patterns = [prev['pattern'] for prev in previous_translations if 'pattern' in prev]
        if patterns: return {'type': 'translation_pattern', 'patterns': patterns}
        return None
    
    def _apply_domain_knowledge(self, text, domain):
        """도메인 지식 적용"""
        hints = {'game': ['게임 용어 우선 사용', '플레이어 중심 표현'], 'technical': ['기술 용어 정확성 중시', '명확한 표현'], 'casual': ['친근한 표현 사용', '간결한 문체']}.get(domain, [])
        if hints: return {'type': 'domain_knowledge', 'domain': domain, 'hints': hints}
        return None
    
    def _calculate_glossary_confidence(self, term, text, position):
        """용어집 매칭 신뢰도 계산"""
        is_word_boundary = (position == 0 or not text[position-1].isalnum()) and (position + len(term) >= len(text) or not text[position + len(term)].isalnum())
        confidence = 0.8 if is_word_boundary else 0.6
        if len(term) >= 4: confidence += 0.1
        return min(1.0, confidence)
    
    def _filter_overlapping_matches(self, matches):
        """겹치는 매칭 필터링"""
        if not matches: return matches
        sorted_matches = sorted(matches, key=lambda x: x['position'])
        filtered = []
        for match in sorted_matches:
            if not any(match['position'] < existing['end_position'] and match['end_position'] > existing['position'] for existing in filtered):
                filtered.append(match)
        return filtered
    
    def _extract_tags(self, text):
        """텍스트에서 태그 추출 (HTML 태그 포함)"""
        return [tag for pattern in [r'<[^>]+>', r'\[#[^]]+\]', r'\{[^}]+\}'] for tag in re.findall(pattern, text)]
    
    def _has_repetitive_patterns(self, text):
        """반복 패턴 감지"""
        words = text.split()
        if len(words) < 4: return False
        return any(words[i] == words[i+1] for i in range(len(words)-1))
    
    def _is_complete_sentence(self, text):
        """완전한 문장인지 검사"""
        if not text or len(text.strip()) < 2 or len(text.split()) < 1: return False
        return True
    
    def _assess_english_quality(self, text):
        """영어 번역 품질 평가"""
        quality = 1.0
        if not re.search(r'[a-zA-Z]', text): quality -= 0.5
        for pattern in [r'\ba a\b', r'\bthe the\b', r'\s{2,}']:
            if re.search(pattern, text, re.IGNORECASE): quality -= 0.1
        return max(0.1, quality)

    def get_processing_stats(self):
        """전처리 통계 반환 (HTML 처리 통계 포함)"""
        return self.text_stats.copy()

    def reset_stats(self):
        """통계 초기화"""
        self.text_stats = {'processed_count': 0, 'glossary_applied': 0, 'tags_protected': 0, 'errors_detected': 0, 'html_processed': 0}