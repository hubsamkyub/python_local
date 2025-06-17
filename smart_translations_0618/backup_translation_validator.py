import re
import json
import time
from datetime import datetime
from difflib import SequenceMatcher
from collections import defaultdict, Counter

class TranslationValidator:
    def __init__(self, manager):
        self.manager = manager
        self.validation_stats = {
            'total_validations': 0,
            'passed_validations': 0,
            'failed_validations': 0,
            'auto_fixes_applied': 0,
            'retranslation_suggested': 0
        }
        
        # 품질 기준 설정
        self.quality_thresholds = {
            'excellent': 0.9,
            'good': 0.7,
            'acceptable': 0.5,
            'poor': 0.3
        }
        
        # 검증 규칙 가중치
        self.validation_weights = {
            'length_ratio': 0.2,
            'tag_preservation': 0.3,
            'glossary_consistency': 0.2,
            'completeness': 0.15,
            'fluency': 0.15
        }
    
    def validate_translation(self, original_kr, translated_text, target_lang='EN', context=None):
        """번역 품질 종합 검증 - 메인 진입점"""
        validation_start_time = time.time()
        
        result = {
            'overall_score': 0.0,
            'quality_level': 'unknown',
            'detailed_scores': {},
            'issues': [],
            'suggestions': [],
            'auto_fixes': [],
            'metadata': {
                'validation_time': 0.0,
                'timestamp': datetime.now().isoformat(),
                'target_language': target_lang,
                'original_length': len(original_kr),
                'translated_length': len(translated_text)
            }
        }
        
        try:
            # 1. 기본 유효성 검사
            if not self._basic_validation_check(original_kr, translated_text, result):
                return result
            
            # 2. 세부 품질 검증
            self._validate_length_ratio(original_kr, translated_text, target_lang, result)
            self._validate_tag_preservation(original_kr, translated_text, result)
            self._validate_glossary_consistency(original_kr, translated_text, result)
            self._validate_translation_completeness(original_kr, translated_text, result)
            self._validate_fluency_and_naturalness(translated_text, target_lang, result)
            
            # 3. 컨텍스트 기반 검증 (시나리오 모드 등)
            if context:
                self._validate_context_consistency(original_kr, translated_text, context, result)
            
            # 4. 전체 점수 계산
            self._calculate_overall_score(result)
            
            # 5. 품질 레벨 결정
            self._determine_quality_level(result)
            
            # 6. 자동 수정 적용
            self._apply_auto_fixes(original_kr, translated_text, result)
            
            # 7. 개선 제안 생성
            self._generate_improvement_suggestions(result)
            
            self.validation_stats['total_validations'] += 1
            if result['overall_score'] >= self.quality_thresholds['acceptable']:
                self.validation_stats['passed_validations'] += 1
            else:
                self.validation_stats['failed_validations'] += 1
            
        except Exception as e:
            result['issues'].append(f"검증 중 오류 발생: {e}")
            result['overall_score'] = 0.0
        
        finally:
            result['metadata']['validation_time'] = time.time() - validation_start_time
        
        return result
    
    def batch_validate_translations(self, translation_pairs, target_lang='EN'):
        """여러 번역들을 일괄 검증"""
        batch_results = []
        batch_stats = {
            'total': len(translation_pairs),
            'excellent': 0,
            'good': 0,
            'acceptable': 0,
            'poor': 0,
            'average_score': 0.0,
            'common_issues': Counter()
        }
        
        total_score = 0.0
        
        for i, (original, translated) in enumerate(translation_pairs):
            result = self.validate_translation(original, translated, target_lang)
            batch_results.append(result)
            
            # 통계 업데이트
            total_score += result['overall_score']
            quality_level = result['quality_level']
            if quality_level in batch_stats:
                batch_stats[quality_level] += 1
            
            # 공통 이슈 수집
            for issue in result['issues']:
                issue_type = self._categorize_issue(issue)
                batch_stats['common_issues'][issue_type] += 1
        
        if batch_stats['total'] > 0:
            batch_stats['average_score'] = total_score / batch_stats['total']
        
        return batch_results, batch_stats
    
    def suggest_retranslation(self, validation_result, original_kr):
        """재번역 필요성 및 개선 방향 제안"""
        retranslation_advice = {
            'should_retranslate': False,
            'priority': 'low',
            'specific_issues': [],
            'improved_prompt_suggestions': [],
            'alternative_approaches': []
        }
        
        overall_score = validation_result['overall_score']
        issues = validation_result['issues']
        
        # 재번역 필요성 판단
        if overall_score < self.quality_thresholds['acceptable']:
            retranslation_advice['should_retranslate'] = True
            retranslation_advice['priority'] = 'high'
        elif overall_score < self.quality_thresholds['good']:
            retranslation_advice['should_retranslate'] = True
            retranslation_advice['priority'] = 'medium'
        
        # 구체적인 이슈 분석
        for issue in issues:
            if '태그' in issue:
                retranslation_advice['specific_issues'].append('특수 태그 처리 개선 필요')
                retranslation_advice['improved_prompt_suggestions'].append('태그 보존을 더 강조하는 프롬프트 사용')
            elif '용어집' in issue:
                retranslation_advice['specific_issues'].append('용어집 일관성 개선 필요')
                retranslation_advice['improved_prompt_suggestions'].append('해당 용어의 정확한 번역을 프롬프트에 명시')
            elif '길이' in issue:
                retranslation_advice['specific_issues'].append('번역 길이 조정 필요')
                retranslation_advice['improved_prompt_suggestions'].append('간결함 또는 상세함을 명시하는 프롬프트 사용')
        
        # 대안적 접근법 제안
        if len(retranslation_advice['specific_issues']) > 2:
            retranslation_advice['alternative_approaches'].append('문장 분할 후 개별 번역')
        
        if overall_score < self.quality_thresholds['poor']:
            retranslation_advice['alternative_approaches'].append('다른 번역 엔진 시도')
            retranslation_advice['alternative_approaches'].append('인간 번역자 검토 요청')
        
        return retranslation_advice
    
    def analyze_translation_patterns(self, translation_history):
        """번역 패턴 분석으로 품질 개선점 찾기"""
        pattern_analysis = {
            'frequent_issues': Counter(),
            'quality_trends': [],
            'engine_performance': defaultdict(list),
            'recommendations': []
        }
        
        for entry in translation_history:
            if 'validation_result' in entry:
                validation = entry['validation_result']
                engine = entry.get('engine', 'unknown')
                
                # 엔진별 성능 추적
                pattern_analysis['engine_performance'][engine].append(validation['overall_score'])
                
                # 빈번한 이슈 수집
                for issue in validation['issues']:
                    issue_type = self._categorize_issue(issue)
                    pattern_analysis['frequent_issues'][issue_type] += 1
        
        # 개선 권장사항 생성
        most_common_issues = pattern_analysis['frequent_issues'].most_common(3)
        for issue_type, count in most_common_issues:
            if count > len(translation_history) * 0.3:  # 30% 이상 발생하는 이슈
                recommendation = self._generate_pattern_recommendation(issue_type)
                pattern_analysis['recommendations'].append(recommendation)
        
        return pattern_analysis
    
    def create_quality_report(self, validation_results, title="번역 품질 분석 보고서"):
        """상세한 품질 분석 보고서 생성"""
        report = {
            'title': title,
            'summary': {},
            'detailed_analysis': {},
            'recommendations': [],
            'charts_data': {},
            'generated_at': datetime.now().isoformat()
        }
        
        if not validation_results:
            report['summary']['message'] = "분석할 번역 결과가 없습니다."
            return report
        
        # 요약 통계
        scores = [r['overall_score'] for r in validation_results]
        report['summary'] = {
            'total_translations': len(validation_results),
            'average_score': sum(scores) / len(scores),
            'highest_score': max(scores),
            'lowest_score': min(scores),
            'quality_distribution': self._calculate_quality_distribution(validation_results)
        }
        
        # 상세 분석
        report['detailed_analysis'] = {
            'common_issues': self._analyze_common_issues(validation_results),
            'score_breakdown': self._analyze_score_breakdown(validation_results),
            'improvement_areas': self._identify_improvement_areas(validation_results)
        }
        
        # 차트 데이터
        report['charts_data'] = {
            'score_distribution': [r['overall_score'] for r in validation_results],
            'quality_levels': [r['quality_level'] for r in validation_results],
            'validation_times': [r['metadata']['validation_time'] for r in validation_results]
        }
        
        # 권장사항
        report['recommendations'] = self._generate_quality_recommendations(validation_results)
        
        return report
    
    # === 세부 검증 메서드들 ===
    
    def _basic_validation_check(self, original, translated, result):
        """기본 유효성 검사"""
        if not original or not original.strip():
            result['issues'].append("원본 텍스트가 비어있습니다")
            return False
        
        if not translated or not translated.strip():
            result['issues'].append("번역 텍스트가 비어있습니다")
            result['overall_score'] = 0.0
            return False
        
        return True
    
    def _validate_length_ratio(self, original, translated, target_lang, result):
        """길이 비율 검증"""
        original_len = len(original.strip())
        translated_len = len(translated.strip())
        
        if original_len == 0:
            ratio = 0
        else:
            ratio = translated_len / original_len
        
        # 언어별 적정 비율 범위
        optimal_ranges = {
            'EN': (0.6, 1.8),  # 한영 번역
            'CN': (0.4, 1.2),  # 한중 번역
            'TW': (0.4, 1.2),  # 한번체중 번역
            'JP': (0.8, 1.5)   # 한일 번역
        }
        
        min_ratio, max_ratio = optimal_ranges.get(target_lang, (0.5, 2.0))
        
        if ratio < min_ratio:
            score = 0.3
            result['issues'].append(f"번역문이 너무 짧습니다 (비율: {ratio:.2f})")
        elif ratio > max_ratio:
            score = 0.4
            result['issues'].append(f"번역문이 너무 깁니다 (비율: {ratio:.2f})")
        else:
            # 비율이 적정 범위 내에 있으면 높은 점수
            score = 1.0 - abs(ratio - 1.0) * 0.3
        
        result['detailed_scores']['length_ratio'] = max(0.0, min(1.0, score))
    
    def _validate_tag_preservation(self, original, translated, result):
        """특수 태그 보존 검증"""
        original_tags = self._extract_all_tags(original)
        translated_tags = self._extract_all_tags(translated)
        
        if not original_tags and not translated_tags:
            score = 1.0  # 태그가 없으면 완벽
        elif len(original_tags) == 0:
            score = 0.8 if len(translated_tags) == 0 else 0.6  # 원본에 태그 없는데 번역에 있으면 약간 감점
        else:
            preserved_count = 0
            for tag in original_tags:
                if tag in translated_tags:
                    preserved_count += 1
                else:
                    result['issues'].append(f"누락된 태그: {tag}")
            
            # 추가된 태그도 체크
            extra_tags = [tag for tag in translated_tags if tag not in original_tags]
            for tag in extra_tags:
                result['issues'].append(f"추가된 태그: {tag}")
            
            if len(original_tags) > 0:
                score = preserved_count / len(original_tags)
                if extra_tags:
                    score *= 0.8  # 추가 태그가 있으면 감점
            else:
                score = 1.0
        
        result['detailed_scores']['tag_preservation'] = score
    
    def _validate_glossary_consistency(self, original, translated, result):
        """용어집 일관성 검증"""
        if not self.manager.glossary:
            result['detailed_scores']['glossary_consistency'] = 1.0
            return
        
        total_terms = 0
        correct_translations = 0
        
        for kr_term, translations in self.manager.glossary.items():
            if kr_term in original:
                total_terms += 1
                expected_translation = translations.get('EN', '')
                
                if expected_translation and expected_translation.lower() in translated.lower():
                    correct_translations += 1
                else:
                    result['issues'].append(f"용어집 불일치: '{kr_term}' → 예상 '{expected_translation}'")
        
        if total_terms == 0:
            score = 1.0  # 용어집 해당 없음
        else:
            score = correct_translations / total_terms
        
        result['detailed_scores']['glossary_consistency'] = score
    
    def _validate_translation_completeness(self, original, translated, result):
        """번역 완성도 검증"""
        score = 1.0
        
        # 번역이 중단된 것 같은 패턴 감지
        incomplete_patterns = [
            r'\.\.\.+$',           # 끝에 점점점
            r'\s+$',               # 끝에 공백만
            r'[^\w\s\.]$',         # 이상한 문자로 끝남
            r'^["\'].*[^"\']$',    # 따옴표 짝 안맞음
        ]
        
        for pattern in incomplete_patterns:
            if re.search(pattern, translated):
                score -= 0.2
                result['issues'].append("번역이 불완전해 보입니다")
                break
        
        # 원문의 중요한 정보가 누락되었는지 간단 체크
        original_words = set(re.findall(r'\w+', original.lower()))
        translated_words = set(re.findall(r'\w+', translated.lower()))
        
        # 한글이 번역문에 남아있으면 감점
        korean_pattern = r'[가-힣]+'
        if re.search(korean_pattern, translated):
            score -= 0.3
            result['issues'].append("번역문에 한글이 남아있습니다")
        
        result['detailed_scores']['completeness'] = max(0.0, score)
    
    def _validate_fluency_and_naturalness(self, translated, target_lang, result):
        """유창성과 자연스러움 검증"""
        score = 1.0
        
        if target_lang == 'EN':
            score = self._validate_english_fluency(translated, result)
        elif target_lang in ['CN', 'TW']:
            score = self._validate_chinese_fluency(translated, result)
        
        result['detailed_scores']['fluency'] = score
    
    def _validate_english_fluency(self, text, result):
        """영어 유창성 검증"""
        score = 1.0
        
        # 기본 영어 패턴 검사
        if not re.search(r'[a-zA-Z]', text):
            score -= 0.5
            result['issues'].append("영어 문자가 없습니다")
        
        # 일반적인 오류 패턴 검사
        error_patterns = [
            (r'\ba a\b', "관사 중복 (a a)"),
            (r'\bthe the\b', "관사 중복 (the the)"),
            (r'\s{2,}', "연속된 공백"),
            (r'\b[A-Z]{2,}\b', "과도한 대문자 (약어 제외)"),
            (r'[.!?]{2,}', "연속된 구두점")
        ]
        
        for pattern, description in error_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                score -= 0.1
                result['issues'].append(f"영어 오류: {description}")
        
        # 문장 구조 간단 체크
        if len(text.split()) < 2:
            score -= 0.2
            result['issues'].append("문장이 너무 짧습니다")
        
        return max(0.0, score)
    
    def _validate_chinese_fluency(self, text, result):
        """중국어 유창성 검증"""
        score = 1.0
        
        # 중국어 문자 확인
        if not re.search(r'[\u4e00-\u9fff]', text):
            score -= 0.5
            result['issues'].append("중국어 문자가 없습니다")
        
        # 간체/번체 일관성 검사는 추후 추가 가능
        
        return max(0.0, score)
    
    def _validate_context_consistency(self, original, translated, context, result):
        """컨텍스트 일관성 검증 (시나리오 모드 등)"""
        context_score = 1.0
        
        # 시나리오 모드에서 화자 일관성 검사
        if 'speaker' in context and self.manager.scenario_manager:
            speaker_name = context['speaker'].get('name')
            if speaker_name:
                consistency_check = self.manager.scenario_manager.validate_translation_consistency(
                    speaker_name, original, translated, 'EN'
                )
                
                if not consistency_check.get('is_consistent', True):
                    context_score = consistency_check.get('confidence', 0.5)
                    result['issues'].append(f"화자 '{speaker_name}'의 기존 번역 패턴과 일치하지 않습니다")
                    
                    # 구체적인 제안사항 추가
                    suggestions = consistency_check.get('suggestions', [])
                    result['suggestions'].extend(suggestions)
        
        result['detailed_scores']['context_consistency'] = context_score
    
    # === 분석 및 계산 메서드들 ===
    
    def _calculate_overall_score(self, result):
        """전체 점수 계산"""
        total_score = 0.0
        used_weights = 0.0
        
        for metric, weight in self.validation_weights.items():
            if metric in result['detailed_scores']:
                total_score += result['detailed_scores'][metric] * weight
                used_weights += weight
        
        if used_weights > 0:
            result['overall_score'] = total_score / used_weights
        else:
            result['overall_score'] = 0.0
    
    def _determine_quality_level(self, result):
        """품질 레벨 결정"""
        score = result['overall_score']
        
        if score >= self.quality_thresholds['excellent']:
            result['quality_level'] = 'excellent'
        elif score >= self.quality_thresholds['good']:
            result['quality_level'] = 'good'
        elif score >= self.quality_thresholds['acceptable']:
            result['quality_level'] = 'acceptable'
        else:
            result['quality_level'] = 'poor'
    
    def _apply_auto_fixes(self, original, translated, result):
        """자동 수정 적용"""
        fixed_text = translated
        fixes_applied = []
        
        # 1. 기본 정제
        if fixed_text != fixed_text.strip():
            fixed_text = fixed_text.strip()
            fixes_applied.append("앞뒤 공백 제거")
        
        # 2. 연속 공백 제거
        if re.search(r'\s{2,}', fixed_text):
            fixed_text = re.sub(r'\s+', ' ', fixed_text)
            fixes_applied.append("연속 공백 정리")
        
        # 3. 대소문자 수정 (영어)
        if fixed_text and not fixed_text[0].isupper() and re.search(r'[a-zA-Z]', fixed_text):
            fixed_text = fixed_text[0].upper() + fixed_text[1:]
            fixes_applied.append("첫 글자 대문자화")
        
        result['auto_fixes'] = fixes_applied
        result['fixed_translation'] = fixed_text
        
        if fixes_applied:
            self.validation_stats['auto_fixes_applied'] += 1
    
    def _generate_improvement_suggestions(self, result):
        """개선 제안 생성"""
        suggestions = []
        
        # 점수 기반 제안
        if result['overall_score'] < self.quality_thresholds['acceptable']:
            suggestions.append("전체적인 번역 품질이 낮습니다. 재번역을 고려하세요.")
            self.validation_stats['retranslation_suggested'] += 1
        
        # 세부 점수 기반 제안
        for metric, score in result['detailed_scores'].items():
            if score < 0.7:
                if metric == 'length_ratio':
                    suggestions.append("번역 길이를 조정하세요. 더 간결하거나 상세하게 번역할 필요가 있습니다.")
                elif metric == 'tag_preservation':
                    suggestions.append("특수 태그 보존에 주의하세요. 원본의 모든 태그를 유지해야 합니다.")
                elif metric == 'glossary_consistency':
                    suggestions.append("용어집 일관성을 확인하세요. 정의된 용어의 번역을 사용하세요.")
                elif metric == 'fluency':
                    suggestions.append("번역의 자연스러움을 개선하세요. 문법과 표현을 다시 확인하세요.")
        
        result['suggestions'].extend(suggestions)
    
    # === 유틸리티 메서드들 ===
    
    def _extract_all_tags(self, text):
        """모든 종류의 태그 추출"""
        patterns = [
            r'<[^>]+>',      # HTML 태그
            r'\[#[^]]+\]',   # 색상 태그
            r'\{[^}]+\}',    # 변수
            r'[@#$]\w+'      # 특수 기호 태그
        ]
        
        tags = []
        for pattern in patterns:
            tags.extend(re.findall(pattern, text))
        
        return tags
    
    def _categorize_issue(self, issue):
        """이슈를 카테고리로 분류"""
        if '태그' in issue:
            return 'tag_issues'
        elif '용어집' in issue:
            return 'glossary_issues'
        elif '길이' in issue:
            return 'length_issues'
        elif '영어' in issue:
            return 'language_issues'
        elif '불완전' in issue:
            return 'completeness_issues'
        else:
            return 'other_issues'
    
    def _generate_pattern_recommendation(self, issue_type):
        """패턴 기반 권장사항 생성"""
        recommendations = {
            'tag_issues': "특수 태그 보호 설정을 강화하고, 번역 후 태그 검증을 추가하세요.",
            'glossary_issues': "용어집 적용 로직을 개선하고, 번역 전 용어 매칭을 강화하세요.",
            'length_issues': "번역 길이 가이드라인을 설정하고, 프롬프트에 길이 지침을 추가하세요.",
            'language_issues': "대상 언어별 문법 검사 규칙을 추가하세요.",
            'completeness_issues': "번역 완성도 검사 로직을 강화하세요."
        }
        
        return recommendations.get(issue_type, "해당 이슈 유형에 대한 개선 방안을 검토하세요.")
    
    def _calculate_quality_distribution(self, validation_results):
        """품질 분포 계산"""
        distribution = {'excellent': 0, 'good': 0, 'acceptable': 0, 'poor': 0}
        
        for result in validation_results:
            quality_level = result['quality_level']
            if quality_level in distribution:
                distribution[quality_level] += 1
        
        return distribution
    
    def _analyze_common_issues(self, validation_results):
        """공통 이슈 분석"""
        issue_counter = Counter()
        
        for result in validation_results:
            for issue in result['issues']:
                issue_type = self._categorize_issue(issue)
                issue_counter[issue_type] += 1
        
        return dict(issue_counter.most_common())
    
    def _analyze_score_breakdown(self, validation_results):
        """점수 세부 분석"""
        breakdown = defaultdict(list)
        
        for result in validation_results:
            for metric, score in result['detailed_scores'].items():
                breakdown[metric].append(score)
        
        # 평균 계산
        avg_breakdown = {}
        for metric, scores in breakdown.items():
            avg_breakdown[metric] = sum(scores) / len(scores) if scores else 0.0
        
        return avg_breakdown
    
    def _identify_improvement_areas(self, validation_results):
        """개선 영역 식별"""
        improvement_areas = []
        score_breakdown = self._analyze_score_breakdown(validation_results)
        
        for metric, avg_score in score_breakdown.items():
            if avg_score < 0.7:
                improvement_areas.append({
                    'area': metric,
                    'current_score': avg_score,
                    'improvement_needed': 0.8 - avg_score
                })
        
        return sorted(improvement_areas, key=lambda x: x['improvement_needed'], reverse=True)
    
    def _generate_quality_recommendations(self, validation_results):
        """품질 개선 권장사항 생성"""
        recommendations = []
        improvement_areas = self._identify_improvement_areas(validation_results)
        
        for area in improvement_areas[:3]:  # 상위 3개 개선 영역
            metric = area['area']
            recommendation = self._generate_pattern_recommendation(f"{metric}_issues")
            recommendations.append(f"{metric}: {recommendation}")
        
        return recommendations

    def get_validation_stats(self):
        """검증 통계 반환"""
        return self.validation_stats.copy()

    def reset_stats(self):
        """통계 초기화"""
        self.validation_stats = {
            'total_validations': 0,
            'passed_validations': 0,
            'failed_validations': 0,
            'auto_fixes_applied': 0,
            'retranslation_suggested': 0
        }