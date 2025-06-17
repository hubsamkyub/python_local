import time

class TranslationMetrics:
    """번역 효율성 메트릭 수집 클래스"""
    def __init__(self):
        self.reset_session()
        
    def reset_session(self):
        """세션별 메트릭 초기화"""
        self.session_start = time.time()
        self.api_calls = {
            'deepl_en': {'count': 0, 'chars': 0, 'time': 0, 'cost': 0},
            'deepl_multi': {'count': 0, 'chars': 0, 'time': 0, 'cost': 0},
            'azure': {'count': 0, 'chars': 0, 'time': 0, 'cost': 0},
            'openai_llm': {'count': 0, 'chars': 0, 'time': 0, 'cost': 0}
        }
        self.tm_usage = {'hits': 0, 'misses': 0, 'partial_hits': 0}
        self.translation_results = {'success': 0, 'failures': 0, 'skipped': 0}
        self.duplicate_prevention = {'prevented': 0, 'unique_texts': set()}
        self.processing_times = []
        
    def log_api_call(self, api_type, char_count, response_time, success=True):
        """API 호출 로깅"""
        if api_type in self.api_calls:
            self.api_calls[api_type]['count'] += 1
            self.api_calls[api_type]['chars'] += char_count
            self.api_calls[api_type]['time'] += response_time
            self.api_calls[api_type]['cost'] += self.calculate_cost(api_type, char_count)
            
    def log_tm_usage(self, kr_text, found_langs):
        """TM 사용 현황 로깅"""
        if found_langs:
            if len(found_langs) >= 2:  # 주요 언어 모두 있음
                self.tm_usage['hits'] += 1
            else:
                self.tm_usage['partial_hits'] += 1
        else:
            self.tm_usage['misses'] += 1
            
    def log_duplicate_check(self, kr_text):
        """중복 텍스트 체크"""
        if kr_text in self.duplicate_prevention['unique_texts']:
            self.duplicate_prevention['prevented'] += 1
        else:
            self.duplicate_prevention['unique_texts'].add(kr_text)
            
    def calculate_cost(self, api_type, char_count):
        """API 호출 비용 계산 (예상 단가)"""
        cost_per_char = {
            'deepl_en': 0.00002,      # $20/1M chars
            'deepl_multi': 0.00002,
            'azure': 0.00001,         # $10/1M chars  
            'openai_llm': 0.00003     # $30/1M chars (추정)
        }
        return cost_per_char.get(api_type, 0) * char_count
        
    def get_session_summary(self):
        """세션 요약 정보 반환"""
        total_time = time.time() - self.session_start
        total_api_calls = sum(api['count'] for api in self.api_calls.values())
        total_cost = sum(api['cost'] for api in self.api_calls.values())
        
        return {
            'session_duration': total_time,
            'total_api_calls': total_api_calls,
            'total_cost': total_cost,
            'tm_hit_rate': self.tm_usage['hits'] / max(1, sum(self.tm_usage.values())),
            'api_efficiency': self.duplicate_prevention['prevented'],
            'detailed_metrics': {
                'api_calls': self.api_calls,
                'tm_usage': self.tm_usage,
                'translation_results': self.translation_results,
                'duplicate_prevention': self.duplicate_prevention
            }
        }

