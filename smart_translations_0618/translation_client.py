# 파일명: translation_client.py

import deepl
import requests
import openai
import re
import time
import uuid
import hashlib
import os
from config import AZURE_API_KEY, AZURE_REGION, OPENAI_API_KEY
from utils import TextProtector

class TranslationApiClient:
    def __init__(self, text_protector):
        """
        다양한 번역 API를 호출하는 클라이언트 클래스입니다.
        """
        self.text_protector = text_protector

        # API 클라이언트 초기화
        self.openai_client = openai.OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY and not OPENAI_API_KEY.startswith('여기에') else None
        self.gemini_api_key = os.getenv('GEMINI_API_KEY')
        if not self.gemini_api_key:
            print("⚠️ GEMINI_API_KEY가 설정되지 않았습니다. Gemini 번역을 사용할 수 없습니다.")

        # LLM 최적화를 위한 속성
        self.translation_cache = {}
        # token_usage_tracker는 제거되었습니다.

    # --- Public 메서드 (외부에서 호출) ---

    def translate(self, engine, text, prompt=None, target_lang_code='EN-US', source_lang_code=None, use_protection=True):
        """엔진에 따라 적절한 번역 메서드를 호출하는 메인 메서드"""
        if engine == 'azure':
            return self.translate_azure(text, target_lang_code, source_lang_code, use_protection)
        elif engine == 'llm':
            return self.translate_llm_optimized(text, prompt)
        else:
            raise ValueError(f"알 수 없는 번역 엔진입니다: {engine}")

    def translate_azure(self, text, target_lang, source_lang=None, use_protection=True):
        """Azure Translator API로 번역합니다."""
        if not all([AZURE_API_KEY, AZURE_REGION]) or AZURE_API_KEY.startswith('여기에'):
            raise ConnectionError("Azure API 키 또는 지역이 설정되지 않았습니다.")
        
        if use_protection:
            protected_text, protection_map = self.text_protector.protect_text(text)
        else:
            protected_text, protection_map = text, {}

        endpoint = "https://api.cognitive.microsofttranslator.com"
        path = '/translate'
        constructed_url = endpoint + path
        params = {'api-version': '3.0', 'to': target_lang}
        if source_lang:
            params['from'] = source_lang

        headers = {
            'Ocp-Apim-Subscription-Key': AZURE_API_KEY,
            'Ocp-Apim-Subscription-Region': AZURE_REGION,
            'Content-type': 'application/json',
            'X-ClientTraceId': str(uuid.uuid4())
        }
        body = [{'text': protected_text}]
        
        try:
            response = requests.post(constructed_url, params=params, headers=headers, json=body)
            response.raise_for_status()
            translated_text = response.json()[0]['translations'][0]['text']
            return self.text_protector.restore_text(translated_text, protection_map) if use_protection else translated_text
        except Exception as e:
            print(f"Azure 번역 오류: {e}")
            return None

    def translate_llm_optimized(self, text, prompt, speaker=None, max_retries=3):
        """GPT-4o-mini 최적화된 LLM 번역 (캐싱, 토큰 체크 포함)"""
        if not self.openai_client:
            raise ConnectionError("OpenAI API 키가 설정되지 않았습니다.")
            
        prompt_hash = hashlib.md5(prompt.encode()).hexdigest()[:8]
        cache_key = self._get_cache_key(text, prompt_hash, speaker)
        
        if cached_result := self._get_cached_translation(cache_key):
            return cached_result
        
        estimated_tokens = self._calculate_token_estimate(text + prompt)
        if estimated_tokens > 3500:
            print(f"⚠️ 토큰 수 초과 예상({estimated_tokens}), 긴 텍스트 분할 처리 시도")
            return self._translate_long_text_llm(text, prompt, speaker)
        
        result = self._call_gpt4o_mini_api(text, prompt, max_retries)
        
        if result:
            self._cache_translation(cache_key, result)
        
        return result

    def get_optimization_stats(self):
        """LLM 최적화 통계를 계산하여 반환합니다."""
        tracker = self.token_usage_tracker
        total_tokens = tracker['total_input_tokens'] + tracker['total_output_tokens']
        input_cost = tracker['total_input_tokens'] * 0.00015 / 1000
        output_cost = tracker['total_output_tokens'] * 0.0006 / 1000
        total_cost = input_cost + output_cost

        return {
            'total_requests': tracker['total_requests'],
            'total_tokens': total_tokens,
            'input_tokens': tracker['total_input_tokens'],
            'output_tokens': tracker['total_output_tokens'],
            'cache_hits': tracker['cache_hits'],
            'cache_hit_rate': tracker['cache_hits'] / max(1, tracker['total_requests']),
            'estimated_cost': total_cost
        }
        
    # --- Private 헬퍼 메서드 ---

    def _call_gpt4o_mini_api(self, text, prompt, max_retries):
        """GPT-4o-mini API를 호출하고 결과를 반환합니다."""
        full_prompt = f"{prompt}\n\n{text}"
        for attempt in range(max_retries):
            try:
                response = self.openai_client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": "You are a professional translator. Return only the translation, no explanations."},
                        {"role": "user", "content": full_prompt}
                    ],
                    max_tokens=min(1200, len(text) * 3 + 200),
                    temperature=0.2,
                    timeout=25
                )
                if response.choices:
                    result = response.choices[0].message.content.strip()
                    return self._post_process_llm_result(result, text)
            except Exception as e:
                print(f"GPT-4o-mini API 호출 실패 (시도 {attempt + 1}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
        return None

    def _translate_long_text_llm(self, text, prompt, speaker):
        """LLM으로 긴 텍스트를 분할하여 번역합니다."""
        sentences = text.split('. ')
        if len(sentences) == 1:
            return self._call_gpt4o_mini_api(text[:3000], prompt, 1)

        translated_parts = []
        current_group = ""
        for sentence in sentences:
            if len(current_group) + len(sentence) > 2500:
                result = self.translate_llm_optimized(current_group, prompt, speaker, 1)
                if result: translated_parts.append(result)
                current_group = sentence
            else:
                current_group += ('. ' if current_group else '') + sentence
        
        if current_group:
            result = self.translate_llm_optimized(current_group, prompt, speaker, 1)
            if result: translated_parts.append(result)
            
        return ' '.join(translated_parts) if translated_parts else None

    def _post_process_llm_result(self, result, original_text):
        """LLM 번역 결과에서 불필요한 부분을 정리합니다."""
        # 이 부분은 기존 로직을 그대로 사용하거나 더 정교하게 만들 수 있습니다.
        unwanted_phrases = ["here's the translation:", "translation:", "영어 번역:", "번역:"]
        for phrase in unwanted_phrases:
            if result.lower().startswith(phrase):
                result = result[len(phrase):].strip()
        if (result.startswith('"') and result.endswith('"')):
            result = result[1:-1]
        return result.strip()

    def _calculate_token_estimate(self, text):
        return len(text) // 3

    def _get_cache_key(self, text, prompt_hash, speaker=None):
        cache_string = f"{text}|{prompt_hash}|{speaker or 'default'}"
        return hashlib.md5(cache_string.encode()).hexdigest()

    def _get_cached_translation(self, cache_key):
        if cache_key in self.translation_cache:
            self.token_usage_tracker['cache_hits'] += 1
            return self.translation_cache[cache_key]
        return None

    def _cache_translation(self, cache_key, result):
        if len(self.translation_cache) > 1000:
            oldest_key = next(iter(self.translation_cache))
            del self.translation_cache[oldest_key]
        self.translation_cache[cache_key] = result