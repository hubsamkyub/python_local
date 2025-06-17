"""
설정 및 상수 정의
"""
import os
from dotenv import load_dotenv

# 환경 변수 로드
load_dotenv()

def get_api_key(key_name, required=True):
    """API 키를 안전하게 가져오는 함수"""
    key = os.getenv(key_name)
    if required and (not key or key.startswith('여기에_')):
        print(f"경고: {key_name}이 제대로 설정되지 않았습니다.")
        return None
    return key

# API 키들
AZURE_API_KEY = get_api_key('AZURE_API_KEY', required=False)
AZURE_REGION = os.getenv('AZURE_REGION', 'koreacentral')
OPENAI_API_KEY = get_api_key('OPENAI_API_KEY', required=False)
ZEMINAI_API_KEY = get_api_key('OPENAI_API_KEY', required=False)

# 언어 코드 매핑
LANG_CODES = {
    "EN": ("en", "EN"),
    "CN": ("zh-Hans", "ZH"),
    "TW": ("zh-Hant", "ZH"),
    "JP": ("ja", "JA"),
    "TH": ("th", "TH"),
    "PT": ("pt", "PT"),
    "ES": ("es", "ES"),
    "DE": ("de", "DE"),
    "FR": ("fr", "FR"),
}

MODEL_CONFIGS = {
    'gpt': {
        'model': 'gpt-4o-mini',  # 기본값: GPT-4o-mini (비용 효율적)
        'max_tokens': 1000,
        'temperature': 0.3,
        # 다른 모델 옵션들:
        # 'model': 'gpt-4o',           # 더 고성능이지만 비용 높음
        # 'model': 'gpt-4-turbo',      # 빠른 GPT-4
        # 'model': 'gpt-3.5-turbo',    # 가장 저렴
    },
    'gemini': {
        'model': 'gemini-2.0-flash',  # 빠르고 저렴한 모델
        'max_tokens': 1000,
        'temperature': 0.3,
        'top_p': 0.9,
        # 'model': 'gemini-2.0-flash',   # 더 고성능 모델
    },
    'azure': {
        'model': 'gpt-4',  # Azure에 배포된 모델명
        'max_tokens': 1000,
        'temperature': 0.3
    }
}