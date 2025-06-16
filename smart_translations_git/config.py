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
DEEPL_API_KEY = get_api_key('DEEPL_API_KEY', required=True)
AZURE_API_KEY = get_api_key('AZURE_API_KEY', required=False)
AZURE_REGION = os.getenv('AZURE_REGION', 'koreacentral')
OPENAI_API_KEY = get_api_key('OPENAI_API_KEY', required=False)

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