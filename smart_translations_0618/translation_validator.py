# 파일명: translation_validator.py

import json

class TranslationValidator:
    def __init__(self, manager):
        """
        번역 품질 검증 및 개선을 담당하는 클래스.
        """
        self.manager = manager
        self.stats = {
            'total_validations': 0,
            'grammar_corrections': 0,
            'failed_corrections': 0
        }

    def correct_grammar_with_llm(self, reassembled_sentence, preprocessed_data, translated_terms_dict):
        """
        LLM을 사용하여 재조립된 문장의 문법을 교정하고 자연스럽게 만듭니다.

        :param reassembled_sentence: 1차 재조립된 문장 (e.g., "Avin의 Adventure")
        :param preprocessed_data: 전처리 결과 데이터 dict
        :param translated_terms_dict: 번역된 용어 dict
        :return: 최종적으로 교정된 문장
        """
        # API 클라이언트가 없거나 문장이 비어있으면 원본 그대로 반환
        if not self.manager.api_client.openai_client or not reassembled_sentence.strip():
            return reassembled_sentence

        self.stats['total_validations'] += 1

        # 1. 프롬프트에 고정시킬 고유명사 목록 생성
        preserved_nouns = []
        if preprocessed_data.get('placeholder_map'):
            for kr_term in preprocessed_data['placeholder_map'].values():
                if kr_term in translated_terms_dict:
                    preserved_nouns.append(translated_terms_dict[kr_term])
        
        # 중복 제거
        preserved_nouns = sorted(list(set(preserved_nouns)))

        # 2. 문법 교정용 프롬프트 구성
        prompt = f"""You are an expert English proofreader for game localization.
Correct the grammar and awkward phrasing of the following sentence to make it sound natural, as if written by a native English speaker.

**IMPORTANT RULES:**
1. You **MUST** preserve the following proper nouns exactly as they are. Do not change their spelling or casing. If they are not in the sentence, do not add them.
   - Preserved Nouns: {json.dumps(preserved_nouns, ensure_ascii=False)}
2. The core meaning of the sentence must not change.
3. Return **ONLY** the corrected sentence, with no additional explanation, preamble, or quotation marks.

**Sentence to Correct:**
{reassembled_sentence}

**Corrected Sentence:**
"""

        try:
            # 3. LLM API 호출
            response = self.manager.api_client.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that corrects grammar while preserving specific proper nouns."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                max_tokens=int(len(reassembled_sentence) * 1.5) + 50, # 원본 길이의 1.5배 + 여유분
                timeout=20
            )

            # 토큰 사용량 로그
            if response.usage:
                print(f"✅ [Grammar Polish] Token Usage - Prompt: {response.usage.prompt_tokens}, Completion: {response.usage.completion_tokens}, Total: {response.usage.total_tokens}")

            corrected_sentence = response.choices[0].message.content.strip()
            
            # LLM이 불필요한 말을 덧붙이는 경우 정리
            if corrected_sentence.startswith('"') and corrected_sentence.endswith('"'):
                corrected_sentence = corrected_sentence[1:-1]

            if corrected_sentence:
                self.stats['grammar_corrections'] += 1
                return corrected_sentence
            else:
                # 비어 있는 응답을 받은 경우
                self.stats['failed_corrections'] += 1
                return reassembled_sentence # 원본 문장 반환

        except Exception as e:
            print(f"❌ 문법 교정 중 API 오류 발생: {e}")
            self.stats['failed_corrections'] += 1
            return reassembled_sentence # 오류 발생 시 원본 문장 반환

    def get_validation_stats(self):
        return self.stats

    def reset_stats(self):
        self.stats = {'total_validations': 0, 'grammar_corrections': 0, 'failed_corrections': 0}