import threading
import json
from difflib import SequenceMatcher
from config import LANG_CODES
from text_processor import TextProcessor
import google.generativeai as genai
import openai 
from config import LANG_CODES, MODEL_CONFIGS

class TranslationEngine:
    def __init__(self, manager):
        self.manager = manager
        self.text_processor = TextProcessor(manager)
        self.model_configs = MODEL_CONFIGS
        self._initialize_api_clients()
        self.current_stats = {
            'total': 0, 'completed': 0, 'failed': 0, 'tm_used': 0,
            'api_calls': 0, 'gemini_calls': 0, 'gpt_calls': 0
        }
    
    def analyze_translations(self):
        """번역 분석 (TM 상태 확인 강화) - 전처리 단계"""
        do_en_trans = self.manager.translate_en_var.get()
        do_multi_trans = self.manager.translate_multi_var.get()
        do_cn_tw_trans = self.manager.translate_cn_tw_var.get()

        target_langs = set()
        if do_en_trans: target_langs.add("EN")
        if do_multi_trans: target_langs.update(self.manager.MULTI_LANG_GROUP)
        if do_cn_tw_trans: target_langs.update(["CN", "TW"])

        if not target_langs:
            from tkinter import messagebox
            messagebox.showinfo("정보", "분석할 언어 옵션을 선택해주세요.")
            return

        self.manager.update_status("번역 분석 중...")
        tm_used_count, api_needed_count = 0, 0
        for trans in self.manager.pending_translations:
            kr_text, methods, needs_api = trans["KR"], set(), False
            if kr_text in self.manager.translation_memory:
                tm_entry = self.manager.translation_memory[kr_text]
                if any(trans["translations"].get(lang) != text for lang, text in tm_entry.items() if lang in target_langs and text):
                    methods.add("DB")
                    tm_used_count += 1
            for lang in target_langs:
                if not trans["translations"].get(lang):
                    if lang in ["CN", "TW"] and do_cn_tw_trans: methods.add("DB필요")
                    else: needs_api = True
            if needs_api:
                methods.add("API필요")
                api_needed_count += 1
            trans["method"] = " / ".join(sorted(list(methods))) if methods else "완료"
        self.manager.update_translation_table()
        self.manager.update_status(f"분석 완료. TM사용: {tm_used_count}, API필요: {api_needed_count}")

    def _initialize_api_clients(self):
        try:
            if hasattr(self.manager.api_client, 'gemini_api_key'):
                genai.configure(api_key=self.manager.api_client.gemini_api_key)
                self.gemini_model = genai.GenerativeModel(self.model_configs['gemini']['model'])
            else: self.gemini_model = None
            print(f"번역 엔진 초기화 완료 - GPT: {self.model_configs['gpt']['model']}, Gemini: {self.model_configs['gemini']['model']}")
        except Exception as e:
            print(f"API 클라이언트 초기화 오류: {e}")
            self.gemini_model = None
            
    def execute_translation(self, items_to_translate):
        if not items_to_translate:
            from tkinter import messagebox
            messagebox.showwarning("알림", "번역할 항목이 선택되지 않았습니다.")
            return
        self.current_stats = {'total': len(items_to_translate), 'completed': 0, 'failed': 0, 'tm_used': 0, 'api_calls': 0, 'gemini_calls': 0, 'gpt_calls': 0}
        threading.Thread(target=self._execute_translation_thread_new, args=(items_to_translate,), daemon=True).start()

    def _execute_translation_thread_new(self, items_to_translate):
        """(최종 버전) 용어집 번역을 직접 사용하고, 문법 교정에만 API를 사용하는 최종 워크플로우"""
        try:
            total_items = len(items_to_translate)
            self.manager.update_progress(0, "번역 준비 중: 텍스트 전처리 시작...")

            # 1. 전처리 단계: 모든 항목을 미리 처리
            preprocessed_items = []
            for i, trans_item in enumerate(items_to_translate):
                self.manager.update_progress((i / total_items) * 50, f"전처리 중 ({i+1}/{total_items})...")
                preprocessed_data = self.text_processor.preprocess_for_translation(trans_item['KR'])
                preprocessed_items.append({'original_item': trans_item, 'preprocessed_data': preprocessed_data})

            # 2. 번역된 용어 딕셔너리 생성 (API 호출 없이, 용어집에서 직접!)
            # 이 단계에서 불필요한 API 호출을 제거합니다.
            self.manager.update_progress(50, "용어집에서 직접 번역본 구성 중...")
            translated_terms_dict = {}
            for item_data in preprocessed_items:
                if item_data['preprocessed_data'].get('glossary_matches'):
                    for match in item_data['preprocessed_data']['glossary_matches']:
                        # 용어집에 있는 'en' 번역을 직접 사용
                        translated_terms_dict[match['term']] = match['translations'].get('en')
            
            print(f"✅ 용어집 기반 번역 구성 완료. {len(translated_terms_dict)}개 용어 처리 준비됨.")
            self.manager.update_progress(60, "문장 재조립 및 최종 교정 시작...")

            # 3. 문장 재조립 및 최종 교정
            final_translated_items = []
            for i, item_data in enumerate(preprocessed_items):
                self.manager.update_progress(60 + ((i / total_items) * 40), f"후처리 중 ({i+1}/{total_items})...")
                trans_item, pre_data = item_data['original_item'], item_data['preprocessed_data']
                
                if not self.manager.translate_en_var.get() or trans_item['translations'].get('EN'):
                    self.current_stats['completed'] += 1
                    continue

                # 용어집 매칭이 없는 경우에만 문장 전체 번역 (기존과 동일)
                if not pre_data['template_text']:
                    en_result = self._translate_single_text(trans_item['KR'])
                    trans_item['translations']['EN'] = en_result
                    trans_item['method'] = "API(Full)"
                else:
                    # 용어집 번역으로 1차 재조립
                    reassembled_sentence = self.text_processor.reassemble_from_placeholders(
                        pre_data['template_text'],
                        pre_data['placeholder_map'],
                        translated_terms_dict
                    )
                    
                    # 최종 문법 교정 (API 호출)
                    final_sentence = self.manager.translation_validator.correct_grammar_with_llm(
                        reassembled_sentence,
                        pre_data,
                        translated_terms_dict
                    )
                    trans_item['translations']['EN'] = final_sentence
                    trans_item['method'] = "API(Glossary+Polish)" # 새로운 번역 방법

                self.current_stats['completed'] += 1
                final_translated_items.append(trans_item)

            # 4. DB 저장
            if final_translated_items:
                updated_krs = self.manager.db_manager.update_translation_memory(final_translated_items)
                if updated_krs: self.manager.translation_memory = self.manager.db_manager.get_translation_memory()
                self.manager.root.after(0, self._on_translation_complete, len(updated_krs))
            else:
                self.manager.root.after(0, self._on_translation_complete, 0)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.manager.root.after(0, lambda: self._show_error(f"번역 스레드에서 오류가 발생했습니다:\n{e}"))
            self.manager.update_status("번역 중 오류 발생")
            
    def _translate_single_text(self, text):
        llm_prompt = self.manager.get_llm_prompt()
        return self.manager.api_client.translate('llm', text, prompt=llm_prompt)

    def _on_translation_complete(self, count):
        self.manager.update_translation_table()
        self.manager.update_status(f"번역 완료. {count}개 항목 DB 업데이트됨.")
        self.manager.progress_bar['value'] = 100
        from tkinter import messagebox
        messagebox.showinfo("완료", f"{count}개 항목의 번역 및 저장이 완료되었습니다.")
        
    def _show_error(self, message):
        from tkinter import messagebox
        messagebox.showerror("번역 오류", message)
        self.manager.update_status("번역 중 오류 발생")

    def _preprocess_text(self, trans_item, engine, speaker_mapping=None):
            """번역 전 텍스트 전처리 - TextProcessor에 위임"""
            kr_text = trans_item['KR']
            
            # 컨텍스트 정보 구성
            context = {
                'string_id': trans_item.get('STRING_ID'),
                'engine': engine
            }
            
            # 시나리오 모드에서 화자 정보 추가
            if speaker_mapping and self.manager.scenario_translation_var.get():
                speaker_name = self._get_speaker_for_item(trans_item, speaker_mapping)
                if speaker_name and self.manager.scenario_manager:
                    context['speaker'] = {
                        'name': speaker_name,
                        'profile': self.manager.scenario_manager.speakers.get(speaker_name)
                    }
            
            # TextProcessor를 사용한 전처리
            preprocessed = self.text_processor.preprocess_for_translation(kr_text, context)
            
            # 기존 형식으로 변환 (호환성 유지)
            return {
                'original_text': preprocessed['original'],
                'cleaned_text': preprocessed['processed'],
                'context': context,
                'glossary_matches': preprocessed['glossary_matches'],
                'speaker_info': context.get('speaker'),
                'complexity_level': preprocessed['metadata']['complexity'],
                'warnings': preprocessed['warnings'],
                'metadata': preprocessed['metadata']
            }


    def _translate_to_english(self, preprocessed_data, engine, llm_prompt, use_protection):
        """영어 번역 실행 - 엔진별 분기 처리"""
        kr_text = preprocessed_data['original_text']
        
        try:
            # 시나리오 모드 번역
            if preprocessed_data['speaker_info'] and engine == 'llm':
                return self._translate_with_scenario(preprocessed_data, llm_prompt, use_protection)
            
            # 엔진별 번역 처리
            if engine == 'gemini':
                result = self._translate_with_gemini(kr_text, llm_prompt, use_protection)
                if result:
                    self.current_stats['gemini_calls'] += 1
            elif engine == 'gpt' or engine == 'llm':  # llm을 gpt의 별칭으로 처리
                result = self._translate_with_gpt(kr_text, llm_prompt, use_protection)
                if result:
                    self.current_stats['gpt_calls'] += 1
            else:
                # 기존 방식 (Azure 등)
                result = self.manager.api_client.translate(
                    engine, 
                    kr_text, 
                    prompt=llm_prompt, 
                    target_lang_code='EN-US', 
                    use_protection=use_protection
                )
            
            return result
            
        except Exception as e:
            print(f"EN 번역 오류 ({engine}): {e}")
            return None

    def _translate_to_multiple_languages(self, trans_item, en_text, engine, use_protection):
        """다국어 번역 실행 - 엔진별 처리"""
        langs_to_translate = []
        
        # 번역할 언어 목록 생성
        if self.manager.translate_multi_var.get():
            langs_to_translate.extend([lang for lang in self.manager.MULTI_LANG_GROUP 
                                    if not trans_item['translations'].get(lang)])
        
        if self.manager.translate_cn_tw_var.get():
            if not trans_item['translations'].get('CN'): 
                langs_to_translate.append('CN')
            if not trans_item['translations'].get('TW'): 
                langs_to_translate.append('TW')
        
        # 각 언어로 번역
        for lang in langs_to_translate:
            try:
                target_lang_code = LANG_CODES[lang][1]
                
                # 엔진별 번역 처리
                if engine == 'gemini':
                    result = self._translate_with_gemini_to_language(en_text, lang, use_protection)
                    if result:
                        self.current_stats['gemini_calls'] += 1
                elif engine == 'gpt' or engine == 'llm':
                    result = self._translate_with_gpt_to_language(en_text, lang, use_protection)
                    if result:
                        self.current_stats['gpt_calls'] += 1
                else:
                    # 기존 방식
                    result = self.manager.api_client.translate(
                        engine, 
                        en_text, 
                        target_lang_code=target_lang_code, 
                        source_lang_code='EN', 
                        use_protection=use_protection
                    )
                
                if result:
                    trans_item['translations'][lang] = result
                    self.current_stats['api_calls'] += 1
                    
            except Exception as e:
                print(f"{lang} 번역 오류: {e}")

    def _translate_with_gpt_to_language(self, en_text, target_lang, use_protection):
        """GPT로 영어를 다른 언어로 번역"""
        try:
            lang_names = {
                'CN': 'Simplified Chinese', 'TW': 'Traditional Chinese',
                'TH': 'Thai', 'PT': 'Portuguese', 'ES': 'Spanish',
                'FR': 'French', 'DE': 'German', 'JP': 'Japanese'
            }
            
            target_lang_name = lang_names.get(target_lang, target_lang)
            prompt = f"Translate the following English text to {target_lang_name}:\n\n{en_text}"
            
            model_name = self.model_configs['gpt']['model']
            
            response = openai.ChatCompletion.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": f"You are a professional translator specializing in English to {target_lang_name} translation."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=self.model_configs['gpt']['max_tokens'],
                temperature=self.model_configs['gpt']['temperature']
            )
            
            return response.choices[0].message.content.strip()
            
        except Exception as e:
            print(f"GPT {target_lang} 번역 오류: {e}")
            return None

    def _translate_with_gemini_to_language(self, en_text, target_lang, use_protection):
        """Gemini로 영어를 다른 언어로 번역"""
        try:
            if not self.gemini_model:
                return None
                
            lang_names = {
                'CN': 'Simplified Chinese', 'TW': 'Traditional Chinese',
                'TH': 'Thai', 'PT': 'Portuguese', 'ES': 'Spanish',
                'FR': 'French', 'DE': 'German', 'JP': 'Japanese'
            }
            
            target_lang_name = lang_names.get(target_lang, target_lang)
            prompt = f"Translate the following English text to {target_lang_name}:\n\n{en_text}"
            
            generation_config = genai.types.GenerationConfig(
                temperature=self.model_configs['gemini']['temperature'],
                max_output_tokens=self.model_configs['gemini']['max_tokens']
            )
            
            response = self.gemini_model.generate_content(prompt, generation_config=generation_config)
            
            if response.text:
                return response.text.strip()
            return None
            
        except Exception as e:
            print(f"Gemini {target_lang} 번역 오류: {e}")
            return None
        
        
    def _postprocess_translation(self, trans_item, engine):
        """번역 후 후처리 - TextProcessor 활용"""
        kr_text = trans_item['KR']
        en_text = trans_item['translations'].get('EN', '')
        
        if en_text:
            # TextProcessor로 후처리
            postprocessed = self.text_processor.postprocess_translation(kr_text, en_text, 'EN')
            
            # 후처리된 결과 적용
            trans_item['translations']['EN'] = postprocessed['processed_translation']
            
            # 품질 점수에 따른 상태 설정
            quality_score = postprocessed['quality_score']
            if quality_score >= 0.8:
                trans_item['status'] = "[완료]"
                trans_item['method'] = engine.upper()
            elif quality_score >= 0.6:
                trans_item['status'] = "[검토권장]"
                trans_item['method'] = engine.upper() + "_REVIEW"
            else:
                trans_item['status'] = "[검토필요]"
                trans_item['method'] = engine.upper() + "_LOW_QUALITY"
            
            # 경고사항이 있으면 로그에 기록
            if postprocessed['warnings']:
                print(f"번역 후처리 경고 ({trans_item['STRING_ID']}): {postprocessed['warnings']}")
        else:
            trans_item['status'] = "[실패]"
            trans_item['method'] = engine.upper() + "_FAILED"
            
    def _translate_with_scenario(self, preprocessed_data, base_prompt, use_protection):
        """시나리오 모드 번역"""
        speaker_info = preprocessed_data['speaker_info']
        kr_text = preprocessed_data['original_text']
        
        if not speaker_info or not self.manager.scenario_manager:
            return None
        
        try:
            # 화자별 맞춤 프롬프트 생성
            enhanced_prompt = self.manager.scenario_manager.generate_enhanced_speaker_prompt(
                speaker_info['name'], kr_text, "EN"
            )
            
            # LLM 번역 실행
            result = self.manager.api_client.translate(
                'llm', 
                kr_text, 
                prompt=enhanced_prompt, 
                target_lang_code='EN-US', 
                use_protection=use_protection
            )
            
            # 일관성 검증
            if result and self.manager.scenario_manager:
                consistency_check = self.manager.scenario_manager.validate_translation_consistency(
                    speaker_info['name'], kr_text, result, "EN"
                )
                
                # 일관성이 낮으면 재번역 시도
                if not consistency_check.get("is_consistent", True) and consistency_check.get("confidence", 1.0) < 0.5:
                    print(f"일관성 부족으로 재번역 시도: {kr_text[:30]}...")
                    
                    refined_prompt = enhanced_prompt + f"""

이전 번역이 '{speaker_info['name']}' 캐릭터의 기존 패턴과 일치하지 않습니다.
다음 점을 개선하여 다시 번역하세요:
{chr(10).join(['- ' + suggestion for suggestion in consistency_check.get("suggestions", [])])}

더 일관된 번역:"""
                    
                    refined_result = self.manager.api_client.translate(
                        'llm', 
                        kr_text, 
                        prompt=refined_prompt, 
                        target_lang_code='EN-US', 
                        use_protection=use_protection
                    )
                    if refined_result:
                        result = refined_result
            
            return result
            
        except Exception as e:
            print(f"시나리오 번역 오류: {e}")
            return None

    def force_retranslate_selected(self, selected_string_ids):
        """선택된 항목들 강제 재번역"""
        threading.Thread(target=self._force_retranslate_thread, 
                        args=(selected_string_ids,), daemon=True).start()

    def _force_retranslate_thread(self, string_ids):
        """강제 재번역 스레드"""
        try:
            self.manager.update_status("강제 재번역 시작...")
            items_to_retranslate = [trans for trans in self.manager.pending_translations 
                                  if trans["STRING_ID"] in string_ids]

            if not items_to_retranslate:
                self.manager.update_status("재번역할 항목을 찾을 수 없습니다.")
                return

            selected_engine = self.manager.api_engine_var.get()
            use_protection = self.manager.protect_tags_var.get()
            llm_prompt = self.manager.get_llm_prompt() if selected_engine == 'llm' else None
            success_count = 0

            for i, trans in enumerate(items_to_retranslate):
                kr_text = trans["KR"]
                self.manager.update_status(f"재번역 중 ({i+1}/{len(items_to_retranslate)}): {kr_text[:20]}...")

                # EN 번역만 강제 수행
                en_result = self.manager.api_client.translate(
                    selected_engine, 
                    kr_text, 
                    prompt=llm_prompt, 
                    target_lang_code='EN-US', 
                    use_protection=use_protection
                )
                if en_result:
                    trans["translations"]["EN"] = en_result
                    trans["method"] = f"재번역({selected_engine.upper()})"
                    trans["status"] = "[재번역완료]"
                    success_count += 1
                else:
                    trans["method"] = "재번역실패"
                    trans["status"] = "[실패]"

            if success_count > 0:
                self.manager.db_manager.update_translation_memory(items_to_retranslate)
                self.manager.translation_memory = self.manager.db_manager.get_translation_memory()

            self.manager.root.after(0, self._on_force_retranslate_complete, success_count, len(items_to_retranslate))

        except Exception as e:
            self.manager.update_status(f"재번역 오류: {e}")

    def check_multilang_prerequisites(self):
        """다국어 번역 선행 조건 체크"""
        do_en_trans = self.manager.translate_en_var.get()
        do_multi_trans = self.manager.translate_multi_var.get()
        do_cn_tw_trans = self.manager.translate_cn_tw_var.get()
        
        # EN이 체크되지 않았지만 다국어가 체크된 경우
        if (do_multi_trans or do_cn_tw_trans) and not do_en_trans:
            return False, "EN 번역이 필요합니다"
        
        # 선택된 항목 중 EN이 비어있는 항목 체크
        empty_en_count = 0
        total_selected = 0
        
        for trans in self.manager.pending_translations:
            item_id = self.manager.find_item_id_by_string_id(trans["STRING_ID"])
            if self.manager.check_states.get(item_id, True):
                total_selected += 1
                if not trans["translations"].get("EN"):
                    empty_en_count += 1
        
        if empty_en_count > 0 and (do_multi_trans or do_cn_tw_trans):
            return False, f"{empty_en_count}개 항목에 EN 번역이 없습니다"
        
        return True, "조건 충족"


    # === 헬퍼 메서드들 ===       
    def _get_speaker_for_item(self, trans_item, speaker_mapping):
        """번역 항목의 화자 정보 가져오기"""
        if not speaker_mapping:
            return None
            
        string_id = trans_item["STRING_ID"]
        for speaker, string_ids in speaker_mapping.items():
            if string_id in string_ids:
                return speaker
        return None
    
    def _on_force_retranslate_complete(self, success_count, total_count):
        """강제 재번역 완료 후 UI 업데이트"""
        if hasattr(self.manager, 'filter_vars') and '재번역완료' in self.manager.filter_vars:
            self.manager.filter_vars['재번역완료'].set(True)
        self.manager.update_translation_table()
        self.manager.update_status(f"재번역 완료: {success_count}/{total_count}개 성공")

    def find_similar_translation(self, kr_text, threshold=0.9):
        """유사 번역 찾기 - 전처리 후 비교"""
        # 텍스트 전처리
        preprocessed = self.text_processor.preprocess_for_translation(kr_text)
        cleaned_text = preprocessed['processed']
        
        best_match = None
        best_similarity = 0
        
        for saved_kr, translations in self.manager.translation_memory.items():
            # 저장된 텍스트도 동일하게 전처리
            saved_preprocessed = self.text_processor.preprocess_for_translation(saved_kr)
            saved_cleaned = saved_preprocessed['processed']
            
            # 전처리된 텍스트로 유사도 계산
            similarity = SequenceMatcher(None, cleaned_text, saved_cleaned).ratio()
            
            if similarity >= threshold and similarity > best_similarity:
                best_similarity = similarity
                best_match = {
                    "kr": saved_kr,
                    "translations": translations,
                    "similarity": similarity,
                    "preprocessed_similarity": True
                }
                
        return best_match

    def apply_glossary(self, kr_text):
        """용어집 적용 - TextProcessor 활용"""
        preprocessed = self.text_processor.preprocess_for_translation(kr_text)
        
        if preprocessed['glossary_matches']:
            # 첫 번째 매칭된 용어의 번역 반환
            first_match = preprocessed['glossary_matches'][0]
            translations = first_match['translations']
            return {"EN": f"[{translations.get('EN', '')}...]"}
        
        return None
    
    def _translate_with_gpt(self, text, prompt, use_protection):
        """GPT를 이용한 번역"""
        try:
            # 프롬프트 구성
            if prompt:
                full_prompt = f"{prompt}\n\n{text}"
            else:
                full_prompt = f"Translate the following Korean text to natural English:\n\n{text}"
            
            # GPT 모델 설정 가져오기
            model_name = self.model_configs['gpt']['model']
            max_tokens = self.model_configs['gpt']['max_tokens']
            temperature = self.model_configs['gpt']['temperature']
            
            # OpenAI API 호출
            response = openai.ChatCompletion.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": "You are a professional translator specializing in Korean to English translation."},
                    {"role": "user", "content": full_prompt}
                ],
                max_tokens=max_tokens,
                temperature=temperature
            )
            
            result = response.choices[0].message.content.strip()
            
            # 태그 보호 처리
            if use_protection:
                result = self._restore_protected_tags(text, result)
            
            return result
            
        except Exception as e:
            print(f"GPT 번역 오류: {e}")
            return None

    def _translate_with_gemini(self, text, prompt, use_protection):
        """Gemini를 이용한 번역"""
        try:
            if not self.gemini_model:
                print("Gemini 모델이 초기화되지 않았습니다.")
                return None
            
            # 프롬프트 구성
            if prompt:
                full_prompt = f"{prompt}\n\n{text}"
            else:
                full_prompt = f"""You are a professional Korean-English translator. 
                
    Translate the following Korean text to natural, fluent English:

    {text}

    Translation:"""
            
            # Gemini 설정
            generation_config = genai.types.GenerationConfig(
                temperature=self.model_configs['gemini']['temperature'],
                max_output_tokens=self.model_configs['gemini']['max_tokens'],
                top_p=self.model_configs['gemini']['top_p']
            )
            
            # Gemini API 호출
            response = self.gemini_model.generate_content(
                full_prompt,
                generation_config=generation_config
            )
            
            if response.text:
                result = response.text.strip()
                
                # 태그 보호 처리
                if use_protection:
                    result = self._restore_protected_tags(text, result)
                    
                return result
            else:
                print("Gemini 응답에 텍스트가 없습니다.")
                return None
            
        except Exception as e:
            print(f"Gemini 번역 오류: {e}")
            return None

    def _restore_protected_tags(self, original_text, translated_text):
        """보호된 태그 복원 (간단 버전)"""
        # 기본적인 태그 복원 로직
        # 실제로는 text_processor의 postprocess_translation을 사용하는 것이 좋음
        import re
        
        # 원본에서 특수 태그 추출
        tags = re.findall(r'(\{[^}]+\}|\[#[^]]+\]|<[^>]+>)', original_text)
        
        # 번역문에 태그가 누락되었으면 끝에 추가 (임시 방식)
        for tag in tags:
            if tag not in translated_text:
                translated_text += f" {tag}"
        
        return translated_text
    
    def get_translation_stats(self):
        """번역 통계 반환"""
        return {
            **self.current_stats,
            'success_rate': (self.current_stats['completed'] / max(1, self.current_stats['total'])) * 100,
            'gpt_usage': self.current_stats['gpt_calls'],
            'gemini_usage': self.current_stats['gemini_calls']
        }
        
