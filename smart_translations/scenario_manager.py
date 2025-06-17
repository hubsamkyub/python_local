"""
시나리오 번역 관리 시스템
"""
import json
import pandas as pd
from collections import defaultdict
import re
import sqlite3  # 새로 추가
from datetime import datetime  # 새로 추가
from tkinter import messagebox  # 새로 추가 (만약 없다면)

class SpeakerProfile:
    """화자 프로필 클래스"""
    def __init__(self, name, gender="중성", tone="보통", style="일반적", examples=None):
        self.name = name
        self.gender = gender  # 남성, 여성, 중성
        self.tone = tone      # 정중, 보통, 친근, 거친 등
        self.style = style    # 말투 설명
        self.examples = examples or []  # 번역 예시들
        self.reference_count = 0
        
    def to_dict(self):
        return {
            'name': self.name,
            'gender': self.gender,
            'tone': self.tone,
            'style': self.style,
            'examples': self.examples,
            'reference_count': self.reference_count
        }
    
    @classmethod
    def from_dict(cls, data):
        speaker = cls(
            data['name'], 
            data.get('gender', '중성'),
            data.get('tone', '보통'),
            data.get('style', '일반적'),
            data.get('examples', [])
        )
        speaker.reference_count = data.get('reference_count', 0)
        return speaker

class ScenarioTranslationManager:
    """시나리오 번역 관리자"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self.speakers = {}  # 화자명: SpeakerProfile
        self.reference_data = []  # 레퍼런스 번역 데이터
        self.speaker_patterns = {}  # 화자별 번역 패턴
        self.load_speakers()
        self.init_reference_tables()
    

    def load_speakers(self):
        """저장된 화자 정보 로드 (디버깅 강화)"""
        try:
            import sqlite3
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 화자 테이블 생성 (없는 경우)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS speakers (
                    name TEXT PRIMARY KEY,
                    gender TEXT,
                    tone TEXT,
                    style TEXT,
                    examples TEXT,
                    reference_count INTEGER DEFAULT 0
                )
            """)
            
            # 기존 화자들 로드
            cursor.execute("SELECT * FROM speakers")
            rows = cursor.fetchall()
            
            # 화자 딕셔너리 초기화
            self.speakers = {}
            
            for row in rows:
                name, gender, tone, style, examples_json, ref_count = row
                try:
                    examples = json.loads(examples_json) if examples_json else []
                except:
                    examples = []
                
                speaker = SpeakerProfile(name, gender, tone, style, examples)
                speaker.reference_count = ref_count or 0
                self.speakers[name] = speaker
            
            conn.close()
            
            print(f"✅ 화자 정보 로드 완료: {len(self.speakers)}명")
            if self.speakers:
                print(f"   로드된 화자: {list(self.speakers.keys())}")
            
        except Exception as e:
            print(f"❌ 화자 정보 로드 오류: {e}")
            import traceback
            traceback.print_exc()
            self.speakers = {}
            

    def save_speaker(self, speaker):
        """화자 정보 저장"""
        try:
            import sqlite3
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT OR REPLACE INTO speakers 
                (name, gender, tone, style, examples, reference_count)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                speaker.name,
                speaker.gender,
                speaker.tone,
                speaker.style,
                json.dumps(speaker.examples, ensure_ascii=False),
                speaker.reference_count
            ))
            
            conn.commit()
            conn.close()
            
            self.speakers[speaker.name] = speaker
            
        except Exception as e:
            print(f"화자 정보 저장 오류: {e}")

    # scenario_manager.py에서 데이터 검증 부분 강화
    def is_valid_text(self, text):
        """텍스트가 유효한지 검사"""
        if pd.isna(text):
            return False
        
        text_str = str(text).strip().lower()
        
        # 무효한 값들
        invalid_values = ['', 'nan', 'none', 'null', '#n/a', '#value!', '#ref!']
        
        return text_str not in invalid_values


    def analyze_reference_data(self, file_path, target_language="EN", skiprows=0):
        """레퍼런스 데이터 분석 (상세 디버깅 포함)"""
        debug_log = []
        
        try:
            debug_log.append("=== 📊 레퍼런스 데이터 분석 시작 ===")
            debug_log.append(f"📁 파일: {file_path}")
            debug_log.append(f"🎯 대상 언어: {target_language}")
            debug_log.append(f"⏭️ skiprows: {skiprows}")
            
            # 1단계: 데이터 읽기
            debug_log.append("\n🔍 1단계: 데이터 읽기")
            df = pd.read_excel(file_path, skiprows=skiprows)
            debug_log.append(f"   📏 데이터프레임 크기: {df.shape}")
            debug_log.append(f"   📋 컬럼명: {list(df.columns)}")
            
            # 2단계: 필수 컬럼 확인
            debug_log.append("\n✅ 2단계: 필수 컬럼 확인")
            required_cols = ['KR', '#화자']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                error_msg = f"필수 컬럼이 없습니다: {missing_cols}"
                debug_log.append(f"   ❌ {error_msg}")
                raise ValueError(error_msg)
            
            debug_log.append(f"   ✅ 필수 컬럼 확인 완료: {required_cols}")
            
            # 3단계: 대상 언어 컬럼 확인
            debug_log.append("\n🌍 3단계: 대상 언어 컬럼 확인")
            if target_language not in df.columns:
                available_langs = [col for col in df.columns if col in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']]
                error_msg = f"{target_language} 컬럼이 없습니다. 사용 가능한 언어: {available_langs}"
                debug_log.append(f"   ❌ {error_msg}")
                raise ValueError(error_msg)
            
            debug_log.append(f"   ✅ 대상 언어 컬럼 확인 완료: {target_language}")
            
            # 4단계: 데이터 품질 분석
            debug_log.append("\n🔬 4단계: 데이터 품질 분석")
            total_rows = len(df)
            debug_log.append(f"   📊 전체 행 수: {total_rows}")
            
            # KR 컬럼 분석
            kr_valid = df['KR'].notna() & (df['KR'].astype(str).str.strip() != '') & (df['KR'].astype(str).str.lower() != 'nan')
            kr_valid_count = kr_valid.sum()
            debug_log.append(f"   📝 유효한 KR 데이터: {kr_valid_count}/{total_rows} ({kr_valid_count/total_rows*100:.1f}%)")
            
            # 화자 컬럼 분석
            speaker_valid = df['#화자'].notna() & (df['#화자'].astype(str).str.strip() != '') & (df['#화자'].astype(str).str.lower() != 'nan')
            speaker_valid_count = speaker_valid.sum()
            debug_log.append(f"   👤 유효한 화자 데이터: {speaker_valid_count}/{total_rows} ({speaker_valid_count/total_rows*100:.1f}%)")
            
            # 대상 언어 컬럼 분석
            lang_valid = df[target_language].notna() & (df[target_language].astype(str).str.strip() != '') & (df[target_language].astype(str).str.lower() != 'nan')
            lang_valid_count = lang_valid.sum()
            debug_log.append(f"   🌍 유효한 {target_language} 데이터: {lang_valid_count}/{total_rows} ({lang_valid_count/total_rows*100:.1f}%)")
            
            # 모든 조건을 만족하는 행
            all_valid = kr_valid & speaker_valid & lang_valid
            all_valid_count = all_valid.sum()
            debug_log.append(f"   ✨ 모든 조건 만족하는 행: {all_valid_count}/{total_rows} ({all_valid_count/total_rows*100:.1f}%)")
            
            if all_valid_count == 0:
                error_msg = "모든 필수 조건을 만족하는 데이터가 없습니다"
                debug_log.append(f"   ❌ {error_msg}")
                # 샘플 데이터 표시
                debug_log.append("\n📋 샘플 데이터 (첫 5행):")
                for i in range(min(5, len(df))):
                    row = df.iloc[i]
                    debug_log.append(f"   행{i+1}: KR='{row.get('KR', 'N/A')}', 화자='{row.get('#화자', 'N/A')}', {target_language}='{row.get(target_language, 'N/A')}'")
                
                self.save_debug_log(debug_log)
                raise ValueError(error_msg)
            
            # 5단계: 화자별 데이터 수집
            debug_log.append("\n👥 5단계: 화자별 데이터 수집")
            speaker_data = defaultdict(lambda: defaultdict(list))
            processed_rows = 0
            valid_rows = 0
            
            unique_speakers = set()
            
            for idx, row in df.iterrows():
                processed_rows += 1
                
                # 데이터 검증 및 정리
                kr_text = self.clean_text(row.get('KR', ''))
                speaker = self.clean_text(row.get('#화자', ''))
                
                if not kr_text or not speaker:
                    continue
                
                valid_rows += 1
                unique_speakers.add(speaker)
                
                # 각 언어별로 번역 데이터 수집
                languages_collected = []
                for lang in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']:
                    if lang in df.columns:
                        translated_text = self.clean_text(row.get(lang, ''))
                        if translated_text:
                            speaker_data[speaker][lang].append({
                                'kr': kr_text,
                                'translated': translated_text
                            })
                            languages_collected.append(lang)
            
            debug_log.append(f"   📊 처리된 행: {processed_rows}")
            debug_log.append(f"   ✅ 유효한 행: {valid_rows}")
            debug_log.append(f"   👤 발견된 화자 수: {len(unique_speakers)}")
            debug_log.append(f"   👤 화자 목록: {list(unique_speakers)}")
            
            # 화자별 상세 정보
            debug_log.append("\n📈 화자별 수집 데이터:")
            for speaker, lang_data in speaker_data.items():
                total_translations = sum(len(translations) for translations in lang_data.values())
                debug_log.append(f"   👤 {speaker}: {total_translations}개 번역문")
                for lang, translations in lang_data.items():
                    if translations:
                        debug_log.append(f"      🌍 {lang}: {len(translations)}개")
            
            if not speaker_data:
                error_msg = "유효한 화자 데이터가 수집되지 않았습니다"
                debug_log.append(f"   ❌ {error_msg}")
                self.save_debug_log(debug_log)
                raise ValueError(error_msg)
            
            # 6단계: 화자별 분석 실행
            debug_log.append("\n🔍 6단계: 화자별 패턴 분석")
            analysis_result = {}
            
            for speaker, lang_data in speaker_data.items():
                debug_log.append(f"\n   🎭 화자 '{speaker}' 분석 중...")
                
                speaker_analysis = {
                    'total_sentences': 0,
                    'languages': {},
                    'main_patterns': {}
                }
                
                # 언어별 패턴 분석
                for lang, translations in lang_data.items():
                    if not translations:
                        continue
                        
                    try:
                        debug_log.append(f"      🌍 {lang} 언어 패턴 분석 중... ({len(translations)}개 문장)")
                        patterns = self.analyze_speaker_patterns_for_language(translations, lang)
                        speaker_analysis['languages'][lang] = {
                            'count': len(translations),
                            'patterns': patterns,
                            'examples': translations[:3]
                        }
                        speaker_analysis['total_sentences'] += len(translations)
                        debug_log.append(f"      ✅ {lang} 언어 분석 완료")
                        
                    except Exception as e:
                        debug_log.append(f"      ❌ {lang} 언어 분석 실패: {e}")
                        import traceback
                        debug_log.append(f"      📜 트레이스백: {traceback.format_exc()}")
                
                # 주요 패턴 추출
                if target_language in speaker_analysis['languages']:
                    speaker_analysis['main_patterns'] = speaker_analysis['languages'][target_language]['patterns']
                    debug_log.append(f"      ✅ 주요 패턴 추출 완료 ({target_language} 기준)")
                
                analysis_result[speaker] = speaker_analysis
                
                # 화자 프로필 업데이트
                try:
                    self.update_or_create_speaker_profile(speaker, speaker_analysis, target_language)
                    debug_log.append(f"      ✅ 화자 프로필 저장 완료")
                except Exception as e:
                    debug_log.append(f"      ❌ 화자 프로필 저장 실패: {e}")
            
            # 7단계: 최종 결과
            debug_log.append("\n🎉 7단계: 분석 완료!")
            debug_log.append(f"   ✅ 성공적으로 분석된 화자: {len(analysis_result)}명")
            debug_log.append(f"   📊 총 처리된 문장: {sum(data['total_sentences'] for data in analysis_result.values())}개")
            
            # 디버그 로그 저장
            self.save_debug_log(debug_log)
            
            return analysis_result
            
        except Exception as e:
            debug_log.append(f"\n💥 오류 발생: {e}")
            debug_log.append(f"📜 상세 오류:")
            import traceback
            debug_log.append(traceback.format_exc())
            
            # 오류 시에도 디버그 로그 저장
            self.save_debug_log(debug_log)
            
            print("\n".join(debug_log))  # 콘솔에도 출력
            return {}

    def clean_text(self, text):
        """텍스트 정리 함수"""
        if pd.isna(text):
            return ""
        
        text_str = str(text).strip().lower()
        
        # 무효한 값들
        invalid_values = ['', 'nan', 'none', 'null', '#n/a', '#value!', '#ref!']
        
        if text_str in invalid_values:
            return ""
        
        return str(text).strip()

    def save_debug_log(self, debug_log):
        """디버그 로그를 파일로 저장"""
        try:
            import os
            from datetime import datetime
            
            log_dir = "debug_logs"
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(log_dir, f"analysis_debug_{timestamp}.txt")
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("\n".join(debug_log))
            
            print(f"디버그 로그 저장됨: {log_file}")
            
        except Exception as e:
            print(f"디버그 로그 저장 실패: {e}")

    def analyze_speaker_patterns_for_language(self, translations, language):
        """특정 언어에 대한 화자 번역 패턴 분석"""
        patterns = {
            'avg_length': 0,
            'sentence_types': {'declarative': 0, 'interrogative': 0, 'exclamatory': 0},
            'tone_indicators': {},
            'vocabulary_level': 'normal',
            'formality_level': 'normal'
        }
        
        if not translations:
            return patterns
        
        translated_texts = [t['translated'] for t in translations]
        
        # 평균 길이 (단어 수)
        patterns['avg_length'] = sum(len(text.split()) for text in translated_texts) / len(translated_texts)
        
        # 문장 유형 분석
        for text in translated_texts:
            if text.endswith('?'):
                patterns['sentence_types']['interrogative'] += 1
            elif text.endswith('!'):
                patterns['sentence_types']['exclamatory'] += 1
            else:
                patterns['sentence_types']['declarative'] += 1
        
        # 언어별 톤 분석
        if language == 'EN':
            patterns['tone_indicators'] = self.analyze_english_tone(translated_texts)
        elif language in ['CN', 'TW']:
            patterns['tone_indicators'] = self.analyze_chinese_tone(translated_texts)
        elif language == 'JP':
            patterns['tone_indicators'] = self.analyze_japanese_tone(translated_texts)
        # 다른 언어들도 필요에 따라 추가
        
        return patterns


    def analyze_english_tone(self, texts):
        """영어 텍스트의 톤 분석"""
        all_text = ' '.join(texts).lower()
        
        tone_indicators = {
            'polite': ['please', 'thank you', 'excuse me', 'i apologize', 'would you', 'could you', 'may i'],
            'casual': ['hey', 'yeah', 'gonna', 'wanna', 'ok', 'cool', 'awesome', "i'm", "you're", "can't"],
            'formal': ['indeed', 'therefore', 'furthermore', 'consequently', 'regarding', 'respectfully'],
            'rough': ['damn', 'hell', 'shut up', 'get out', 'what the', 'screw', 'crap'],
            'cute': ['aww', 'yay', 'wow', 'amazing', 'adorable', 'sweetie', 'cutie']
        }
        
        scores = {}
        for tone, indicators in tone_indicators.items():
            scores[tone] = sum(all_text.count(indicator) for indicator in indicators)
        
        return scores

    def analyze_chinese_tone(self, texts):
        """중국어 텍스트의 톤 분석 (간단한 예시)"""
        all_text = ''.join(texts)
        
        tone_indicators = {
            'polite': ['请', '谢谢', '不好意思', '劳烦', '麻烦'],
            'casual': ['嗯', '哦', '呢', '啊', '哈'],
            'formal': ['因此', '所以', '然而', '不过', '另外'],
            'cute': ['呀', '哇', '好可爱', '太棒了']
        }
        
        scores = {}
        for tone, indicators in tone_indicators.items():
            scores[tone] = sum(all_text.count(indicator) for indicator in indicators)
        
        return scores

    def analyze_japanese_tone(self, texts):
        """일본어 텍스트의 톤 분석 (간단한 예시)"""
        all_text = ''.join(texts)
        
        tone_indicators = {
            'polite': ['です', 'ます', 'ございます', 'でしょう', 'していただき'],
            'casual': ['だよ', 'だね', 'じゃん', 'ちゃう', '～や'],
            'cute': ['だわ', '～なの', '～よ', '～ね', 'にゃ', '～です♪'],
            'rough': ['だろ', 'てめぇ', '～やがる', 'ちくしょう']
        }
        
        scores = {}
        for tone, indicators in tone_indicators.items():
            scores[tone] = sum(all_text.count(indicator) for indicator in indicators)
        
        return scores

    def update_or_create_speaker_profile(self, speaker_name, analysis_data, target_language):
        """분석 데이터를 바탕으로 화자 프로필 업데이트 또는 생성"""
        if speaker_name in self.speakers:
            # 기존 화자 업데이트
            speaker = self.speakers[speaker_name]
            speaker.reference_count = analysis_data['total_sentences']
            
            # 새로운 예시 추가 (중복 제거)
            if target_language in analysis_data['languages']:
                new_examples = analysis_data['languages'][target_language]['examples']
                for example in new_examples:
                    if example not in speaker.examples:
                        speaker.examples.append(example)
                # 예시는 최대 5개까지만 유지
                speaker.examples = speaker.examples[:5]
        else:
            # 새 화자 생성
            inferred_profile = self.infer_speaker_profile_from_analysis(speaker_name, analysis_data, target_language)
            self.speakers[speaker_name] = inferred_profile
        
        # DB에 저장
        self.save_speaker(self.speakers[speaker_name])
        
        
  
    def analyze_speaker_patterns(self, translations):
        """화자의 번역 패턴 분석"""
        patterns = {
            'avg_length': 0,
            'common_words': [],
            'tone_indicators': [],
            'sentence_types': {'declarative': 0, 'interrogative': 0, 'exclamatory': 0}
        }
        
        if not translations:
            return patterns
        
        en_texts = [t['en'] for t in translations]
        
        # 평균 길이
        patterns['avg_length'] = sum(len(text.split()) for text in en_texts) / len(en_texts)
        
        # 문장 유형 분석
        for text in en_texts:
            if text.endswith('?'):
                patterns['sentence_types']['interrogative'] += 1
            elif text.endswith('!'):
                patterns['sentence_types']['exclamatory'] += 1
            else:
                patterns['sentence_types']['declarative'] += 1
        
        # 말투 특성 추론
        all_text = ' '.join(en_texts).lower()
        
        # 정중한 말투 지표
        polite_indicators = ['please', 'thank you', 'excuse me', 'i apologize', 'would you', 'could you']
        # 친근한 말투 지표  
        casual_indicators = ['hey', 'yeah', 'gonna', 'wanna', 'ok', 'cool', 'awesome']
        # 거친 말투 지표
        rough_indicators = ['damn', 'hell', 'shut up', 'get out', 'what the']
        
        polite_count = sum(all_text.count(indicator) for indicator in polite_indicators)
        casual_count = sum(all_text.count(indicator) for indicator in casual_indicators)
        rough_count = sum(all_text.count(indicator) for indicator in rough_indicators)
        
        patterns['tone_indicators'] = {
            'polite': polite_count,
            'casual': casual_count,
            'rough': rough_count
        }
        
        return patterns
    
    def infer_speaker_profile(self, speaker_name, translations):
        """번역 데이터로부터 화자 특성 자동 추론"""
        patterns = self.analyze_speaker_patterns(translations)
        
        # 성별 추론 (이름 기반, 단순화)
        gender = "중성"
        if any(name in speaker_name.lower() for name in ['king', 'prince', 'duke', 'sir', 'mr']):
            gender = "남성"
        elif any(name in speaker_name.lower() for name in ['queen', 'princess', 'duchess', 'lady', 'ms', 'mrs']):
            gender = "여성"
        
        # 말투 추론
        tone_scores = patterns['tone_indicators']
        if tone_scores['polite'] > tone_scores['casual'] and tone_scores['polite'] > tone_scores['rough']:
            tone = "정중"
        elif tone_scores['rough'] > tone_scores['casual']:
            tone = "거친"
        elif tone_scores['casual'] > 0:
            tone = "친근"
        else:
            tone = "보통"
        
        # 스타일 설명 생성
        style_parts = []
        
        if patterns['avg_length'] > 15:
            style_parts.append("상세한 설명을 선호")
        elif patterns['avg_length'] < 8:
            style_parts.append("간결한 표현 선호")
            
        exclamatory_ratio = patterns['sentence_types']['exclamatory'] / len(translations)
        if exclamatory_ratio > 0.3:
            style_parts.append("감정 표현이 풍부")
            
        style = ", ".join(style_parts) if style_parts else "일반적인 말투"
        
        speaker = SpeakerProfile(
            name=speaker_name,
            gender=gender,
            tone=tone,
            style=style,
            examples=translations[:3]
        )
        speaker.reference_count = len(translations)
        
        return speaker
    
    
    
    def infer_speaker_profile_from_analysis(self, speaker_name, analysis_data, target_language):
        """분석 데이터를 바탕으로 화자 프로필 자동 추론 (새로운 데이터 구조용)"""
        
        # 기본값 설정
        gender = "중성"
        tone = "보통"
        style_parts = []
        examples = []
        
        # 성별 추론 (이름 기반)
        name_lower = speaker_name.lower()
        male_keywords = ['king', 'prince', 'duke', 'sir', 'mr', 'lord', 'knight', '왕', '왕자', '공작']
        female_keywords = ['queen', 'princess', 'duchess', 'lady', 'ms', 'mrs', 'miss', '여왕', '공주', '공작부인']
        
        for keyword in male_keywords:
            if keyword in name_lower:
                gender = "남성"
                break
        else:
            for keyword in female_keywords:
                if keyword in name_lower:
                    gender = "여성"
                    break
        
        # 대상 언어의 패턴 데이터 추출
        if target_language in analysis_data.get('languages', {}):
            lang_data = analysis_data['languages'][target_language]
            patterns = lang_data.get('patterns', {})
            examples = lang_data.get('examples', [])[:3]  # 최대 3개 예시만
            
            # 평균 길이 기반 스타일 추론
            avg_length = patterns.get('avg_length', 0)
            if avg_length > 15:
                style_parts.append("상세한 설명을 선호")
            elif avg_length < 8:
                style_parts.append("간결한 표현 선호")
            
            # 문장 유형 기반 특성 추론
            sentence_types = patterns.get('sentence_types', {})
            total_sentences = sum(sentence_types.values())
            if total_sentences > 0:
                exclamatory_ratio = sentence_types.get('exclamatory', 0) / total_sentences
                interrogative_ratio = sentence_types.get('interrogative', 0) / total_sentences
                
                if exclamatory_ratio > 0.3:
                    style_parts.append("감정 표현이 풍부")
                if interrogative_ratio > 0.4:
                    style_parts.append("질문을 자주 사용")
            
            # 톤 지표 기반 말투 추론
            tone_indicators = patterns.get('tone_indicators', {})
            if isinstance(tone_indicators, dict):
                max_tone = max(tone_indicators.items(), key=lambda x: x[1], default=('normal', 0))
                tone_map = {
                    'polite': '정중',
                    'casual': '친근',
                    'rough': '거친',
                    'cute': '귀여운',
                    'formal': '격식'
                }
                if max_tone[1] > 0:  # 0보다 큰 값이 있으면
                    tone = tone_map.get(max_tone[0], '보통')
        
        # 스타일 설명 생성
        if not style_parts:
            style_parts.append("일반적인 말투")
        
        # 총 문장 수 정보 추가
        total_sentences = analysis_data.get('total_sentences', 0)
        if total_sentences > 50:
            style_parts.append("풍부한 대사량")
        elif total_sentences < 5:
            style_parts.append("제한적인 대사량")
        
        style = ", ".join(style_parts)
        
        # SpeakerProfile 객체 생성
        speaker = SpeakerProfile(
            name=speaker_name,
            gender=gender,
            tone=tone,
            style=style,
            examples=examples
        )
        speaker.reference_count = total_sentences
        
        return speaker
    

    def generate_speaker_prompt(self, speaker_name):
        """화자별 맞춤 프롬프트 생성 (데이터 구조 호환성 개선)"""
        if speaker_name not in self.speakers:
            return self.get_default_prompt()
        
        speaker = self.speakers[speaker_name]
        
        prompt = f"""다음은 '{speaker.name}' 캐릭터의 대사를 번역하는 작업입니다.

    캐릭터 정보:
    - 이름: {speaker.name}
    - 성별: {speaker.gender}
    - 말투: {speaker.tone}
    - 특징: {speaker.style}

    번역 예시:"""
        
        # 번역 예시 추가 (데이터 구조 호환성 확인)
        for i, example in enumerate(speaker.examples[:3], 1):
            if isinstance(example, dict):
                # 새로운 구조: {'kr': ..., 'translated': ...}
                kr_text = example.get('kr', '')
                en_text = example.get('translated', '') or example.get('en', '')
            else:
                # 기존 구조나 기타 형태 처리
                kr_text = str(example) if example else ''
                en_text = ''
            
            if kr_text and en_text:
                prompt += f"\n{i}. 한국어: {kr_text}\n   영어: {en_text}"
        
        prompt += f"""

    위 예시를 참고하여 '{speaker.name}' 캐릭터의 말투와 특징을 유지하면서 다음 한국어 대사를 영어로 번역해주세요:

    """
        return prompt


    
    def get_default_prompt(self):
        """기본 프롬프트"""
        return """다음 한국어 대사를 자연스러운 영어로 번역해주세요:

"""

    def analyze_reference_data_with_mapping(self, file_path, target_language, skiprows, column_mapping):
        """사용자 정의 컬럼 매핑으로 레퍼런스 데이터 분석"""
        try:
            print(f"=== 사용자 정의 매핑 분석 시작 ===")
            print(f"파일: {file_path}")
            print(f"skiprows: {skiprows}")
            print(f"컬럼 매핑: {column_mapping}")
            
            # 데이터 읽기
            df = pd.read_excel(file_path, skiprows=skiprows)
            print(f"데이터프레임 크기: {df.shape}")
            
            # 매핑된 컬럼명으로 변환
            rename_mapping = {v: k for k, v in column_mapping.items() if v in df.columns}
            df_renamed = df.rename(columns=rename_mapping)
            
            print(f"컬럼 이름 변경: {rename_mapping}")
            print(f"변경 후 컬럼: {list(df_renamed.columns)}")
            
            # 필수 컬럼 확인
            if 'KR' not in df_renamed.columns or '#화자' not in df_renamed.columns:
                raise ValueError("KR 또는 #화자 컬럼 매핑이 잘못되었습니다.")
            
            # 기존 분석 로직 사용 (컬럼명이 정규화된 데이터프레임 사용)
            return self.analyze_dataframe(df_renamed, target_language)
            
        except Exception as e:
            print(f"사용자 정의 매핑 분석 오류: {e}")
            import traceback
            traceback.print_exc()
            return {}


    def analyze_dataframe(self, df, target_language):
        """정규화된 데이터프레임 분석 (analyze_reference_data의 핵심 로직)"""
        debug_log = []
        
        try:
            debug_log.append("=== 📊 정규화된 데이터프레임 분석 시작 ===")
            debug_log.append(f"🎯 대상 언어: {target_language}")
            debug_log.append(f"📏 데이터프레임 크기: {df.shape}")
            debug_log.append(f"📋 컬럼명: {list(df.columns)}")
            
            # 1단계: 필수 컬럼 확인
            debug_log.append("\n✅ 1단계: 필수 컬럼 확인")
            required_cols = ['KR', '#화자']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                error_msg = f"필수 컬럼이 없습니다: {missing_cols}"
                debug_log.append(f"   ❌ {error_msg}")
                raise ValueError(error_msg)
            
            debug_log.append(f"   ✅ 필수 컬럼 확인 완료: {required_cols}")
            
            # 2단계: 대상 언어 컬럼 확인
            debug_log.append("\n🌍 2단계: 대상 언어 컬럼 확인")
            if target_language not in df.columns:
                available_langs = [col for col in df.columns if col in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']]
                error_msg = f"{target_language} 컬럼이 없습니다. 사용 가능한 언어: {available_langs}"
                debug_log.append(f"   ❌ {error_msg}")
                raise ValueError(error_msg)
            
            debug_log.append(f"   ✅ 대상 언어 컬럼 확인 완료: {target_language}")
            
            # 3단계: 데이터 품질 분석
            debug_log.append("\n🔬 3단계: 데이터 품질 분석")
            total_rows = len(df)
            debug_log.append(f"   📊 전체 행 수: {total_rows}")
            
            # KR 컬럼 분석
            kr_valid = df['KR'].notna() & (df['KR'].astype(str).str.strip() != '') & (df['KR'].astype(str).str.lower() != 'nan')
            kr_valid_count = kr_valid.sum()
            debug_log.append(f"   📝 유효한 KR 데이터: {kr_valid_count}/{total_rows} ({kr_valid_count/total_rows*100:.1f}%)")
            
            # 화자 컬럼 분석
            speaker_valid = df['#화자'].notna() & (df['#화자'].astype(str).str.strip() != '') & (df['#화자'].astype(str).str.lower() != 'nan')
            speaker_valid_count = speaker_valid.sum()
            debug_log.append(f"   👤 유효한 화자 데이터: {speaker_valid_count}/{total_rows} ({speaker_valid_count/total_rows*100:.1f}%)")
            
            # 대상 언어 컬럼 분석
            lang_valid = df[target_language].notna() & (df[target_language].astype(str).str.strip() != '') & (df[target_language].astype(str).str.lower() != 'nan')
            lang_valid_count = lang_valid.sum()
            debug_log.append(f"   🌍 유효한 {target_language} 데이터: {lang_valid_count}/{total_rows} ({lang_valid_count/total_rows*100:.1f}%)")
            
            # 모든 조건을 만족하는 행
            all_valid = kr_valid & speaker_valid & lang_valid
            all_valid_count = all_valid.sum()
            debug_log.append(f"   ✨ 모든 조건 만족하는 행: {all_valid_count}/{total_rows} ({all_valid_count/total_rows*100:.1f}%)")
            
            if all_valid_count == 0:
                error_msg = "모든 필수 조건을 만족하는 데이터가 없습니다"
                debug_log.append(f"   ❌ {error_msg}")
                # 샘플 데이터 표시
                debug_log.append("\n📋 샘플 데이터 (첫 5행):")
                for i in range(min(5, len(df))):
                    row = df.iloc[i]
                    debug_log.append(f"   행{i+1}: KR='{row.get('KR', 'N/A')}', 화자='{row.get('#화자', 'N/A')}', {target_language}='{row.get(target_language, 'N/A')}'")
                
                self.save_debug_log(debug_log)
                raise ValueError(error_msg)
            
            # 4단계: 화자별 데이터 수집
            debug_log.append("\n👥 4단계: 화자별 데이터 수집")
            speaker_data = defaultdict(lambda: defaultdict(list))
            processed_rows = 0
            valid_rows = 0
            
            unique_speakers = set()
            
            for idx, row in df.iterrows():
                processed_rows += 1
                
                # 데이터 검증 및 정리
                kr_text = self.clean_text(row.get('KR', ''))
                speaker = self.clean_text(row.get('#화자', ''))
                
                if not kr_text or not speaker:
                    continue
                
                valid_rows += 1
                unique_speakers.add(speaker)
                
                # 각 언어별로 번역 데이터 수집
                for lang in ['EN', 'CN', 'TW', 'TH', 'PT', 'ES', 'DE', 'FR', 'JP']:
                    if lang in df.columns:
                        translated_text = self.clean_text(row.get(lang, ''))
                        if translated_text:
                            speaker_data[speaker][lang].append({
                                'kr': kr_text,
                                'translated': translated_text
                            })
            
            debug_log.append(f"   📊 처리된 행: {processed_rows}")
            debug_log.append(f"   ✅ 유효한 행: {valid_rows}")
            debug_log.append(f"   👤 발견된 화자 수: {len(unique_speakers)}")
            debug_log.append(f"   👤 화자 목록: {list(unique_speakers)}")
            
            # 화자별 상세 정보
            debug_log.append("\n📈 화자별 수집 데이터:")
            for speaker, lang_data in speaker_data.items():
                total_translations = sum(len(translations) for translations in lang_data.values())
                debug_log.append(f"   👤 {speaker}: {total_translations}개 번역문")
                for lang, translations in lang_data.items():
                    if translations:
                        debug_log.append(f"      🌍 {lang}: {len(translations)}개")
            
            if not speaker_data:
                error_msg = "유효한 화자 데이터가 수집되지 않았습니다"
                debug_log.append(f"   ❌ {error_msg}")
                self.save_debug_log(debug_log)
                raise ValueError(error_msg)
            
            # 5단계: 화자별 분석 실행
            debug_log.append("\n🔍 5단계: 화자별 패턴 분석")
            analysis_result = {}
            
            for speaker, lang_data in speaker_data.items():
                debug_log.append(f"\n   🎭 화자 '{speaker}' 분석 중...")
                
                speaker_analysis = {
                    'total_sentences': 0,
                    'languages': {},
                    'main_patterns': {}
                }
                
                # 언어별 패턴 분석
                for lang, translations in lang_data.items():
                    if not translations:
                        continue
                        
                    try:
                        debug_log.append(f"      🌍 {lang} 언어 패턴 분석 중... ({len(translations)}개 문장)")
                        patterns = self.analyze_speaker_patterns_for_language(translations, lang)
                        speaker_analysis['languages'][lang] = {
                            'count': len(translations),
                            'patterns': patterns,
                            'examples': translations[:3]
                        }
                        speaker_analysis['total_sentences'] += len(translations)
                        debug_log.append(f"      ✅ {lang} 언어 분석 완료")
                        
                    except Exception as e:
                        debug_log.append(f"      ❌ {lang} 언어 분석 실패: {e}")
                        import traceback
                        debug_log.append(f"      📜 트레이스백: {traceback.format_exc()}")
                
                # 주요 패턴 추출
                if target_language in speaker_analysis['languages']:
                    speaker_analysis['main_patterns'] = speaker_analysis['languages'][target_language]['patterns']
                    debug_log.append(f"      ✅ 주요 패턴 추출 완료 ({target_language} 기준)")
                
                analysis_result[speaker] = speaker_analysis
                
                # 화자 프로필 업데이트
                try:
                    self.update_or_create_speaker_profile(speaker, speaker_analysis, target_language)
                    debug_log.append(f"      ✅ 화자 프로필 저장 완료")
                except Exception as e:
                    debug_log.append(f"      ❌ 화자 프로필 저장 실패: {e}")
            
            # 6단계: 최종 결과
            debug_log.append("\n🎉 6단계: 분석 완료!")
            debug_log.append(f"   ✅ 성공적으로 분석된 화자: {len(analysis_result)}명")
            debug_log.append(f"   📊 총 처리된 문장: {sum(data['total_sentences'] for data in analysis_result.values())}개")
            
            # 디버그 로그 저장
            self.save_debug_log(debug_log)
            
            return analysis_result
            
        except Exception as e:
            debug_log.append(f"\n💥 오류 발생: {e}")
            debug_log.append(f"📜 상세 오류:")
            import traceback
            debug_log.append(traceback.format_exc())
            
            # 오류 시에도 디버그 로그 저장
            self.save_debug_log(debug_log)
            
            print("\n".join(debug_log))  # 콘솔에도 출력
            return {}
        
        # ===== scenario_manager.py에 추가할 메서드들 =====

    def init_reference_tables(self):
        """레퍼런스 데이터 저장용 테이블 초기화"""
        try:
            import sqlite3
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 레퍼런스 데이터셋 테이블
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS reference_datasets (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    description TEXT,
                    source_type TEXT NOT NULL,  -- 'file' or 'gsheet'
                    source_path TEXT,
                    target_language TEXT,
                    total_speakers INTEGER,
                    total_sentences INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    last_used TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # 레퍼런스 번역 데이터 테이블
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS reference_translations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dataset_id INTEGER,
                    speaker_name TEXT NOT NULL,
                    kr_text TEXT NOT NULL,
                    language TEXT NOT NULL,
                    translated_text TEXT NOT NULL,
                    string_id TEXT,
                    FOREIGN KEY (dataset_id) REFERENCES reference_datasets (id)
                )
            """)
            
            # 인덱스 생성
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ref_speaker ON reference_translations (speaker_name)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ref_dataset ON reference_translations (dataset_id)")
            
            conn.commit()
            conn.close()
            
        except Exception as e:
            print(f"레퍼런스 테이블 초기화 오류: {e}")

    def save_reference_dataset(self, dataset_name, source_info, analysis_result, target_language, confirm_callback=None):
        """분석된 레퍼런스 데이터셋을 DB에 저장"""
        try:
            import sqlite3
            import json
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 기존 동일 이름 데이터셋 확인
            cursor.execute("SELECT id FROM reference_datasets WHERE name = ?", (dataset_name,))
            existing = cursor.fetchone()
            
            if existing:
                # 기존 데이터 업데이트 여부 확인 (콜백 사용)
                if confirm_callback:
                    should_overwrite = confirm_callback(f"'{dataset_name}' 데이터셋이 이미 존재합니다.\n기존 데이터를 덮어쓰시겠습니까?")
                    if not should_overwrite:
                        conn.close()
                        return False
                else:
                    # 콜백이 없으면 자동으로 덮어쓰기
                    pass
                
                # 기존 데이터 삭제
                dataset_id = existing[0]
                cursor.execute("DELETE FROM reference_translations WHERE dataset_id = ?", (dataset_id,))
                cursor.execute("DELETE FROM reference_datasets WHERE id = ?", (dataset_id,))
            
            # 통계 계산
            total_speakers = len(analysis_result)
            total_sentences = sum(data['total_sentences'] for data in analysis_result.values())
            
            # 데이터셋 메타정보 저장
            cursor.execute("""
                INSERT INTO reference_datasets 
                (name, description, source_type, source_path, target_language, total_speakers, total_sentences)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                dataset_name,
                f"자동 생성됨 - {total_speakers}명 화자, {total_sentences}개 문장",
                source_info['type'],  # 'file' or 'gsheet'
                source_info['path'],
                target_language,
                total_speakers,
                total_sentences
            ))
            
            dataset_id = cursor.lastrowid
            
            # 상세 번역 데이터 저장
            saved_count = 0
            for speaker_name, speaker_data in analysis_result.items():
                for lang, lang_data in speaker_data.get('languages', {}).items():
                    examples = lang_data.get('examples', [])
                    for example in examples:
                        cursor.execute("""
                            INSERT INTO reference_translations 
                            (dataset_id, speaker_name, kr_text, language, translated_text, string_id)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (
                            dataset_id,
                            speaker_name,
                            example['kr'],
                            lang,
                            example['translated'],
                            example.get('string_id', '')
                        ))
                        saved_count += 1
            
            conn.commit()
            conn.close()
            
            print(f"레퍼런스 데이터셋 저장 완료: {dataset_name} ({saved_count}개 번역 예시)")
            return True
            
        except Exception as e:
            print(f"레퍼런스 데이터셋 저장 오류: {e}")
            import traceback
            traceback.print_exc()
            return False

    def load_reference_dataset(self, dataset_name):
        """저장된 레퍼런스 데이터셋 로드"""
        try:
            import sqlite3
            from collections import defaultdict
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 데이터셋 정보 확인
            cursor.execute("SELECT id, target_language FROM reference_datasets WHERE name = ?", (dataset_name,))
            dataset_info = cursor.fetchone()
            
            if not dataset_info:
                return None
            
            dataset_id, target_language = dataset_info
            
            # 번역 데이터 로드
            cursor.execute("""
                SELECT speaker_name, kr_text, language, translated_text, string_id
                FROM reference_translations 
                WHERE dataset_id = ?
                ORDER BY speaker_name, language
            """, (dataset_id,))
            
            # 데이터 재구성
            speaker_data = defaultdict(lambda: defaultdict(list))
            
            for speaker_name, kr_text, language, translated_text, string_id in cursor.fetchall():
                speaker_data[speaker_name][language].append({
                    'kr': kr_text,
                    'translated': translated_text,
                    'string_id': string_id
                })
            
            # 마지막 사용 시간 업데이트
            cursor.execute("UPDATE reference_datasets SET last_used = CURRENT_TIMESTAMP WHERE id = ?", (dataset_id,))
            conn.commit()
            conn.close()
            
            return {
                'target_language': target_language,
                'speaker_data': dict(speaker_data)
            }
            
        except Exception as e:
            print(f"레퍼런스 데이터셋 로드 오류: {e}")
            return None

    def get_available_datasets(self):
        """사용 가능한 레퍼런스 데이터셋 목록 반환"""
        try:
            import sqlite3
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT name, description, source_type, target_language, 
                    total_speakers, total_sentences, created_at, last_used
                FROM reference_datasets 
                ORDER BY last_used DESC
            """)
            
            datasets = []
            for row in cursor.fetchall():
                datasets.append({
                    'name': row[0],
                    'description': row[1],
                    'source_type': row[2],
                    'target_language': row[3],
                    'total_speakers': row[4],
                    'total_sentences': row[5],
                    'created_at': row[6],
                    'last_used': row[7]
                })
            
            conn.close()
            return datasets
            
        except Exception as e:
            print(f"데이터셋 목록 조회 오류: {e}")
            return []

    def find_similar_references(self, speaker_name, current_kr_text, target_language="EN", max_examples=8):
        """현재 번역할 문장과 유사한 레퍼런스 찾기 (일관성 개선 핵심)"""
        try:
            from difflib import SequenceMatcher
            import sqlite3
            
            if speaker_name not in self.speakers:
                return []
            
            # DB에서 해당 화자의 모든 레퍼런스 가져오기
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT kr_text, translated_text 
                FROM reference_translations 
                WHERE speaker_name = ? AND language = ?
                ORDER BY id
            """, (speaker_name, target_language))
            
            all_references = cursor.fetchall()
            conn.close()
            
            if not all_references:
                return []
            
            # 유사도 계산 및 정렬
            similarities = []
            for kr_ref, translated_ref in all_references:
                similarity = SequenceMatcher(None, current_kr_text.lower(), kr_ref.lower()).ratio()
                similarities.append({
                    'kr': kr_ref,
                    'translated': translated_ref,
                    'similarity': similarity
                })
            
            # 유사도 순으로 정렬 후 상위 max_examples개 반환
            similarities.sort(key=lambda x: x['similarity'], reverse=True)
            return similarities[:max_examples]
            
        except Exception as e:
            print(f"유사 레퍼런스 검색 오류: {e}")
            return []

    def analyze_tone_patterns(self, references, current_translation):
        """말투 패턴 분석"""
        try:
            # 간단한 말투 지표들
            formal_indicators = ["please", "thank you", "excuse me", "would you", "could you"]
            casual_indicators = ["hey", "yeah", "gonna", "wanna", "ok", "cool"]
            
            # 레퍼런스들의 패턴
            ref_formal_count = sum(sum(indicator in ref['translated'].lower() for indicator in formal_indicators) for ref in references)
            ref_casual_count = sum(sum(indicator in ref['translated'].lower() for indicator in casual_indicators) for ref in references)
            
            # 현재 번역의 패턴
            current_formal = sum(indicator in current_translation.lower() for indicator in formal_indicators)
            current_casual = sum(indicator in current_translation.lower() for indicator in casual_indicators)
            
            # 패턴 매칭 여부
            ref_tends_formal = ref_formal_count > ref_casual_count
            current_tends_formal = current_formal > current_casual
            
            matches = (ref_tends_formal == current_tends_formal) or (ref_formal_count == 0 and ref_casual_count == 0)
            
            return {
                "matches": matches,
                "ref_formal": ref_formal_count,
                "ref_casual": ref_casual_count,
                "current_formal": current_formal,
                "current_casual": current_casual
            }
            
        except Exception as e:
            print(f"말투 패턴 분석 오류: {e}")
            return {"matches": True}
        
    def generate_enhanced_speaker_prompt(self, speaker_name, current_kr_text, target_language="EN"):
        """향상된 화자별 맞춤 프롬프트 생성 (GPT-4o-mini 최적화)"""
        if speaker_name not in self.speakers:
            return self.get_default_enhanced_prompt()
        
        speaker = self.speakers[speaker_name]
        
        # 1. 현재 문장과 유사한 레퍼런스 찾기 (개선된 버전)
        similar_refs = self.find_similar_references(speaker_name, current_kr_text, target_language, max_examples=4)
        
        # 2. 화자 특성 기반 번역 스타일 지침 생성
        style_instructions = self.generate_character_style_guide(speaker)
        
        # 3. GPT-4o-mini 최적화 프롬프트 구성
        prompt = f"""You are localizing dialogue for '{speaker.name}' in a game. Translate Korean to natural English maintaining character consistency.

    CHARACTER: {speaker.name}
    • Gender: {speaker.gender}
    • Speaking Style: {speaker.tone}
    • Personality: {speaker.style}

    TRANSLATION STYLE:
    {style_instructions}

    REFERENCE EXAMPLES (maintain this style):"""

        # 4. 유사한 레퍼런스들을 프롬프트에 포함 (효율적으로)
        if similar_refs:
            for i, ref in enumerate(similar_refs[:3], 1):  # 최대 3개만 사용 (토큰 절약)
                prompt += f"\n{i}. KR: {ref['kr']}"
                prompt += f"\n   EN: {ref['translated']}"
        else:
            # 유사한 레퍼런스가 없으면 일반 예시 사용
            for i, example in enumerate(speaker.examples[:2], 1):
                if isinstance(example, dict):
                    kr_text = example.get('kr', '')
                    en_text = example.get('translated', '') or example.get('en', '')
                    if kr_text and en_text:
                        prompt += f"\n{i}. KR: {kr_text}"
                        prompt += f"\n   EN: {en_text}"

        prompt += f"""

    CURRENT TEXT: {current_kr_text}

    Keep the same character voice and style. Translate naturally:"""
        
        return prompt

    def generate_character_style_guide(self, speaker):
        """화자 특성 기반 번역 스타일 지침 생성"""
        guidelines = []
        
        # 성별 기반 지침
        if speaker.gender == "남성":
            guidelines.append("Use confident, direct language typical of male characters")
        elif speaker.gender == "여성":
            guidelines.append("Use expressive language appropriate for female characters")
        
        # 말투 기반 지침
        tone_guides = {
            "정중": "Formal, polite speech with respectful language",
            "친근": "Casual, friendly tone with contractions and informal expressions",
            "거친": "Rough, direct speech with strong language",
            "귀여운": "Cute, endearing expressions with softer language",
            "격식": "Formal, ceremonial language",
            "보통": "Natural, balanced tone"
        }
        
        if speaker.tone in tone_guides:
            guidelines.append(tone_guides[speaker.tone])
        
        # 스타일 특성 기반 추가 지침
        if "간결" in speaker.style:
            guidelines.append("Keep responses concise and to the point")
        if "감정" in speaker.style:
            guidelines.append("Use emotionally expressive language")
        if "질문" in speaker.style:
            guidelines.append("Maintain questioning patterns when present")
        
        return "• " + "\n• ".join(guidelines) if guidelines else "• Use natural, character-appropriate language"

    def get_default_enhanced_prompt(self):
        """개선된 기본 프롬프트 (화자 정보가 없을 때)"""
        return """Translate Korean game dialogue to natural English for young English speakers.

    GUIDELINES:
    • Sound natural and engaging, not translated
    • Use appropriate tone for the context
    • Preserve special tags: {}, [#color#], etc.
    • Make it feel like native English dialogue

    Text to translate:"""

    def validate_translation_consistency(self, speaker_name, kr_text, translated_text, target_language="EN"):
        """번역 일관성 검증 (개선된 버전)"""
        try:
            if speaker_name not in self.speakers:
                return {"is_consistent": True, "confidence": 0.5, "suggestions": []}
            
            # 유사한 레퍼런스들과 비교
            similar_refs = self.find_similar_references(speaker_name, kr_text, target_language, max_examples=3)
            
            if not similar_refs:
                return {"is_consistent": True, "confidence": 0.3, "suggestions": ["레퍼런스 부족으로 검증 제한적"]}
            
            analysis = {
                "is_consistent": True,
                "confidence": 0.8,
                "suggestions": []
            }
            
            # 1. 길이 패턴 검증 (개선)
            ref_lengths = [len(ref['translated'].split()) for ref in similar_refs]
            avg_length = sum(ref_lengths) / len(ref_lengths)
            current_length = len(translated_text.split())
            
            length_ratio = current_length / avg_length if avg_length > 0 else 1
            if length_ratio > 2.5 or length_ratio < 0.4:  # 기준 완화
                analysis["suggestions"].append(f"번역 길이 검토 필요 (평균: {avg_length:.1f}단어, 현재: {current_length}단어)")
                analysis["confidence"] *= 0.85
            
            # 2. 말투 패턴 검증 (강화)
            tone_analysis = self.analyze_enhanced_tone_patterns(similar_refs, translated_text)
            if not tone_analysis["matches"]:
                analysis["suggestions"].append("화자 말투 패턴 확인 필요")
                analysis["confidence"] *= 0.8
            
            # 3. 특수 표현 일관성 체크
            consistency_score = self.check_expression_consistency(similar_refs, translated_text)
            if consistency_score < 0.7:
                analysis["suggestions"].append("표현 스타일 일관성 검토")
                analysis["confidence"] *= 0.9
            
            # 4. 전체 신뢰도 기반 일관성 판단
            analysis["is_consistent"] = analysis["confidence"] > 0.65  # 기준 완화
            
            return analysis
            
        except Exception as e:
            print(f"일관성 검증 오류: {e}")
            return {"is_consistent": True, "confidence": 0.5, "suggestions": []}

    def analyze_enhanced_tone_patterns(self, references, current_translation):
        """강화된 말투 패턴 분석"""
        try:
            # 확장된 말투 지표들
            tone_indicators = {
                'formal': ["please", "thank you", "excuse me", "would you", "could you", "may I", "I appreciate"],
                'casual': ["hey", "yeah", "gonna", "wanna", "ok", "cool", "awesome", "kinda", "sorta"],
                'rough': ["damn", "hell", "shut up", "get out", "what the", "screw", "crap", "dammit"],
                'cute': ["aww", "yay", "wow", "amazing", "adorable", "sweetie", "cutie", "lovely"],
                'confident': ["absolutely", "definitely", "of course", "obviously", "clearly", "certainly"],
                'questioning': ["right?", "don't you think?", "you know?", "isn't it?", "wouldn't you say?"]
            }
            
            # 레퍼런스들의 패턴 분석
            ref_scores = {tone: 0 for tone in tone_indicators}
            total_refs = len(references)
            
            for ref in references:
                ref_text = ref['translated'].lower()
                for tone, indicators in tone_indicators.items():
                    score = sum(1 for indicator in indicators if indicator in ref_text)
                    ref_scores[tone] += score
            
            # 평균 점수 계산
            ref_avg_scores = {tone: score / total_refs for tone, score in ref_scores.items()}
            
            # 현재 번역의 패턴
            current_text = current_translation.lower()
            current_scores = {}
            for tone, indicators in tone_indicators.items():
                current_scores[tone] = sum(1 for indicator in indicators if indicator in current_text)
            
            # 주요 패턴 매칭 확인
            dominant_ref_tone = max(ref_avg_scores, key=ref_avg_scores.get)
            dominant_current_tone = max(current_scores, key=current_scores.get) if any(current_scores.values()) else dominant_ref_tone
            
            # 매칭도 계산
            matches = (dominant_ref_tone == dominant_current_tone) or (ref_avg_scores[dominant_ref_tone] < 0.5)
            
            return {
                "matches": matches,
                "ref_dominant": dominant_ref_tone,
                "current_dominant": dominant_current_tone,
                "ref_scores": ref_avg_scores,
                "current_scores": current_scores
            }
            
        except Exception as e:
            print(f"말투 패턴 분석 오류: {e}")
            return {"matches": True}

    def check_expression_consistency(self, references, current_translation):
        """표현 스타일 일관성 체크"""
        try:
            if not references:
                return 0.8
            
            # 문장 구조 패턴 분석
            ref_patterns = []
            for ref in references:
                text = ref['translated']
                patterns = {
                    'avg_word_length': sum(len(word) for word in text.split()) / len(text.split()) if text.split() else 0,
                    'exclamation_ratio': text.count('!') / len(text.split()) if text.split() else 0,
                    'question_ratio': text.count('?') / len(text.split()) if text.split() else 0,
                    'contraction_count': sum(1 for word in text.split() if "'" in word),
                    'uppercase_words': sum(1 for word in text.split() if word.isupper() and len(word) > 1)
                }
                ref_patterns.append(patterns)
            
            # 평균 패턴 계산
            avg_patterns = {}
            for key in ref_patterns[0].keys():
                avg_patterns[key] = sum(p[key] for p in ref_patterns) / len(ref_patterns)
            
            # 현재 번역 패턴
            current_text = current_translation
            current_patterns = {
                'avg_word_length': sum(len(word) for word in current_text.split()) / len(current_text.split()) if current_text.split() else 0,
                'exclamation_ratio': current_text.count('!') / len(current_text.split()) if current_text.split() else 0,
                'question_ratio': current_text.count('?') / len(current_text.split()) if current_text.split() else 0,
                'contraction_count': sum(1 for word in current_text.split() if "'" in word),
                'uppercase_words': sum(1 for word in current_text.split() if word.isupper() and len(word) > 1)
            }
            
            # 일관성 점수 계산
            consistency_scores = []
            for key in avg_patterns.keys():
                if avg_patterns[key] == 0 and current_patterns[key] == 0:
                    consistency_scores.append(1.0)
                elif avg_patterns[key] == 0:
                    consistency_scores.append(0.8)  # 참조에 없는 특성이 나타남
                else:
                    ratio = min(current_patterns[key], avg_patterns[key]) / max(current_patterns[key], avg_patterns[key])
                    consistency_scores.append(ratio)
            
            return sum(consistency_scores) / len(consistency_scores)
            
        except Exception as e:
            print(f"표현 일관성 체크 오류: {e}")
            return 0.8
        
        
