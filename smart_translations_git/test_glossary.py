"""
용어집 매칭 시스템 종합 테스트
사용법: python test_glossary.py
"""
import time
import os
import sys

def test_text_preprocessor():
    """텍스트 전처리 모듈 테스트"""
    print("=" * 60)
    print("🧪 1. TextPreprocessor 테스트")
    print("=" * 60)
    
    try:
        from text_preprocessor import TextPreprocessor
        
        processor = TextPreprocessor()
        
        test_cases = [
            "[@michel]이 [무기]를 장착했습니다",
            "[#color#red#]경고[#color#] 메시지입니다",
            "캐릭터가 스킬을 사용했습니다",
            "[@system] [아이템] 획득! [@player_name] 레벨업 [#sound#beep#]",
            "[무기] [방어구] [스킬] 모두 업그레이드!",
            "[@variable] [일반태그] [#special#] 혼합 테스트"
        ]
        
        for i, text in enumerate(test_cases, 1):
            print(f"\n📝 테스트 {i}: {text}")
            
            start_time = time.time()
            debug_info = processor.get_debug_info(text)
            process_time = (time.time() - start_time) * 1000  # ms
            
            print(f"   🔍 검색대상: '{debug_info['searchable']}'")
            print(f"   🚫 제외태그: {debug_info['excluded_tags']}")
            print(f"   ✅ 포함태그: {debug_info['included_tags']}")
            print(f"   📏 길이감소: {debug_info['length_reduction']}글자")
            print(f"   ⏱️ 처리시간: {process_time:.2f}ms")
        
        print(f"\n📊 전체 처리 통계:")
        stats = processor.get_stats()
        for key, value in stats.items():
            print(f"   {key}: {value}")
        
        print("✅ TextPreprocessor 테스트 완료")
        return True
        
    except ImportError:
        print("❌ text_preprocessor 모듈을 찾을 수 없습니다.")
        return False
    except Exception as e:
        print(f"❌ TextPreprocessor 테스트 실패: {e}")
        return False

def test_glossary_matcher():
    """용어집 매칭 모듈 테스트"""
    print("\n" + "=" * 60)
    print("🧪 2. SmartGlossaryMatcher 테스트")
    print("=" * 60)
    
    # DB 파일 존재 확인
    db_path = "smart_translations.db"
    if not os.path.exists(db_path):
        print(f"❌ DB 파일을 찾을 수 없습니다: {db_path}")
        print("   다른 경로의 DB 파일을 사용하려면 코드를 수정하세요.")
        return False
    
    try:
        from glossary_matcher import SmartGlossaryMatcher
        
        print(f"📂 DB 파일 로드 중: {db_path}")
        start_time = time.time()
        matcher = SmartGlossaryMatcher(db_path)
        load_time = time.time() - start_time
        
        # 성능 통계 확인
        stats = matcher.get_performance_stats()
        print(f"📊 로드 완료: {stats['glossary_terms']}개 용어, {load_time:.2f}초")
        
        # 테스트 케이스들
        test_cases = [
            "[@michel]이 [무기]를 장착했습니다",
            "캐릭터가 새로운 스킬을 배웠습니다",
            "[#color#red#]경고[#color#] 메시지입니다",
            "[@system] [아이템] 획득! 레벨업!",
            "[전설의검] [마법방패] [고급물약] 조합",
            "빈 텍스트 테스트를 위한 일반 문장"
        ]
        
        for i, text in enumerate(test_cases, 1):
            print(f"\n📝 테스트 {i}: {text}")
            
            # 용어 검색 성능 측정
            start_time = time.time()
            relevant_terms = matcher.find_relevant_terms(text, max_terms=8)
            search_time = (time.time() - start_time) * 1000  # ms
            
            print(f"   🎯 관련용어: {relevant_terms}")
            print(f"   📊 매칭개수: {len(relevant_terms)}개")
            print(f"   ⏱️ 검색시간: {search_time:.2f}ms")
            
            # 프롬프트 생성 테스트
            base_prompt = "다음 텍스트를 영어로 번역하세요:"
            enhanced_prompt = matcher.create_enhanced_prompt(text, base_prompt)
            
            improvement = len(enhanced_prompt) - len(base_prompt) - len(text)
            print(f"   📝 프롬프트: +{improvement}글자 개선")
            
            # 상세 디버그 정보 (첫 번째 케이스만)
            if i == 1:
                debug_info = matcher.get_debug_info(text)
                print(f"   🔍 디버그:")
                print(f"      - 원본 매칭: {len(debug_info.get('raw_matches', []))}개")
                print(f"      - 우선순위 적용 후: {len(debug_info.get('prioritized_matches', []))}개")
        
        # 성능 통계 출력
        print(f"\n📊 최종 성능 통계:")
        final_stats = matcher.get_performance_stats()
        for key, value in final_stats.items():
            if isinstance(value, float):
                print(f"   {key}: {value:.3f}")
            else:
                print(f"   {key}: {value}")
        
        print("✅ SmartGlossaryMatcher 테스트 완료")
        return True
        
    except ImportError:
        print("❌ glossary_matcher 모듈을 찾을 수 없습니다.")
        return False
    except Exception as e:
        print(f"❌ SmartGlossaryMatcher 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_integration():
    """통합 테스트 - 실제 사용 시나리오"""
    print("\n" + "=" * 60)
    print("🧪 3. 통합 테스트 - 실제 번역 시나리오")
    print("=" * 60)
    
    try:
        from glossary_matcher import SmartGlossaryMatcher
        
        db_path = "smart_translations.db"
        if not os.path.exists(db_path):
            print(f"❌ DB 파일 없음: {db_path}")
            return False
        
        matcher = SmartGlossaryMatcher(db_path)
        
        # 실제 번역 시나리오 테스트
        translation_scenarios = [
            {
                "korean": "[@player]가 [레전드 무기]를 장착하고 [파이어볼] 스킬을 사용했습니다",
                "base_prompt": "다음 한국어를 자연스러운 영어로 번역하세요:"
            },
            {
                "korean": "[#color#red#]경고![#color#] 던전 입장 시 [HP포션]이 필요합니다",
                "base_prompt": "게임 UI 텍스트를 영어로 번역하세요:"
            },
            {
                "korean": "캐릭터 레벨업! 새로운 스킬을 배울 수 있습니다",
                "base_prompt": "다음 메시지를 영어로 번역하세요:"
            }
        ]
        
        total_token_saved = 0
        
        for i, scenario in enumerate(translation_scenarios, 1):
            korean_text = scenario["korean"]
            base_prompt = scenario["base_prompt"]
            
            print(f"\n📝 시나리오 {i}: {korean_text}")
            
            # 1. 기본 프롬프트
            basic_prompt = f"{base_prompt}\n\n{korean_text}"
            basic_length = len(basic_prompt)
            
            # 2. 용어집 적용 프롬프트
            enhanced_prompt = matcher.create_enhanced_prompt(korean_text, base_prompt)
            enhanced_length = len(enhanced_prompt)
            
            # 3. 개선 효과 분석
            token_increase = enhanced_length - basic_length
            glossary_portion = enhanced_length - basic_length - len(korean_text)
            
            print(f"   📊 분석:")
            print(f"      - 기본 프롬프트: {basic_length}글자")
            print(f"      - 용어집 프롬프트: {enhanced_length}글자")
            print(f"      - 용어집 추가분: {glossary_portion}글자")
            print(f"      - 예상 토큰 증가: ~{token_increase // 3}토큰")
            
            # 4. 용어집 효과 확인
            relevant_terms = matcher.find_relevant_terms(korean_text)
            if relevant_terms:
                print(f"   🎯 적용 용어: {relevant_terms}")
                print(f"   ✅ 번역 일관성 향상 예상")
            else:
                print(f"   ℹ️ 관련 용어 없음 - 기본 번역 방식")
            
            total_token_saved += max(0, glossary_portion)
        
        print(f"\n📈 통합 테스트 결과:")
        print(f"   - 총 {len(translation_scenarios)}개 시나리오 테스트")
        print(f"   - 예상 추가 토큰: ~{total_token_saved // 3}토큰")
        print(f"   - 번역 일관성 향상: 용어집 매칭된 항목들")
        print("✅ 통합 테스트 완료")
        
        return True
        
    except Exception as e:
        print(f"❌ 통합 테스트 실패: {e}")
        return False

def test_performance():
    """성능 테스트"""
    print("\n" + "=" * 60)
    print("🧪 4. 성능 테스트")
    print("=" * 60)
    
    try:
        from glossary_matcher import SmartGlossaryMatcher
        
        db_path = "smart_translations.db"
        if not os.path.exists(db_path):
            print(f"❌ DB 파일 없음: {db_path}")
            return False
        
        # 초기화 시간 측정
        print("⏱️ 시스템 초기화 성능 측정...")
        start_time = time.time()
        matcher = SmartGlossaryMatcher(db_path)
        init_time = time.time() - start_time
        
        stats = matcher.get_performance_stats()
        print(f"   초기화 시간: {init_time:.2f}초")
        print(f"   로드된 용어: {stats['glossary_terms']}개")
        print(f"   메모리 효율성: {stats['memory_efficiency']:.3f}")
        
        # 검색 성능 테스트
        print("\n⏱️ 검색 성능 측정...")
        test_texts = [
            "짧은 텍스트",
            "캐릭터가 무기를 장착하고 스킬을 사용하는 중간 길이 텍스트",
            "[@player]가 [레전드 무기]를 장착하고 [@enemy]와 전투를 시작했습니다. [파이어볼] 스킬과 [힐링] 스킬을 번갈아 사용하면서 [HP포션]과 [MP포션]을 소비하고 있습니다. [#color#red#]위험![#color#] 신호가 나타났습니다."
        ]
        
        for i, text in enumerate(test_texts):
            # 여러 번 실행하여 평균 시간 측정
            times = []
            for _ in range(10):
                start = time.time()
                relevant_terms = matcher.find_relevant_terms(text)
                times.append((time.time() - start) * 1000)  # ms
            
            avg_time = sum(times) / len(times)
            min_time = min(times)
            max_time = max(times)
            
            print(f"   텍스트 {i+1} ({len(text)}글자): {avg_time:.2f}ms (범위: {min_time:.2f}~{max_time:.2f}ms)")
            print(f"      매칭 결과: {len(relevant_terms)}개 용어")
        
        # 캐시 효과 테스트
        print("\n⏱️ 캐시 효과 측정...")
        test_text = "캐릭터가 무기를 장착했습니다"
        
        # 첫 번째 검색 (캐시 미스)
        start = time.time()
        matcher.find_relevant_terms(test_text)
        first_time = (time.time() - start) * 1000
        
        # 두 번째 검색 (캐시 히트)
        start = time.time()
        matcher.find_relevant_terms(test_text)
        second_time = (time.time() - start) * 1000
        
        speedup = first_time / max(second_time, 0.001)
        print(f"   캐시 미스: {first_time:.2f}ms")
        print(f"   캐시 히트: {second_time:.2f}ms")
        print(f"   속도 향상: {speedup:.1f}배")
        
        # 최종 성능 통계
        final_stats = matcher.get_performance_stats()
        print(f"\n📊 최종 캐시 통계:")
        print(f"   히트율: {final_stats['cache_hit_rate']:.1%}")
        print(f"   캐시 크기: {final_stats['cache_size']}개")
        
        print("✅ 성능 테스트 완료")
        return True
        
    except Exception as e:
        print(f"❌ 성능 테스트 실패: {e}")
        return False

def main():
    """메인 테스트 함수"""
    print("🚀 용어집 매칭 시스템 종합 테스트 시작")
    print(f"📅 테스트 시간: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 파이썬 버전 확인
    print(f"🐍 Python 버전: {sys.version}")
    
    test_results = []
    
    # 개별 모듈 테스트
    test_results.append(("TextPreprocessor", test_text_preprocessor()))
    test_results.append(("SmartGlossaryMatcher", test_glossary_matcher()))
    test_results.append(("Integration", test_integration()))
    test_results.append(("Performance", test_performance()))
    
    # 결과 요약
    print("\n" + "=" * 60)
    print("📋 테스트 결과 요약")
    print("=" * 60)
    
    passed = 0
    total = len(test_results)
    
    for test_name, result in test_results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status} {test_name}")
        if result:
            passed += 1
    
    print(f"\n📊 전체 결과: {passed}/{total} 테스트 통과")
    
    if passed == total:
        print("🎉 모든 테스트가 성공적으로 완료되었습니다!")
        print("✨ 용어집 매칭 시스템을 사용할 준비가 되었습니다.")
    else:
        print("⚠️ 일부 테스트가 실패했습니다.")
        print("💡 실패한 테스트의 오류 메시지를 확인하고 문제를 해결하세요.")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)