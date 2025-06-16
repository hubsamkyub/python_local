"""
데이터 품질 관리 모듈
용도: TM/용어집의 빈칸 분석, 자동 보완, 품질 향상
"""
import sqlite3
import json
# import pandas as pd # 개선: 사용되지 않으므로 제거
from typing import Dict, List, Optional, Tuple
from collections import defaultdict, Counter

class DataQualityManager:
    """TM/용어집 데이터 품질 관리 클래스"""

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.visible_langs = ["EN", "CN", "TW", "TH", "PT", "ES", "FR", "DE"]

    def analyze_data_quality(self) -> Dict:
        """데이터 품질 전체 분석"""
        print("📊 데이터 품질 분석 시작...")

        results = {
            'tm_analysis': self.analyze_tm_completeness(),
            'glossary_analysis': self.analyze_glossary_completeness(),
            'consistency_analysis': self.analyze_translation_consistency(),
            'recommendations': []
        }

        # 권장사항 생성
        results['recommendations'] = self._generate_recommendations(results)

        return results

    def analyze_tm_completeness(self) -> Dict:
        """TM 완성도 분석"""
        print("🔍 TM 완성도 분석 중...")
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT kr_text, translations FROM translation_memory")
                
                tm_data = []
                for kr_text, trans_json in cursor.fetchall():
                    try:
                        translations = json.loads(trans_json)
                        tm_data.append({
                            'kr': kr_text,
                            'translations': translations
                        })
                    except json.JSONDecodeError:
                        continue
        except sqlite3.Error as e:
            print(f"DB 오류 (TM 분석): {e}")
            return {}

        total_entries = len(tm_data)
        if total_entries == 0:
            return {'total_entries': 0, 'language_stats': {}}

        lang_stats = {}
        for lang in self.visible_langs:
            filled_count = sum(1 for item in tm_data if item['translations'].get(lang, '').strip())
            lang_stats[lang] = {
                'filled': filled_count,
                'empty': total_entries - filled_count,
                'completeness_rate': filled_count / total_entries
            }

        complete_entries = sum(1 for item in tm_data if all(item['translations'].get(lang, '').strip() for lang in self.visible_langs))

        return {
            'total_entries': total_entries,
            'complete_entries': complete_entries,
            'complete_rate': complete_entries / total_entries,
            'language_stats': lang_stats,
            'most_complete_lang': max(lang_stats, key=lambda x: lang_stats[x]['completeness_rate']),
            'least_complete_lang': min(lang_stats, key=lambda x: lang_stats[x]['completeness_rate'])
        }

    def analyze_glossary_completeness(self) -> Dict:
        """용어집 완성도 분석"""
        print("📚 용어집 완성도 분석 중...")
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                # 용어집 데이터 로드 (STRING_ID 제거된 새 스키마)
                cursor.execute("SELECT kr, en, cn, tw, th, pt, es, de, fr FROM glossary")
                columns = [desc[0] for desc in cursor.description]
                glossary_data = [dict(zip(columns, row)) for row in cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"DB 오류 (용어집 분석): {e}")
            return {}

        total_entries = len(glossary_data)
        if total_entries == 0:
            return {'total_entries': 0, 'language_stats': {}}
            
        lang_stats = {}
        target_langs = ['en', 'cn', 'tw', 'th', 'pt', 'es', 'de', 'fr']

        for lang in target_langs:
            # ### 수정된 부분: 로직을 명확하고 단순하게 변경 ###
            filled_count = sum(1 for item in glossary_data if item.get(lang, '').strip())
            
            lang_stats[lang.upper()] = {
                'filled': filled_count,
                'empty': total_entries - filled_count,
                'completeness_rate': filled_count / total_entries
            }

        return {
            'total_entries': total_entries,
            'language_stats': lang_stats
        }

    def analyze_translation_consistency(self) -> Dict:
        """번역 일관성 분석 (같은 한국어의 다른 번역들)"""
        print("🔍 번역 일관성 분석 중...")

        kr_variations = defaultdict(lambda: defaultdict(list))
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT kr_text, translations FROM translation_memory")
                
                for kr_text, trans_json in cursor.fetchall():
                    try:
                        translations = json.loads(trans_json)
                        for lang, trans_text in translations.items():
                            if trans_text and trans_text.strip():
                                kr_variations[kr_text][lang].append(trans_text.strip())
                    except json.JSONDecodeError:
                        continue
        except sqlite3.Error as e:
            print(f"DB 오류 (일관성 분석): {e}")
            return {}

        inconsistent_items = []
        for kr_text, lang_dict in kr_variations.items():
            for lang, trans_list in lang_dict.items():
                unique_translations = list(set(trans_list))
                if len(unique_translations) > 1:
                    inconsistent_items.append({
                        'korean': kr_text,
                        'language': lang,
                        'variations': unique_translations,
                        'variation_count': len(unique_translations)
                    })
        
        inconsistent_items.sort(key=lambda x: x['variation_count'], reverse=True)

        return {
            'total_inconsistent': len(inconsistent_items),
            'inconsistent_items': inconsistent_items[:20],  # 상위 20개만
            'languages_with_issues': list(set(item['language'] for item in inconsistent_items))
        }

    # ### 수정된 부분: DB 연결 관리를 with 구문으로 변경하여 안정성 확보 ###
    def auto_fill_missing_translations(self, dry_run: bool = True) -> Dict:
        """빈 번역을 자동으로 채우기"""
        print(f"🔧 누락된 번역 자동 보완 시작... (모드: {'시뮬레이션' if dry_run else '실제 적용'})")
        
        results = {
            'tm_filled': 0,
            'glossary_filled': 0,
            'fill_patterns': defaultdict(int),
            'errors': []
        }

        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # 1. TM 보완
                tm_fill_result = self._auto_fill_tm(cursor)
                results['tm_filled'] = tm_fill_result['filled_count']
                results['fill_patterns'].update(tm_fill_result['patterns'])
                
                # 2. 용어집 보완
                glossary_fill_result = self._auto_fill_glossary(cursor)
                results['glossary_filled'] = glossary_fill_result['filled_count']
                results['fill_patterns'].update(glossary_fill_result['patterns'])

                # 실제 적용 시에만 DB에 커밋
                if not dry_run:
                    conn.commit()
                    print(f"✅ 자동 보완 완료: TM {results['tm_filled']}개, 용어집 {results['glossary_filled']}개")
                else:
                    conn.rollback() # 시뮬레이션이므로 롤백
                    print(f"🔍 시뮬레이션 결과: TM {results['tm_filled']}개, 용어집 {results['glossary_filled']}개 보완 가능")

        except sqlite3.Error as e:
            results['errors'].append(str(e))
            print(f"❌ 자동 보완 중 DB 오류 발생: {e}")
        except Exception as e:
            results['errors'].append(str(e))
            print(f"❌ 자동 보완 중 예기치 않은 오류 발생: {e}")

        return results

    # ### 수정된 부분: cursor를 인자로 받아 DB 연결 재사용 ###
    def _auto_fill_tm(self, cursor: sqlite3.Cursor) -> Dict:
        """TM 자동 보완 (헬퍼 메서드)"""
        cursor.execute("SELECT rowid, kr_text, translations FROM translation_memory")
        
        all_tm_data = cursor.fetchall()
        filled_count = 0
        patterns = defaultdict(int)
        
        for rowid, kr_text, trans_json in all_tm_data:
            try:
                translations = json.loads(trans_json)
                original_translations = translations.copy()
                
                # EN 기준 다국어 추론
                if translations.get('EN'):
                    en_text = translations['EN']
                    
                    # 간단한 패턴 기반 보완 (실제로는 더 복잡한 로직이나 API 호출 필요)
                    if not translations.get('CN') and 'Chinese' in self._get_language_hints(en_text):
                        translations['CN'] = f"[AUTO_FILLED_CN] {en_text}"
                        patterns['EN→CN'] += 1
                
                # 변경사항이 있으면 업데이트
                if translations != original_translations:
                    filled_count += 1
                    # dry_run 여부는 상위 메서드에서 commit/rollback으로 제어하므로 여기서는 항상 execute
                    cursor.execute(
                        "UPDATE translation_memory SET translations = ? WHERE rowid = ?",
                        (json.dumps(translations, ensure_ascii=False), rowid)
                    )
            except (json.JSONDecodeError, TypeError):
                continue
        
        return {'filled_count': filled_count, 'patterns': dict(patterns)}

    def _auto_fill_glossary(self, cursor: sqlite3.Cursor) -> Dict:
        """용어집 자동 보완 (헬퍼 메서드)"""
        # 현재는 시뮬레이션만 구현, 실제 로직 추가 필요
        return {'filled_count': 0, 'patterns': {}}

    def _get_language_hints(self, text: str) -> List[str]:
        """텍스트에서 언어 힌트 추출 (단순 버전)"""
        hints = []
        if any(c in text for c in '中国台湾繁体简体'): hints.append('Chinese')
        if any(c in text for c in 'ไทยวาดภาษา'): hints.append('Thai')
        return hints

    def _generate_recommendations(self, analysis_results: Dict) -> List[str]:
        """분석 결과 기반 권장사항 생성"""
        recommendations = []
        
        tm_analysis = analysis_results.get('tm_analysis', {})
        consistency_analysis = analysis_results.get('consistency_analysis', {})
        
        if not tm_analysis: return ["TM 분석 데이터를 가져올 수 없어 권장사항을 생성할 수 없습니다."]

        if tm_analysis.get('complete_rate', 1.0) < 0.3:
            recommendations.append(
                f"🚨 TM 완성도가 매우 낮습니다 ({tm_analysis['complete_rate']:.1%}). "
                f"'{tm_analysis.get('most_complete_lang', '알 수 없음')}' 기준으로 다른 언어 번역을 보완하세요."
            )
        
        if consistency_analysis.get('total_inconsistent', 0) > 50:
            issues = consistency_analysis.get('languages_with_issues', [])
            recommendations.append(
                f"⚠️ 번역 일관성 문제가 {consistency_analysis['total_inconsistent']}개 발견되었습니다. "
                f"특히 {', '.join(issues[:3])} 언어를 점검하세요."
            )

        for lang, stats in tm_analysis.get('language_stats', {}).items():
            if stats['completeness_rate'] < 0.5:
                recommendations.append(
                    f"📝 {lang} 언어 완성도가 낮습니다 ({stats['completeness_rate']:.1%}). 우선적으로 보완이 필요합니다."
                )
        return recommendations

    def export_quality_report(self, analysis_results: Dict, file_path: str = None):
        """품질 분석 리포트 내보내기"""
        if not file_path:
            from datetime import datetime
            file_path = f"data_quality_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("=== 데이터 품질 분석 리포트 ===\n\n")
            
            # TM 분석 결과
            tm = analysis_results.get('tm_analysis', {})
            f.write(f"📊 TM 분석:\n")
            f.write(f"  총 항목: {tm.get('total_entries', 0):,}개\n")
            f.write(f"  완전한 항목: {tm.get('complete_entries', 0):,}개 ({tm.get('complete_rate', 0):.1%})\n")
            f.write(f"  가장 완성도 높은 언어: {tm.get('most_complete_lang', 'N/A')}\n")
            f.write(f"  가장 완성도 낮은 언어: {tm.get('least_complete_lang', 'N/A')}\n\n")
            
            # 언어별 상세
            f.write("언어별 완성도:\n")
            for lang, stats in tm.get('language_stats', {}).items():
                f.write(f"  {lang}: {stats.get('filled', 0):,}개 ({stats.get('completeness_rate', 0):.1%})\n")
            f.write("\n")
            
            # 용어집 분석 결과
            glossary = analysis_results.get('glossary_analysis', {})
            f.write(f"📚 용어집 분석:\n")
            f.write(f"  총 항목: {glossary.get('total_entries', 0):,}개\n")
            for lang, stats in glossary.get('language_stats', {}).items():
                f.write(f"  {lang}: {stats.get('filled', 0):,}개 ({stats.get('completeness_rate', 0):.1%})\n")
            f.write("\n")
            
            # 일관성 분석 결과
            consistency = analysis_results.get('consistency_analysis', {})
            f.write(f"🔍 일관성 분석:\n")
            f.write(f"  일관성 문제: {consistency.get('total_inconsistent', 0)}개\n")
            f.write(f"  문제 언어: {', '.join(consistency.get('languages_with_issues', []))}\n\n")
            
            # 권장사항
            f.write("💡 권장사항:\n")
            for i, rec in enumerate(analysis_results.get('recommendations', []), 1):
                f.write(f"  {i}. {rec}\n")
        
        print(f"📄 리포트 저장 완료: {file_path}")
        return file_path

# 스마트 번역 매니저 통합용 메서드
class SmartTranslationManagerIntegration:
    """기존 SmartTranslationManager와의 통합 인터페이스"""
    
    @staticmethod
    def add_data_quality_buttons(tm_management_tab):
        """TM 관리 탭에 데이터 품질 버튼 추가"""
        import tkinter as tk
        from tkinter import ttk
        
        # 기존 TM 관리 탭에 새 버튼 추가
        quality_frame = ttk.LabelFrame(tm_management_tab, text="📊 데이터 품질 관리")
        quality_frame.pack(fill="x", padx=10, pady=5)
        
        button_frame = ttk.Frame(quality_frame)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        # 품질 분석 버튼
        ttk.Button(button_frame, text="📊 품질 분석", 
                   command=lambda: SmartTranslationManagerIntegration.run_quality_analysis(tm_management_tab), 
                   width=12).pack(side="left", padx=2)
        
        # 자동 보완 버튼  
        ttk.Button(button_frame, text="🔧 자동 보완", 
                   command=lambda: SmartTranslationManagerIntegration.run_auto_improvement(tm_management_tab),
                   width=12).pack(side="left", padx=2)
        
        # 리포트 버튼
        ttk.Button(button_frame, text="📄 품질 리포트",
                   command=lambda: SmartTranslationManagerIntegration.generate_report(tm_management_tab),
                   width=12).pack(side="left", padx=2)
    
    @staticmethod 
    def run_quality_analysis(parent_widget):
        """품질 분석 실행"""
        from tkinter import messagebox
        try:
            # 부모 위젯에서 DB 경로 가져오기 (winfo_toplevel()로 최상위 앱 인스턴스 접근)
            # 참고: 이 방식은 GUI 구조에 따라 불안정할 수 있으므로 주의가 필요합니다.
            main_app = parent_widget.winfo_toplevel()
            db_path = getattr(main_app, 'translation_db_path', 'smart_translations.db')
            
            manager = DataQualityManager(db_path)
            analysis = manager.analyze_data_quality()
            
            # 결과를 새 창에 표시
            SmartTranslationManagerIntegration.show_analysis_results(parent_widget, analysis)
            
        except Exception as e:
            messagebox.showerror("분석 오류", f"품질 분석 중 오류 발생: {e}")
    
    @staticmethod
    def run_auto_improvement(parent_widget):
        """자동 개선 실행"""
        from tkinter import messagebox
        try:
            if not messagebox.askyesno("자동 보완 확인", 
                                       "TM/용어집의 빈 칸을 자동으로 보완하시겠습니까?\n"
                                       "먼저 시뮬레이션을 실행하여 결과를 확인합니다."):
                return
            
            main_app = parent_widget.winfo_toplevel()
            db_path = getattr(main_app, 'translation_db_path', 'smart_translations.db')
            
            manager = DataQualityManager(db_path)
            
            # 시뮬레이션 먼저 실행
            sim_results = manager.auto_fill_missing_translations(dry_run=True)
            
            sim_msg = f"""시뮬레이션 결과:
            
TM 보완 가능: {sim_results['tm_filled']}개
용어집 보완 가능: {sim_results['glossary_filled']}개

실제로 적용하시겠습니까?"""
            
            if messagebox.askyesno("시뮬레이션 결과", sim_msg):
                # 실제 실행
                real_results = manager.auto_fill_missing_translations(dry_run=False)
                messagebox.showinfo("보완 완료", 
                                    f"데이터 보완이 완료되었습니다!\n"
                                    f"TM: {real_results['tm_filled']}개\n"
                                    f"용어집: {real_results['glossary_filled']}개")
                
                # 부모 앱의 데이터 리로드 메서드가 있다면 호출
                if hasattr(main_app, 'load_translation_memory'):
                    main_app.load_translation_memory()
                    
        except Exception as e:
            messagebox.showerror("보완 오류", f"자동 보완 중 오류 발생: {e}")
    
    @staticmethod
    def generate_report(parent_widget):
        """품질 리포트 생성"""
        from tkinter import messagebox
        try:
            main_app = parent_widget.winfo_toplevel()
            db_path = getattr(main_app, 'translation_db_path', 'smart_translations.db')
            
            manager = DataQualityManager(db_path)
            analysis = manager.analyze_data_quality()
            report_file = manager.export_quality_report(analysis)
            
            messagebox.showinfo("리포트 생성 완료", f"품질 리포트가 생성되었습니다:\n{report_file}")
            
        except Exception as e:
            messagebox.showerror("리포트 오류", f"리포트 생성 중 오류 발생: {e}")
    
    @staticmethod
    def show_analysis_results(parent_widget, analysis_results):
        """분석 결과를 새 창에 표시"""
        import tkinter as tk
        from tkinter import ttk
        
        # 새 창 생성
        result_window = tk.Toplevel(parent_widget)
        result_window.title("📊 데이터 품질 분석 결과")
        result_window.geometry("800x600")
        
        main_frame = ttk.Frame(result_window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 탭 생성
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True, pady=5)
        
        # --- 요약 및 권장사항 탭 ---
        summary_tab = ttk.Frame(notebook, padding="10")
        notebook.add(summary_tab, text="💡 요약 및 권장사항")
        
        summary_text = tk.Text(summary_tab, wrap="word", font=("맑은 고딕", 10), relief="flat", background=summary_tab.cget('bg'))
        summary_text.pack(fill="both", expand=True)
        
        rec_content = "💡 데이터 품질 개선 권장사항:\n\n"
        for i, rec in enumerate(analysis_results.get('recommendations', []), 1):
            summary_text.insert(tk.END, f"{i}. {rec}\n\n")
        summary_text.config(state="disabled")

        # --- TM 분석 탭 ---
        tm_tab = ttk.Frame(notebook, padding="10")
        notebook.add(tm_tab, text="TM 분석")
        
        tm_text = tk.Text(tm_tab, wrap="none", font=("Consolas", 10))
        tm_scroll_y = ttk.Scrollbar(tm_tab, orient="vertical", command=tm_text.yview)
        tm_scroll_x = ttk.Scrollbar(tm_tab, orient="horizontal", command=tm_text.xview)
        tm_text.configure(yscrollcommand=tm_scroll_y.set, xscrollcommand=tm_scroll_x.set)
        
        tm_analysis = analysis_results.get('tm_analysis', {})
        tm_content = f"""📊 TM 완성도 분석 결과

총 항목 수: {tm_analysis.get('total_entries', 0):,}개
완전한 항목: {tm_analysis.get('complete_entries', 0):,}개 ({tm_analysis.get('complete_rate', 0):.1%})

가장 완성도 높은 언어: {tm_analysis.get('most_complete_lang', 'N/A')}
가장 완성도 낮은 언어: {tm_analysis.get('least_complete_lang', 'N/A')}

{"-"*50}
언어별 상세 통계:
{"-"*50}
"""
        tm_text.insert("1.0", tm_content)
        
        for lang, stats in sorted(tm_analysis.get('language_stats', {}).items()):
            line = f"{lang:>4}: {stats.get('filled', 0):>6,}개 ({stats.get('completeness_rate', 0):>6.1%}) | 빈칸: {stats.get('empty', 0):>6,}개\n"
            tm_text.insert(tk.END, line)
        
        tm_text.config(state="disabled")
        
        tm_scroll_y.pack(side="right", fill="y")
        tm_scroll_x.pack(side="bottom", fill="x")
        tm_text.pack(side="left", fill="both", expand=True)

        # --- 일관성 분석 탭 ---
        consistency_tab = ttk.Frame(notebook, padding="10")
        notebook.add(consistency_tab, text="일관성 분석")

        cons_text = tk.Text(consistency_tab, wrap="word", font=("맑은 고딕", 10))
        cons_scroll = ttk.Scrollbar(consistency_tab, command=cons_text.yview)
        cons_text.configure(yscrollcommand=cons_scroll.set)
        
        consistency = analysis_results.get('consistency_analysis', {})
        cons_content = f"🔍 총 {consistency.get('total_inconsistent', 0)}개의 일관성 문제 발견\n\n"
        cons_text.insert("1.0", cons_content)
        
        for item in consistency.get('inconsistent_items', []):
            item_str = f"KR: {item['korean']} ({item['language']})\n"
            for i, variation in enumerate(item['variations']):
                item_str += f"  └─ v{i+1}: {variation}\n"
            item_str += "\n"
            cons_text.insert(tk.END, item_str)

        cons_text.config(state="disabled")
        cons_text.pack(side="left", fill="both", expand=True)
        cons_scroll.pack(side="right", fill="y")
        
        # 닫기 버튼
        ttk.Button(main_frame, text="닫기", command=result_window.destroy).pack(pady=10, side="bottom")