# tools/main_logic.py

import os
from PyQt5.QtWidgets import QMessageBox, QApplication

from .translate.translation_request_extractor import TranslationRequestExtractor
from .translate.string_sync_manager import StringSyncManager
from .translate.word_replacement_manager import WordReplacementManager
from .translate.translation_db_manager import TranslationDBManager
from .translate.translation_apply_manager import TranslationApplyManager
from .db_compare_manager import DBCompareManager
from ..ui.progress_manager import ProgressManager


class TranslationLogic:
    def __init__(self, main_window):
        self.main_window = main_window
        self.log_message = main_window.log_message
        self.config = main_window.config
        
        # =====================================================================================
        # 각 매니저 클래스를 생성할 때, self.config 객체를 인자로 전달해줍니다.
        # =====================================================================================
        self.progress_manager = ProgressManager(self.log_message)
        self.extractor = TranslationRequestExtractor(self.log_message, self.progress_manager, self.config)
        self.sync_manager = StringSyncManager(self.log_message, self.progress_manager, self.config)
        self.apply_manager = TranslationApplyManager(self.log_message, self.progress_manager, self.config)
        self.word_replacement_manager = WordReplacementManager(self.main_window, self.log_message, self.config)
        # DBCompareManager에도 config를 전달하도록 수정합니다.
        self.db_compare_manager = DBCompareManager(self.main_window, self.config)

    def run_extraction(self):
        source_file = self.main_window.source_file_path.text()
        template_file = self.main_window.template_file_path.text()
        output_dir = self.main_window.output_dir_path.text()

        if not all([source_file, template_file, output_dir]):
            QMessageBox.warning(self.main_window, "입력 오류", "모든 경로를 설정해야 합니다.")
            return

        try:
            self.log_message("번역 요청서 추출을 시작합니다.")
            self.extractor.extract_and_save(source_file, template_file, output_dir)
            self.log_message("번역 요청서 추출이 완료되었습니다.")
            QMessageBox.information(self.main_window, "완료", "번역 요청서 추출이 완료되었습니다.")
        except Exception as e:
            self.log_message(f"오류 발생: {e}")
            QMessageBox.critical(self.main_window, "오류", f"추출 중 오류가 발생했습니다: {e}")

    def run_synchronization(self):
        source_path = self.main_window.sync_source_path.text()
        target_path = self.main_window.sync_target_path.text()
        db_name = self.main_window.sync_db_selector.currentText()

        if not all([source_path, target_path, db_name]):
            QMessageBox.warning(self.main_window, "입력 오류", "모든 경로와 DB를 선택해야 합니다.")
            return

        try:
            self.log_message(f"{db_name} DB에 원본/번역문 동기화를 시작합니다.")
            self.sync_manager.sync_strings_to_db(source_path, target_path, db_name)
            self.log_message("동기화가 완료되었습니다.")
            QMessageBox.information(self.main_window, "완료", "문자열 동기화가 완료되었습니다.")
        except Exception as e:
            self.log_message(f"오류 발생: {e}")
            QMessageBox.critical(self.main_window, "오류", f"동기화 중 오류가 발생했습니다: {e}")

    def run_word_replacement(self):
        target_file = self.main_window.replace_file_path.text()
        glossary_file = self.main_window.glossary_file_path.text()

        if not all([target_file, glossary_file]):
            QMessageBox.warning(self.main_window, "입력 오류", "대상 파일과 용어집 파일 경로를 모두 지정해야 합니다.")
            return

        try:
            self.log_message("용어 교체를 시작합니다.")
            self.word_replacement_manager.replace_words_in_excel(target_file, glossary_file)
            self.log_message("용어 교체가 완료되었습니다.")
            QMessageBox.information(self.main_window, "완료", "용어 교체가 완료되었습니다.")
        except Exception as e:
            self.log_message(f"오류 발생: {e}")
            QMessageBox.critical(self.main_window, "오류", f"용어 교체 중 오류가 발생했습니다: {e}")

    def run_db_comparison(self):
        db1_name = self.main_window.db1_selector.currentText()
        db2_name = self.main_window.db2_selector.currentText()

        if db1_name == "DB 선택" or db2_name == "DB 선택" or db1_name == db2_name:
            QMessageBox.warning(self.main_window, "입력 오류", "서로 다른 두 개의 DB를 선택해야 합니다.")
            return

        try:
            self.log_message(f"{db1_name}와 {db2_name} 비교를 시작합니다.")
            self.db_compare_manager.compare_databases_and_export(db1_name, db2_name)
            # 완료 메시지는 db_compare_manager 내부에서 처리하도록 변경할 수 있습니다.
            # self.log_message("DB 비교 및 엑셀 출력이 완료되었습니다.")
            # QMessageBox.information(self.main_window, "완료", "DB 비교 및 엑셀 출력이 완료되었습니다.")
        except Exception as e:
            self.log_message(f"오류 발생: {e}")
            QMessageBox.critical(self.main_window, "오류", f"DB 비교 중 오류가 발생했습니다: {e}")

    def run_translation_apply(self):
        source_dir = self.main_window.apply_source_dir.text()
        target_file = self.main_window.apply_target_file.text()
        db_name = self.main_window.apply_db_selector.currentText()

        if not all([source_dir, target_file, db_name]) or db_name == "DB 선택":
            QMessageBox.warning(self.main_window, "입력 오류", "모든 경로와 DB를 올바르게 설정해야 합니다.")
            return

        try:
            self.log_message("번역문 적용을 시작합니다.")
            self.apply_manager.apply_translations_to_excel(source_dir, target_file, db_name)
            self.log_message("번역문 적용이 완료되었습니다.")
            QMessageBox.information(self.main_window, "완료", "번역문 적용이 완료되었습니다.")
        except Exception as e:
            self.log_message(f"오류 발생: {e}")
            QMessageBox.critical(self.main_window, "오류", f"번역문 적용 중 오류가 발생했습니다: {e}")