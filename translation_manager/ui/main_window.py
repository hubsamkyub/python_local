# ui/main_window.py

import sys
import logging
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QTabWidget, QTextEdit, QHBoxLayout, QLabel, QComboBox, QFileDialog)
from PyQt5.QtGui import QIcon

from ..tools.main_logic import TranslationLogic
from ..tools.translate.translation_db_manager import TranslationDBManager
from ..utils.config_utils import AppConfig
from .common_components import create_group_box, create_file_input, create_button, create_db_selector


class TranslationTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = AppConfig("config.ini")
        self.db_manager = TranslationDBManager(self)
        
        # setup_logging()과 init_ui()가 config를 사용하므로 먼저 호출합니다.
        self.setup_logging()
        self.logic = TranslationLogic(self)
        self.init_ui()

    def setup_logging(self):
        # =====================================================================================
        # 하드코딩된 로그 파일 이름 대신, self.config에서 값을 읽어옵니다.
        # =====================================================================================
        log_file = self.config.get('Paths', 'log_file', fallback='translation_tool.log')
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            handlers=[logging.FileHandler(log_file, 'w', 'utf-8'),
                                      logging.StreamHandler()])
        self.logger = logging.getLogger(__name__)

    def log_message(self, message):
        self.logger.info(message)
        self.log_display.append(message)
        QApplication.processEvents()

    def init_ui(self):
        # =====================================================================================
        # 하드코딩된 창 제목과 아이콘 경로 대신, self.config에서 값을 읽어옵니다.
        # =====================================================================================
        window_title = self.config.get('Application', 'title', fallback='번역 요청 툴')
        self.setWindowTitle(window_title)
        self.setGeometry(100, 100, 1200, 800)
        
        icon_path = self.config.get('Application', 'icon', fallback='icon.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.setup_translation_request_tab()
        self.setup_string_sync_tab()
        self.setup_word_replacement_tab()
        self.setup_db_compare_tab()
        self.setup_translation_apply_tab()

        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        main_layout.addWidget(self.log_display)

    def setup_translation_request_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.tabs.addTab(tab, "번역 요청서 추출")
        
        path_group_box = create_group_box("경로 설정")
        path_layout = QVBoxLayout()
        self.source_file_path, _ = create_file_input("번역 대상 파일:", self, path_layout, is_folder=False)
        self.template_file_path, _ = create_file_input("템플릿 파일:", self, path_layout, is_folder=False)
        self.output_dir_path, _ = create_file_input("결과물 저장 폴더:", self, path_layout, is_folder=True)
        path_group_box.setLayout(path_layout)
        layout.addWidget(path_group_box)
        
        self.extract_button = create_button("번역 요청서 추출 실행", self.logic.run_extraction, layout)
        layout.addStretch(1)

    def setup_string_sync_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.tabs.addTab(tab, "원본/번역문 동기화")

        path_group_box = create_group_box("파일/폴더 경로")
        path_layout = QVBoxLayout()
        self.sync_source_path, _ = create_file_input("기준 원본 파일/폴더:", self, path_layout)
        self.sync_target_path, _ = create_file_input("번역문 파일/폴더:", self, path_layout)
        path_group_box.setLayout(path_layout)
        layout.addWidget(path_group_box)

        db_group_box = create_group_box("DB 선택")
        db_layout = QVBoxLayout()
        self.sync_db_selector, _ = create_db_selector(self, db_layout)
        db_group_box.setLayout(db_layout)
        layout.addWidget(db_group_box)

        self.sync_button = create_button("동기화 실행", self.logic.run_synchronization, layout)
        layout.addStretch(1)

    def setup_word_replacement_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.tabs.addTab(tab, "용어 교체")

        path_group_box = create_group_box("파일 경로")
        path_layout = QVBoxLayout()
        self.replace_file_path, _ = create_file_input("대상 파일:", self, path_layout, is_folder=False)
        self.glossary_file_path, _ = create_file_input("용어집 파일:", self, path_layout, is_folder=False)
        path_group_box.setLayout(path_layout)
        layout.addWidget(path_group_box)

        self.replace_button = create_button("용어 교체 실행", self.logic.run_word_replacement, layout)
        layout.addStretch(1)

    def setup_db_compare_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.tabs.addTab(tab, "DB 비교")

        compare_group_box = create_group_box("비교할 DB 선택")
        compare_layout = QVBoxLayout()

        db1_layout = QHBoxLayout()
        db1_label = QLabel("DB 1:")
        self.db1_selector = QComboBox()
        self.populate_db_selectors([self.db1_selector])
        db1_layout.addWidget(db1_label)
        db1_layout.addWidget(self.db1_selector)
        compare_layout.addLayout(db1_layout)

        db2_layout = QHBoxLayout()
        db2_label = QLabel("DB 2:")
        self.db2_selector = QComboBox()
        self.populate_db_selectors([self.db2_selector])
        db2_layout.addWidget(db2_label)
        db2_layout.addWidget(self.db2_selector)
        compare_layout.addLayout(db2_layout)

        compare_group_box.setLayout(compare_layout)
        layout.addWidget(compare_group_box)

        self.compare_button = create_button("DB 비교 실행", self.logic.run_db_comparison, layout)
        layout.addStretch(1)

    def setup_translation_apply_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.tabs.addTab(tab, "번역문 적용")

        settings_group_box = create_group_box("경로 및 DB 설정")
        settings_layout = QVBoxLayout()
        self.apply_source_dir, _ = create_file_input("번역 원본 폴더:", self, settings_layout, is_folder=True)
        self.apply_target_file, _ = create_file_input("번역 적용 대상 파일:", self, settings_layout, is_folder=False)
        self.apply_db_selector, _ = create_db_selector(self, settings_layout)
        settings_group_box.setLayout(settings_layout)
        layout.addWidget(settings_group_box)

        self.apply_button = create_button("번역문 적용 실행", self.logic.run_translation_apply, layout)
        layout.addStretch(1)

    def populate_db_selectors(self, selectors):
        self.db_manager.populate_db_selectors(selectors)

    def browse_file(self, line_edit):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "All Files (*);;Excel Files (*.xlsx)")
        if file_path:
            line_edit.setText(file_path)

    def browse_folder(self, line_edit):
        folder_path = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if folder_path:
            line_edit.setText(folder_path)

    def browse_path(self, line_edit):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.AnyFile)
        if dialog.exec_():
            path = dialog.selectedFiles()[0]
            line_edit.setText(path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TranslationTool()
    ex.show()
    sys.exit(app.exec_())