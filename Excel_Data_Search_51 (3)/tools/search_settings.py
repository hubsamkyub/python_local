import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


class SearchFilesManager:
    def __init__(self, folder_path, config_file="search_files_config.json"):
        self.folder_path = folder_path
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_file)
        self.search_files = self.load_search_files()
        
    def load_search_files(self):
        """JSON 파일에서 검색 파일 목록을 로드합니다."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 존재하는 파일만 필터링
                    return [file for file in data.get('files', []) 
                           if os.path.exists(os.path.join(self.folder_path, file))]
            except Exception as e:
                print(f"설정 파일 로드 오류: {e}")
                return []
        return []
    
    def save_search_files(self):
        """검색 파일 목록을 JSON 파일에 저장합니다."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({'files': self.search_files}, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"설정 파일 저장 오류: {e}")
            return False
    
    def validate_and_update_files(self, excel_cache):
        """excel_cache를 기반으로 검색 파일 목록을 업데이트합니다."""
        # 존재하지 않는 파일 제거
        valid_files = [f for f in self.search_files 
                      if f in excel_cache or os.path.exists(os.path.join(self.folder_path, f))]
        
        # 변경사항이 있으면 저장
        if set(valid_files) != set(self.search_files):
            self.search_files = valid_files
            self.save_search_files()
        
        return self.search_files
    
    def should_search_file(self, filename):
        """파일을 검색할지 여부를 결정합니다."""
        # 검색 파일 목록이 비어있으면 모든 파일 검색
        if not self.search_files:
            return filename.endswith('.xlsx')
        
        # 그렇지 않으면 목록에 있는 파일만 검색
        return filename in self.search_files

    def get_search_files(self):
        """현재 검색 파일 목록을 반환합니다."""
        return self.search_files
    
    def add_search_file(self, filename):
        """검색 파일 목록에 파일을 추가합니다."""
        if not filename.endswith('.xlsx'):
            return False, "엑셀 파일(.xlsx)만 추가할 수 있습니다."
            
        # 이미 목록에 있는지 확인
        if filename in self.search_files:
            return False, "이미 목록에 있는, 파일입니다."
            
        # 파일이 실제로 존재하는지 확인
        if not os.path.exists(os.path.join(self.folder_path, filename)):
            return False, "폴더에 존재하지 않는 파일입니다."
            
        self.search_files.append(filename)
        self.save_search_files()
        return True, "파일이 추가되었습니다."
    
    def remove_search_file(self, filename):
        """검색 파일 목록에서 파일을 제거합니다."""
        if filename in self.search_files:
            self.search_files.remove(filename)
            self.save_search_files()
            return True
        return False