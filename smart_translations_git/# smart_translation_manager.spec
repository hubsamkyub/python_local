# smart_translation_manager.spec

# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# 현재 작업 디렉토리
WORK_DIR = os.getcwd()

# 데이터 파일들 수집
datas = [
    # 설정 파일들
    ('config.json', '.'),
    ('translation_config.json', '.'),
    ('credentials.json', '.'),  # 있는 경우
    
    # 데이터베이스 파일
    ('smart_translations.db', '.'),
    
    # 템플릿 파일들 (있는 경우)
    ('.env.template', '.') if os.path.exists('.env.template') else None,
    ('credentials.template.json', '.') if os.path.exists('credentials.template.json') else None,
    
    # dialogs 폴더 전체
    ('dialogs', 'dialogs'),
    
    # utils 폴더 전체
    ('utils', 'utils'),
]

# None 값 제거
datas = [d for d in datas if d is not None]

# 숨겨진 imports (외부 라이브러리들)
hiddenimports = [
    # GUI 관련
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.simpledialog',
    'tkinterdnd2',
    
    # 데이터 처리
    'pandas',
    'openpyxl',
    'openpyxl.styles',
    'sqlite3',
    'json',
    
    # API 관련
    'deepl',
    'openai',
    'requests',
    'google.oauth2.service_account',
    'googleapiclient.discovery',
    
    # 기타 라이브러리
    'dotenv',
    'difflib',
    're',
    'uuid',
    'threading',
    'time',
    'datetime',
    'collections',
    'os',
    'sys',
    
    # 커스텀 모듈들
    'config',
    'translation_helpers',
    'scenario_manager',
    'glossary_matcher',
    'text_preprocessor',
    'exclusion_manager',
    'translation_consolidator',
    
    # dialogs 모듈들
    'dialogs.edit_dialogs',
    'dialogs.preview_dialogs',
    'dialogs.selection_dialogs',
    'dialogs.speaker_dialog',
    
    # utils 모듈들
    'utils',
]

# 제외할 모듈들 (불필요한 모듈들)
excludes = [
    'matplotlib',
    'scipy',
    'numpy.testing',
    'pytest',
    'unittest',
    'test',
    'setuptools',
    'pip',
]

# 바이너리 파일들 (필요한 경우)
binaries = []

# PyInstaller Analysis 설정
a = Analysis(
    ['smart_translation_manager.py'],  # 메인 스크립트
    pathex=[WORK_DIR],  # 경로
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# PYZ 아카이브 생성
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# 실행 파일 생성 (원클릭 실행)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SmartTranslationManager',  # 실행 파일 이름
    debug=False,  # 디버그 모드 비활성화
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # UPX 압축 사용 (파일 크기 줄임)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 콘솔 창 숨김 (GUI 앱이므로)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,  # 아이콘 파일이 있으면 사용
)

# COLLECT는 폴더 형태로 배포할 때 사용 (현재는 단일 exe 파일로 설정)
# coll = COLLECT(
#     exe,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     strip=False,
#     upx=True,
#     upx_exclude=[],
#     name='SmartTranslationManager'
# )