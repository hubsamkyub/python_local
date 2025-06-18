# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['Excel_Data_Search_50.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('table_relationships.json', '.'), 
        ('typecode_mapping.json', '.'),
        ('dulcet-antler-462703-n8-d2fbdb362407.json', '.'),
        ('ui', 'ui'),                    # ui 폴더 전체 포함
        ('utils', 'utils'),              # utils 폴더 전체 포함
        ('tools', 'tools'),              # tools 폴더 전체 포함
        ('db_relationships.py', '.'),    # db_relationships 모듈 포함
    ],
    hiddenimports=[
        'psutil',
        # utils 모듈들
        'utils.cache_utils',
        'utils.excel_utils', 
        'utils.config_utils',
        'utils.type_mappings',
        # ui 모듈들 (필요한 것들 추가)
        'ui',
        # tools 모듈들 (필요한 것들 추가)
        'tools',
        'tools.translate',
        # 기타 모듈들
        'db_relationships',
        # 추가로 필요할 수 있는 모듈들
        'tkinter',
        'tkinter.ttk',
        'pandas',
        'sqlite3',
        'openpyxl',
        'xlsxwriter',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Excel_Data_Search_50',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['converted_icon.ico'],
)