# build.spec (오류 수정 및 클린 버전)

# -*- mode: python ; coding: utf-8 -*-

# tkinterdnd2의 라이브러리 파일 경로를 직접 찾아주기 위한 코드
import os
import tkinterdnd2
tkinterdnd2_path = os.path.dirname(tkinterdnd2.__file__)
tkdnd_data_path = os.path.join(tkinterdnd2_path, 'tkdnd')


block_cipher = None

a = Analysis(
    ['smart_translation_manager.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        (tkdnd_data_path, 'tkinterdnd2/tkdnd')
    ],
    hiddenimports=[
        'pandas', 'openpyxl', 'tkinterdnd2',
        'googleapiclient.discovery', 'google_auth_httplib2',
        'google.oauth2.credentials', 'dotenv', 'requests', 'deepl'
    ],
    hookspath=[], # ◀◀◀ numpy.core.hook 관련 내용 제거
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SmartTranslationManager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SmartTranslationApp'
)